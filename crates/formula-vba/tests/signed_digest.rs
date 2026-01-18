#![cfg(not(target_arch = "wasm32"))]

use formula_vba::extract_vba_signature_signed_digest;

mod signature_test_utils;

use signature_test_utils::{make_pkcs7_detached_signature, make_pkcs7_signed_message};

fn wrap_in_digsig_info_serialized(pkcs7: &[u8]) -> Vec<u8> {
    // Synthetic DigSigInfoSerialized-like blob:
    // [cbSignature, cbSigningCertStore, cchProjectName] (LE u32)
    // [projectName UTF-16LE] [certStore bytes] [signature bytes]
    let project_name_utf16: Vec<u16> = "VBAProject\0".encode_utf16().collect();
    let mut project_name_bytes = Vec::new();
    for ch in &project_name_utf16 {
        project_name_bytes.extend_from_slice(&ch.to_le_bytes());
    }

    let cert_store = vec![0xAA, 0xBB, 0xCC, 0xDD];

    let cb_signature = pkcs7.len() as u32;
    let cb_cert_store = cert_store.len() as u32;
    let cch_project = project_name_utf16.len() as u32;

    let mut out = Vec::new();
    out.extend_from_slice(&cb_signature.to_le_bytes());
    out.extend_from_slice(&cb_cert_store.to_le_bytes());
    out.extend_from_slice(&cch_project.to_le_bytes());
    out.extend_from_slice(&project_name_bytes);
    out.extend_from_slice(&cert_store);
    out.extend_from_slice(pkcs7);
    out
}

fn der_len(len: usize) -> Vec<u8> {
    if len < 0x80 {
        return vec![len as u8];
    }
    let mut bytes = Vec::new();
    let mut tmp = len;
    while tmp > 0 {
        bytes.push((tmp & 0xFF) as u8);
        tmp >>= 8;
    }
    bytes.reverse();
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(1usize.saturating_add(bytes.len()));
    out.push(0x80 | (bytes.len() as u8));
    out.extend_from_slice(&bytes);
    out
}

fn der_tlv(tag: u8, value: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(tag);
    out.extend_from_slice(&der_len(value.len()));
    out.extend_from_slice(value);
    out
}

fn der_sequence(children: &[Vec<u8>]) -> Vec<u8> {
    let mut value = Vec::new();
    for child in children {
        value.extend_from_slice(child);
    }
    der_tlv(0x30, &value)
}

fn der_octet_string(bytes: &[u8]) -> Vec<u8> {
    der_tlv(0x04, bytes)
}

fn der_null() -> Vec<u8> {
    vec![0x05, 0x00]
}

fn der_oid(oid: &str) -> Vec<u8> {
    let arcs: Vec<u32> = oid
        .split('.')
        .map(|s| s.parse::<u32>().expect("numeric arc"))
        .collect();
    assert!(arcs.len() >= 2, "OID needs at least 2 arcs");
    let mut out = Vec::new();
    out.push((arcs[0] * 40 + arcs[1]) as u8);
    for &arc in &arcs[2..] {
        let mut tmp = arc;
        let mut buf = Vec::new();
        buf.push((tmp & 0x7F) as u8);
        tmp >>= 7;
        while tmp > 0 {
            buf.push(((tmp & 0x7F) as u8) | 0x80);
            tmp >>= 7;
        }
        buf.reverse();
        out.extend_from_slice(&buf);
    }
    der_tlv(0x06, &out)
}

fn der_integer_u32(n: u32) -> Vec<u8> {
    // DER INTEGER encoding (positive).
    let mut bytes = Vec::new();
    let mut v = n;
    while v > 0 {
        bytes.push((v & 0xFF) as u8);
        v >>= 8;
    }
    if bytes.is_empty() {
        bytes.push(0);
    }
    bytes.reverse();
    // Ensure the integer is interpreted as positive.
    if bytes[0] & 0x80 != 0 {
        bytes.insert(0, 0);
    }
    der_tlv(0x02, &bytes)
}

fn make_spc_indirect_data_content_sha256(digest: &[u8]) -> Vec<u8> {
    // data SpcAttributeTypeAndOptionalValue ::= SEQUENCE { type OBJECT IDENTIFIER, value [0] EXPLICIT ANY OPTIONAL }
    let data = der_sequence(&[der_oid("1.3.6.1.4.1.311.2.1.15")]);

    // messageDigest DigestInfo ::= SEQUENCE { digestAlgorithm AlgorithmIdentifier, digest OCTET STRING }
    let alg = der_sequence(&[der_oid("2.16.840.1.101.3.4.2.1"), der_null()]);
    let digest_info = der_sequence(&[alg, der_octet_string(digest)]);

    der_sequence(&[data, digest_info])
}

fn make_spc_indirect_data_content_v2_sha256(source_hash: &[u8]) -> Vec<u8> {
    // MS-OSHARED §2.3.2.4.3.2
    // - data.type = 1.3.6.1.4.1.311.2.1.31
    // - data.value = OCTET STRING containing DER SigFormatDescriptorV1
    // - messageDigest.digest = OCTET STRING containing DER SigDataV1Serialized
    let sig_format_descriptor_v1 = der_sequence(&[
        der_integer_u32(0), // size (ignored by our parser)
        der_integer_u32(1), // version
        der_integer_u32(1), // format
    ]);

    let data = der_sequence(&[
        der_oid("1.3.6.1.4.1.311.2.1.31"),
        // value [0] EXPLICIT OCTET STRING containing SigFormatDescriptorV1 DER.
        der_tlv(0xA0, &der_octet_string(&sig_format_descriptor_v1)),
    ]);

    // SigDataV1Serialized ::= SEQUENCE { 6 INTEGERs, algorithmId OID, compiledHash OCTET STRING, sourceHash OCTET STRING }
    let sig_data = der_sequence(&[
        der_integer_u32(0), // algorithmIdSize
        der_integer_u32(0), // compiledHashSize
        der_integer_u32(source_hash.len() as u32), // sourceHashSize
        der_integer_u32(0), // algorithmIdOffset
        der_integer_u32(0), // compiledHashOffset
        der_integer_u32(0), // sourceHashOffset
        der_oid("2.16.840.1.101.3.4.2.1"), // algorithmId (sha256)
        der_octet_string(&[]),            // compiledHash (empty)
        der_octet_string(source_hash),    // sourceHash (VBA project digest, MD5 bytes)
    ]);

    let alg = der_sequence(&[der_oid("2.16.840.1.101.3.4.2.1"), der_null()]);
    let digest_info = der_sequence(&[alg, der_octet_string(&sig_data)]);

    der_sequence(&[data, digest_info])
}

fn wrap_in_digsig_blob(pkcs7: &[u8]) -> Vec<u8> {
    // Minimal DigSigBlob (MS-OSHARED §2.3.2.2) containing a DigSigInfoSerialized (MS-OSHARED
    // §2.3.2.1) header and the pbSignatureBuffer (PKCS#7 SignedData bytes).
    let signature_offset = 8 + 36;
    let mut blob = Vec::new();
    blob.extend_from_slice(&0u32.to_le_bytes()); // cb placeholder
    blob.extend_from_slice(&8u32.to_le_bytes()); // serializedPointer

    // DigSigInfoSerialized fixed header (9 u32s).
    blob.extend_from_slice(&(pkcs7.len() as u32).to_le_bytes()); // cbSignature
    blob.extend_from_slice(&(signature_offset as u32).to_le_bytes()); // signatureOffset
    blob.extend_from_slice(&0u32.to_le_bytes()); // cbSigningCertStore
    blob.extend_from_slice(&0u32.to_le_bytes()); // certStoreOffset
    blob.extend_from_slice(&0u32.to_le_bytes()); // cbProjectName
    blob.extend_from_slice(&0u32.to_le_bytes()); // projectNameOffset
    blob.extend_from_slice(&0u32.to_le_bytes()); // fTimestamp
    blob.extend_from_slice(&0u32.to_le_bytes()); // cbTimestampUrl
    blob.extend_from_slice(&0u32.to_le_bytes()); // timestampUrlOffset

    blob.extend_from_slice(pkcs7);

    // Pad signatureInfo to 4-byte alignment.
    while (blob.len() - 8) % 4 != 0 {
        blob.push(0);
    }

    let cb = (blob.len() - 8) as u32;
    blob[0..4].copy_from_slice(&cb.to_le_bytes());
    blob
}

#[test]
fn extracts_source_hash_from_spc_indirect_data_content_v2() {
    // Simulate the MS-OSHARED SpcIndirectDataContentV2 variant, where DigestInfo.digest contains
    // a DER-encoded `SigDataV1Serialized` structure instead of raw hash bytes.
    let source_hash = (10u8..26u8).collect::<Vec<_>>();
    assert_eq!(source_hash.len(), 16);

    let sigdata = der_sequence(&[der_octet_string(&source_hash)]);
    let spc = make_spc_indirect_data_content_sha256(&sigdata);
    let pkcs7 = make_pkcs7_signed_message(&spc);

    let got = extract_vba_signature_signed_digest(&pkcs7)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1");
    assert_eq!(got.digest, source_hash);
}

#[test]
fn extracts_signed_digest_from_embedded_pkcs7() {
    let digest = (0u8..32).collect::<Vec<_>>();
    let spc = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);

    let got = extract_vba_signature_signed_digest(&pkcs7)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1");
    assert_eq!(got.digest, digest);
}

#[test]
fn extracts_signed_digest_from_detached_pkcs7_using_prefix_content() {
    let digest = (42u8..74).collect::<Vec<_>>();
    let spc = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_detached_signature(&spc);

    let mut stream = spc.clone();
    stream.extend_from_slice(&pkcs7);

    let got = extract_vba_signature_signed_digest(&stream)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1");
    assert_eq!(got.digest, digest);
}

#[test]
fn extracts_signed_digest_when_pkcs7_is_prefixed_by_header_bytes() {
    let digest = b"this-is-a-test-digest-32-bytes!!".to_vec();
    assert_eq!(digest.len(), 32);
    let spc = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);

    let mut stream = b"VBA\0SIG\0HDR".to_vec();
    stream.extend_from_slice(&pkcs7);

    let got = extract_vba_signature_signed_digest(&stream)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1");
    assert_eq!(got.digest, digest);
}

#[test]
fn extracts_signed_digest_when_pkcs7_is_wrapped_in_digsig_info_serialized() {
    let digest = (100u8..132).collect::<Vec<_>>();
    assert_eq!(digest.len(), 32);
    let spc = make_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&spc);

    let stream = wrap_in_digsig_info_serialized(&pkcs7);

    let got = extract_vba_signature_signed_digest(&stream)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1");
    assert_eq!(got.digest, digest);
}

#[test]
fn extracts_signed_digest_from_digsig_blob_preferring_wrapper_over_trailing_pkcs7() {
    // Digest that should be extracted from the DigSigBlob's pbSignatureBuffer (V2 format).
    let source_hash = (0u8..16).collect::<Vec<_>>();
    let spc_v2 = make_spc_indirect_data_content_v2_sha256(&source_hash);
    let pkcs7_v2 = make_pkcs7_signed_message(&spc_v2);
    let mut stream = wrap_in_digsig_blob(&pkcs7_v2);

    // Append a second PKCS#7 blob containing a *different* digest. A naive scanner that just picks
    // the last SignedData in the stream would return this digest instead of the DigSigBlob one.
    let trailing_digest = vec![0xAAu8; 16];
    let spc_trailing = make_spc_indirect_data_content_sha256(&trailing_digest);
    let pkcs7_trailing = make_pkcs7_signed_message(&spc_trailing);
    stream.extend_from_slice(&pkcs7_trailing);

    let got = extract_vba_signature_signed_digest(&stream)
        .expect("extract digest")
        .expect("digest present");

    assert_eq!(got.digest_algorithm_oid, "2.16.840.1.101.3.4.2.1");
    assert_eq!(got.digest, source_hash);
}
