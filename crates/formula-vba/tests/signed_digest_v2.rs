#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, compute_vba_project_digest, extract_vba_signature_signed_digest,
    verify_vba_digital_signature, DigestAlg, OleFile, VbaSignatureBinding, VbaSignatureVerification,
};

mod signature_test_utils;

use signature_test_utils::{make_pkcs7_detached_signature, make_pkcs7_signed_message};

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

fn der_integer_u32(n: u32) -> Vec<u8> {
    // Minimal DER INTEGER encoding for a non-negative u32.
    if n == 0 {
        return vec![0x02, 0x01, 0x00];
    }
    let mut bytes = Vec::new();
    let mut tmp = n;
    while tmp > 0 {
        bytes.push((tmp & 0xFF) as u8);
        tmp >>= 8;
    }
    bytes.reverse();
    // Ensure the integer is interpreted as positive by prefixing 0x00 when the high bit is set.
    if bytes.first().is_some_and(|b| b & 0x80 != 0) {
        bytes.insert(0, 0x00);
    }
    der_tlv(0x02, &bytes)
}

fn der_null() -> Vec<u8> {
    vec![0x05, 0x00]
}

fn der_oid(oid_content: &[u8]) -> Vec<u8> {
    der_tlv(0x06, oid_content)
}

fn build_sig_data_v1_serialized_with_source_hash(source_hash: &[u8]) -> Vec<u8> {
    // Minimal binary SigDataV1Serialized-ish blob:
    // [version u32 LE] [cbSourceHash u32 LE] [sourceHash bytes]
    let mut out = Vec::new();
    out.extend_from_slice(&1u32.to_le_bytes());
    out.extend_from_slice(&(source_hash.len() as u32).to_le_bytes());
    out.extend_from_slice(source_hash);
    out
}

fn build_sig_data_v1_serialized_asn1(source_hash: &[u8]) -> Vec<u8> {
    // Minimal ASN.1-ish SigDataV1Serialized payload. Real-world `SigDataV1Serialized` is a
    // serialized structure, but some producers may embed it as ASN.1.
    der_sequence(&[der_integer_u32(1), der_octet_string(source_hash)])
}

fn build_spc_indirect_data_content_v2(source_hash: &[u8]) -> Vec<u8> {
    // Minimal SpcIndirectDataContentV2-like payload:
    // SEQUENCE { data ANY, sigData OCTET STRING }
    //
    // The classic Authenticode `SpcIndirectDataContent` stores the digest in a `DigestInfo`
    // (SEQUENCE) as the second element. By making the second element an OCTET STRING instead, we
    // ensure the classic parser fails and our V2 parser path is exercised.
    let data = der_null();
    let sig_data = build_sig_data_v1_serialized_with_source_hash(source_hash);
    der_sequence(&[data, der_octet_string(&sig_data)])
}

fn build_spc_indirect_data_content_v2_with_sigdata_element(sigdata_element: Vec<u8>) -> Vec<u8> {
    let data = der_null();
    der_sequence(&[data, sigdata_element])
}

fn wrap_in_digsig_info_serialized(pkcs7: &[u8]) -> Vec<u8> {
    // Synthetic DigSigInfoSerialized-like blob:
    // [cbSignature, cbSigningCertStore, cchProjectName] (LE u32)
    // [projectName UTF-16LE] [certStore bytes] [signature bytes]
    let project_name_utf16: Vec<u16> = "VBAProject\0".encode_utf16().collect();
    let mut project_name_bytes = Vec::new();
    for ch in &project_name_utf16 {
        project_name_bytes.extend_from_slice(&ch.to_le_bytes());
    }

    // Include a decoy 0x30 prefix to ensure we don't accidentally scan the cert store as the PKCS#7
    // payload.
    let cert_store = vec![0x30, 0xAA, 0xBB, 0xCC, 0xDD];

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

#[test]
fn extracts_signed_digest_from_spc_indirect_data_content_v2_source_hash() {
    let source_hash = (0u8..16).collect::<Vec<_>>();
    let spc_v2 = build_spc_indirect_data_content_v2(&source_hash);
    let pkcs7 = make_pkcs7_signed_message(&spc_v2);

    // Store PKCS#7 in a minimal OLE file signature stream.
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignature")
            .expect("create signature stream");
        s.write_all(&pkcs7).expect("write signature");
    }
    let ole_bytes = ole.into_inner().into_inner();

    // Read it back from the OLE file and extract the digest.
    let mut ole = OleFile::open(&ole_bytes).expect("open ole");
    let signature_stream = ole
        .read_stream_opt("\u{0005}DigitalSignature")
        .expect("read signature stream")
        .expect("signature stream present");

    let got = extract_vba_signature_signed_digest(&signature_stream)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "1.2.840.113549.2.5");
    assert_eq!(got.digest, source_hash);
}

#[test]
fn extracts_signed_digest_from_spc_indirect_data_content_v2_sigdata_as_asn1_sequence() {
    let source_hash = (0u8..16).collect::<Vec<_>>();
    let sigdata = build_sig_data_v1_serialized_asn1(&source_hash);
    let spc_v2 = build_spc_indirect_data_content_v2_with_sigdata_element(sigdata);

    let pkcs7 = make_pkcs7_signed_message(&spc_v2);
    let got = extract_vba_signature_signed_digest(&pkcs7)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "1.2.840.113549.2.5");
    assert_eq!(got.digest, source_hash);
}

#[test]
fn extracts_signed_digest_from_spc_indirect_data_content_v2_sigdata_as_octet_wrapped_asn1() {
    let source_hash = (0u8..16).collect::<Vec<_>>();
    let sigdata_asn1 = build_sig_data_v1_serialized_asn1(&source_hash);
    let spc_v2 =
        build_spc_indirect_data_content_v2_with_sigdata_element(der_octet_string(&sigdata_asn1));

    let pkcs7 = make_pkcs7_signed_message(&spc_v2);
    let got = extract_vba_signature_signed_digest(&pkcs7)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "1.2.840.113549.2.5");
    assert_eq!(got.digest, source_hash);
}

#[test]
fn extracts_signed_digest_from_spc_indirect_data_content_v2_sigdata_as_algorithm_id_first_asn1() {
    // Some producers encode SigDataV1Serialized as an ASN.1 SEQUENCE beginning with an
    // AlgorithmIdentifier instead of a version INTEGER. Ensure we can still locate `sourceHash`.
    let source_hash = (0u8..16).collect::<Vec<_>>();

    // AlgorithmIdentifier ::= SEQUENCE { algorithm OBJECT IDENTIFIER, parameters NULL }
    let sha256_oid = der_oid(&[0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01]);
    let alg_id = der_sequence(&[sha256_oid, der_null()]);
    let sigdata = der_sequence(&[alg_id, der_octet_string(&source_hash)]);

    // Wrap SigData in an OCTET STRING to ensure we exercise the V2 parser path.
    let spc_v2 = build_spc_indirect_data_content_v2_with_sigdata_element(der_octet_string(&sigdata));
    let pkcs7 = make_pkcs7_signed_message(&spc_v2);

    let got = extract_vba_signature_signed_digest(&pkcs7)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "1.2.840.113549.2.5");
    assert_eq!(got.digest, source_hash);
}

#[test]
fn v2_parser_rejects_plain_16_byte_octet_string_instead_of_sigdata() {
    // Regression test: our V2 parser used to have an overly-permissive fallback that would accept
    // *any* 16-byte OCTET STRING as the VBA project hash. This is unsafe for binding verification.
    //
    // Ensure we only accept properly shaped SigDataV1Serialized structures.
    let plain_16 = (0u8..16).collect::<Vec<_>>();
    let spc_v2 =
        build_spc_indirect_data_content_v2_with_sigdata_element(der_octet_string(&plain_16));
    let pkcs7 = make_pkcs7_signed_message(&spc_v2);

    assert!(
        extract_vba_signature_signed_digest(&pkcs7).is_err(),
        "expected digest extraction to fail for non-SigData payload"
    );
}

#[test]
fn extracts_signed_digest_from_v2_detached_pkcs7_using_prefix_content() {
    let source_hash = (0u8..16).collect::<Vec<_>>();
    let spc_v2 = build_spc_indirect_data_content_v2(&source_hash);
    let pkcs7 = make_pkcs7_detached_signature(&spc_v2);

    let mut stream = spc_v2.clone();
    stream.extend_from_slice(&pkcs7);

    let got = extract_vba_signature_signed_digest(&stream)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "1.2.840.113549.2.5");
    assert_eq!(got.digest, source_hash);
}

#[test]
fn extracts_signed_digest_from_v2_when_wrapped_in_digsig_info_serialized() {
    let source_hash = (0u8..16).collect::<Vec<_>>();
    let spc_v2 = build_spc_indirect_data_content_v2(&source_hash);
    let pkcs7 = make_pkcs7_signed_message(&spc_v2);

    let stream = wrap_in_digsig_info_serialized(&pkcs7);
    let got = extract_vba_signature_signed_digest(&stream)
        .expect("extract digest")
        .expect("digest present");
    assert_eq!(got.digest_algorithm_oid, "1.2.840.113549.2.5");
    assert_eq!(got.digest, source_hash);
}

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_minimal_vba_project_bin(module1: &[u8], signature_blob: Option<&[u8]>) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\nModule=Module1\r\n")
            .expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("VBA storage");

    {
        // Minimal decompressed `VBA/dir` stream that `content_normalized_data` can parse.
        let dir_decompressed = {
            let mut out = Vec::new();
            // PROJECTNAME
            push_record(&mut out, 0x0004, b"VBAProject");
            // MODULENAME
            push_record(&mut out, 0x0019, b"Module1");
            // MODULESTREAMNAME + reserved u16
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(b"Module1");
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name);
            // MODULETYPE (standard)
            push_record(&mut out, 0x0021, &0u16.to_le_bytes());
            // MODULETEXTOFFSET (0)
            push_record(&mut out, 0x0031, &0u32.to_le_bytes());
            out
        };
        let dir_container = compress_container(&dir_decompressed);

        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        let module_container = compress_container(module1);
        s.write_all(&module_container).expect("write module");
    }

    if let Some(sig) = signature_blob {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignature")
            .expect("signature stream");
        s.write_all(sig).expect("write signature");
    }

    ole.into_inner().into_inner()
}

#[test]
fn v2_signed_digest_is_used_for_signature_binding() {
    let module1 = b"module1-bytes";
    let unsigned = build_minimal_vba_project_bin(module1, None);
    let digest = compute_vba_project_digest(&unsigned, DigestAlg::Md5).expect("digest");
    assert_eq!(digest.len(), 16, "VBA project digest must be MD5 (16 bytes)");

    let spc_v2 = build_spc_indirect_data_content_v2(&digest);
    let pkcs7 = make_pkcs7_signed_message(&spc_v2);

    let signed = build_minimal_vba_project_bin(module1, Some(&pkcs7));
    let sig = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
}
