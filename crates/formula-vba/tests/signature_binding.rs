#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, verify_vba_digital_signature,
    verify_vba_digital_signature_bound, VbaProjectBindingVerification, VbaSignatureBinding,
    VbaSignatureVerification,
};
use md5::{Digest as _, Md5};

mod signature_test_utils;

use signature_test_utils::{make_pkcs7_detached_signature, make_pkcs7_signed_message};

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_minimal_vba_project_bin(module1_code: &[u8], signature_blob: Option<&[u8]>) -> Vec<u8> {
    // Store the module as a plain compressed container (text_offset = 0).
    let module_container = compress_container(module1_code);

    // Minimal `VBA/dir` listing one module.
    let dir_decompressed = {
        let mut out = Vec::new();

        // PROJECTNAME + PROJECTCONSTANTS are incorporated into ContentNormalizedData when present.
        push_record(&mut out, 0x0004, b"VBAProject");
        push_record(&mut out, 0x000C, b"");

        // MODULENAME
        push_record(&mut out, 0x0019, b"Module1");

        // MODULESTREAMNAME + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULETYPE (standard)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());

        // MODULETEXTOFFSET
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());

        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\nModule=Module1\r\n")
            .expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("VBA storage");

    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }

    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
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

fn der_len(len: usize) -> Vec<u8> {
    if len < 0x80 {
        return vec![len as u8];
    }
    let mut out = Vec::new();
    let mut n = len;
    let mut buf = Vec::new();
    while n > 0 {
        buf.push((n & 0xFF) as u8);
        n >>= 8;
    }
    buf.reverse();
    out.push(0x80 | (buf.len() as u8));
    out.extend_from_slice(&buf);
    out
}

fn der_tlv(tag: u8, value: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(tag);
    out.extend_from_slice(&der_len(value.len()));
    out.extend_from_slice(value);
    out
}

fn der_sequence(content: &[u8]) -> Vec<u8> {
    der_tlv(0x30, content)
}

fn der_null() -> Vec<u8> {
    vec![0x05, 0x00]
}

fn der_oid_raw(oid: &[u8]) -> Vec<u8> {
    der_tlv(0x06, oid)
}

fn der_integer_u32(n: u32) -> Vec<u8> {
    // DER INTEGER encoding, positive.
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
    if bytes[0] & 0x80 != 0 {
        bytes.insert(0, 0);
    }
    der_tlv(0x02, &bytes)
}

fn der_octet_string(bytes: &[u8]) -> Vec<u8> {
    der_tlv(0x04, bytes)
}

fn build_spc_indirect_data_content_sha256(project_digest: &[u8]) -> Vec<u8> {
    // SHA-256 OID: 2.16.840.1.101.3.4.2.1
    let sha256_oid = [0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01];

    let mut alg_id = Vec::new();
    alg_id.extend_from_slice(&der_oid_raw(&sha256_oid));
    alg_id.extend_from_slice(&der_null());
    let alg_id = der_sequence(&alg_id);

    let mut digest_info = Vec::new();
    digest_info.extend_from_slice(&alg_id);
    digest_info.extend_from_slice(&der_octet_string(project_digest));
    let digest_info = der_sequence(&digest_info);

    let mut spc = Vec::new();
    // `data` (ignored by our parser) – use NULL.
    spc.extend_from_slice(&der_null());
    spc.extend_from_slice(&digest_info);
    der_sequence(&spc)
}

fn build_spc_indirect_data_content_md5(project_digest: &[u8]) -> Vec<u8> {
    // MD5 OID: 1.2.840.113549.2.5
    let md5_oid = [0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x02, 0x05];

    let mut alg_id = Vec::new();
    alg_id.extend_from_slice(&der_oid_raw(&md5_oid));
    alg_id.extend_from_slice(&der_null());
    let alg_id = der_sequence(&alg_id);

    let mut digest_info = Vec::new();
    digest_info.extend_from_slice(&alg_id);
    digest_info.extend_from_slice(&der_octet_string(project_digest));
    let digest_info = der_sequence(&digest_info);

    let mut spc = Vec::new();
    // `data` (ignored by our parser) – use NULL.
    spc.extend_from_slice(&der_null());
    spc.extend_from_slice(&digest_info);
    der_sequence(&spc)
}

fn build_spc_indirect_data_content_v2_sha256(project_digest_md5: &[u8]) -> Vec<u8> {
    // MS-OSHARED §2.3.2.4.3.2: SpcIndirectDataContentV2
    //
    // The VBA project hash bytes are stored in SigDataV1Serialized.sourceHash (MD5 per MS-OSHARED
    // §4.3).
    let spc_indirect_data_v2_oid = [0x2B, 0x06, 0x01, 0x04, 0x01, 0x82, 0x37, 0x02, 0x01, 0x1F];
    let sha256_oid = [0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01];

    let sig_format_descriptor_v1 = der_sequence(&{
        let mut v = Vec::new();
        v.extend_from_slice(&der_integer_u32(0)); // size (ignored by our parser)
        v.extend_from_slice(&der_integer_u32(1)); // version
        v.extend_from_slice(&der_integer_u32(1)); // format
        v
    });

    let mut spc_attr = Vec::new();
    spc_attr.extend_from_slice(&der_oid_raw(&spc_indirect_data_v2_oid));
    // value [0] EXPLICIT OCTET STRING containing DER SigFormatDescriptorV1.
    spc_attr.extend_from_slice(&der_tlv(0xA0, &der_octet_string(&sig_format_descriptor_v1)));
    let spc_attr = der_sequence(&spc_attr);

    // SigDataV1Serialized ::= SEQUENCE { 6 INTEGERs, algorithmId OID, compiledHash OCTET STRING, sourceHash OCTET STRING }
    let sig_data_v1 = der_sequence(&{
        let mut v = Vec::new();
        v.extend_from_slice(&der_integer_u32(0)); // algorithmIdSize
        v.extend_from_slice(&der_integer_u32(0)); // compiledHashSize
        v.extend_from_slice(&der_integer_u32(project_digest_md5.len() as u32)); // sourceHashSize
        v.extend_from_slice(&der_integer_u32(0)); // algorithmIdOffset
        v.extend_from_slice(&der_integer_u32(0)); // compiledHashOffset
        v.extend_from_slice(&der_integer_u32(0)); // sourceHashOffset
        v.extend_from_slice(&der_oid_raw(&sha256_oid)); // algorithmId
        v.extend_from_slice(&der_octet_string(&[])); // compiledHash (empty)
        v.extend_from_slice(&der_octet_string(project_digest_md5)); // sourceHash (MD5 bytes)
        v
    });

    let alg_id = der_sequence(&{
        let mut v = Vec::new();
        v.extend_from_slice(&der_oid_raw(&sha256_oid));
        v.extend_from_slice(&der_null());
        v
    });

    let digest_info = der_sequence(&{
        let mut v = Vec::new();
        v.extend_from_slice(&alg_id);
        v.extend_from_slice(&der_octet_string(&sig_data_v1));
        v
    });

    let mut spc = Vec::new();
    spc.extend_from_slice(&spc_attr);
    spc.extend_from_slice(&digest_info);
    der_sequence(&spc)
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
fn bound_signature_sets_binding_bound() {
    // Include an Attribute line and LF-only newlines to exercise normalization.
    let module1 = b"Attribute VB_Name = \"Module1\"\nSub A()\nEnd Sub\n";
    let unsigned = build_minimal_vba_project_bin(module1, None);
    let normalized = content_normalized_data(&unsigned).expect("content normalized data");
    let digest: [u8; 16] = Md5::digest(&normalized).into();

    // MS-OSHARED: the DigestInfo.algorithm may be SHA-256 even when the VBA project digest bytes
    // are a 16-byte MD5.
    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_detached_signature(&signed_content);

    let mut signature_stream = signed_content.clone();
    signature_stream.extend_from_slice(&pkcs7);

    let signed = build_minimal_vba_project_bin(module1, Some(&signature_stream));
    let sig = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);

    let bound = verify_vba_digital_signature_bound(&signed)
        .expect("bound verify")
        .expect("signature present");
    assert_eq!(
        bound.signature.verification,
        VbaSignatureVerification::SignedVerified
    );
    assert!(matches!(
        bound.binding,
        VbaProjectBindingVerification::BoundVerified(_)
    ));
}

#[test]
fn tampering_project_changes_binding_but_not_pkcs7_verification() {
    let module1 = b"Sub A()\r\nEnd Sub\r\n";
    let unsigned = build_minimal_vba_project_bin(module1, None);
    let normalized = content_normalized_data(&unsigned).expect("content normalized data");
    let digest: [u8; 16] = Md5::digest(&normalized).into();

    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_detached_signature(&signed_content);
    let mut signature_stream = signed_content.clone();
    signature_stream.extend_from_slice(&pkcs7);

    // Tamper with a project stream (module bytes) but keep the signature blob intact.
    let mut tampered_module = module1.to_vec();
    tampered_module[0] ^= 0xFF;

    let tampered = build_minimal_vba_project_bin(&tampered_module, Some(&signature_stream));
    let sig = verify_vba_digital_signature(&tampered)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::NotBound);

    let bound = verify_vba_digital_signature_bound(&tampered)
        .expect("bound verify")
        .expect("signature present");
    assert_eq!(
        bound.signature.verification,
        VbaSignatureVerification::SignedVerified
    );
    assert!(matches!(
        bound.binding,
        VbaProjectBindingVerification::BoundMismatch(_)
    ));
}

#[test]
fn embedded_pkcs7_content_is_used_for_binding() {
    let module1 = b"Sub A()\r\nEnd Sub\r\n";
    let unsigned = build_minimal_vba_project_bin(module1, None);
    let normalized = content_normalized_data(&unsigned).expect("content normalized data");
    let digest: [u8; 16] = Md5::digest(&normalized).into();

    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&signed_content);

    let signed = build_minimal_vba_project_bin(module1, Some(&pkcs7));
    let sig = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
}

#[test]
fn md5_binding_is_supported() {
    let module1 = b"Sub Hello()\r\nEnd Sub\r\n";
    let unsigned = build_minimal_vba_project_bin(module1, None);
    let normalized = content_normalized_data(&unsigned).expect("content normalized data");
    let digest: [u8; 16] = Md5::digest(&normalized).into();
    assert_eq!(digest.len(), 16, "MD5 digest should be 16 bytes");

    let signed_content = build_spc_indirect_data_content_md5(&digest);
    let pkcs7 = make_pkcs7_detached_signature(&signed_content);
    let mut signature_stream = signed_content.clone();
    signature_stream.extend_from_slice(&pkcs7);

    // Baseline: signed project should be bound.
    let signed = build_minimal_vba_project_bin(module1, Some(&signature_stream));
    let sig = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);

    let bound = verify_vba_digital_signature_bound(&signed)
        .expect("bound verify")
        .expect("signature present");
    match bound.binding {
        VbaProjectBindingVerification::BoundVerified(info) => {
            assert_eq!(info.hash_algorithm_name.as_deref(), Some("MD5"));
        }
        other => panic!("expected BoundVerified, got {other:?}"),
    }

    // Tamper with a project stream (module bytes) but keep the signature blob intact.
    let mut tampered_module = module1.to_vec();
    tampered_module[0] ^= 0xFF;

    let tampered = build_minimal_vba_project_bin(&tampered_module, Some(&signature_stream));
    let sig = verify_vba_digital_signature(&tampered)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::NotBound);
}

#[test]
fn v2_signed_digest_source_hash_is_used_for_binding() {
    let module1 = b"Sub Hello()\r\nEnd Sub\r\n";
    let unsigned = build_minimal_vba_project_bin(module1, None);
    let normalized = content_normalized_data(&unsigned).expect("content normalized data");
    let digest: [u8; 16] = Md5::digest(&normalized).into();

    let signed_content = build_spc_indirect_data_content_v2_sha256(&digest);
    let pkcs7 = make_pkcs7_signed_message(&signed_content);
    let signature_stream = wrap_in_digsig_blob(&pkcs7);

    let signed = build_minimal_vba_project_bin(module1, Some(&signature_stream));
    let sig = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);

    let bound = verify_vba_digital_signature_bound(&signed)
        .expect("bound verify")
        .expect("signature present");
    match bound.binding {
        VbaProjectBindingVerification::BoundVerified(info) => {
            assert_eq!(info.signed_digest.as_deref(), Some(digest.as_ref()));
            assert_eq!(info.computed_digest.as_deref(), Some(digest.as_ref()));
        }
        other => panic!("expected BoundVerified, got {other:?}"),
    }
}
