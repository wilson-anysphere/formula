#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, forms_normalized_data, verify_vba_digital_signature,
    verify_vba_digital_signature_bound, verify_vba_project_signature_binding,
    VbaProjectBindingVerification, VbaSignatureBinding, VbaSignatureVerification,
};
use md5::{Digest as _, Md5};

mod signature_test_utils;

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_project_with_missing_designer_storage(signature_blob: Option<&[u8]>) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // PROJECT stream references a designer via BaseClass=, but we intentionally omit the required
    // root-level designer storage to force `FormsNormalizedData` computation to fail.
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\n")
            .expect("write PROJECT");
        s.write_all(b"Module=Module1\r\n")
            .expect("write PROJECT");
        s.write_all(b"BaseClass=UserForm1\r\n")
            .expect("write PROJECT");
    }

    ole.create_storage("VBA").expect("VBA storage");

    // Standard module.
    let module1_source = b"Sub Hello()\r\nEnd Sub\r\n";
    let module1_container = compress_container(module1_source);

    // Designer module (UserForm) referenced by BaseClass=.
    let userform_source = b"Sub FormInit()\r\nEnd Sub\r\n";
    let userform_container = compress_container(userform_source);

    // Minimal `dir` stream (decompressed form) with two modules.
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTNAME (included in ContentNormalizedData when present).
        push_record(&mut out, 0x0004, b"VBAProject");

        // ---- Module1 ----
        push_record(&mut out, 0x0019, b"Module1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // standard module
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // text offset

        // ---- UserForm1 ----
        push_record(&mut out, 0x0019, b"UserForm1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0021, &3u16.to_le_bytes()); // MODULETYPE (UserForm)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // text offset
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module1_container).expect("write module");
    }
    {
        let mut s = ole
            .create_stream("VBA/UserForm1")
            .expect("designer module stream");
        s.write_all(&userform_container)
            .expect("write designer module");
    }

    // Intentionally do NOT create the required root-level `UserForm1/*` streams.

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

fn der_octet_string(bytes: &[u8]) -> Vec<u8> {
    der_tlv(0x04, bytes)
}

fn build_spc_indirect_data_content_sha1(digest: &[u8]) -> Vec<u8> {
    // SHA-1 OID: 1.3.14.3.2.26
    let sha1_oid = [0x2B, 0x0E, 0x03, 0x02, 0x1A];

    let mut alg_id = Vec::new();
    alg_id.extend_from_slice(&der_oid_raw(&sha1_oid));
    alg_id.extend_from_slice(&der_null());
    let alg_id = der_sequence(&alg_id);

    let mut digest_info = Vec::new();
    digest_info.extend_from_slice(&alg_id);
    digest_info.extend_from_slice(&der_octet_string(digest));
    let digest_info = der_sequence(&digest_info);

    let mut spc = Vec::new();
    // `data` (ignored by our parser) – use NULL.
    spc.extend_from_slice(&der_null());
    spc.extend_from_slice(&digest_info);
    der_sequence(&spc)
}

fn make_signature_stream_detached(digest: &[u8]) -> Vec<u8> {
    let signed_content = build_spc_indirect_data_content_sha1(digest);
    let pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&signed_content);

    let mut signature_stream = signed_content;
    signature_stream.extend_from_slice(&pkcs7);
    signature_stream
}

#[test]
fn binding_is_unknown_when_forms_normalized_data_is_unavailable() {
    // FormsNormalizedData should error due to the missing root designer storage.
    let unsigned = build_project_with_missing_designer_storage(None);
    assert!(
        forms_normalized_data(&unsigned).is_err(),
        "expected FormsNormalizedData computation to fail for a missing designer storage"
    );

    // Sign a digest that does NOT match ContentHash, and ensure binding is reported as Unknown
    // (we can't rule out an Agile binding when FormsNormalizedData is unavailable).
    let content = content_normalized_data(&unsigned).expect("ContentNormalizedData");
    let expected_content_hash: [u8; 16] = Md5::digest(&content).into();
    let mut wrong_digest: [u8; 16] = Md5::digest(&content).into();
    wrong_digest[0] ^= 0xFF;

    let sig_stream = make_signature_stream_detached(&wrong_digest);
    let signed = build_project_with_missing_designer_storage(Some(&sig_stream));

    let sig = verify_vba_digital_signature(&signed)
        .expect("signature inspection should succeed")
        .expect("signature should be present");
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Unknown);

    // The richer helper should also treat this as unknown (we can't compute FormsNormalizedData so
    // we can't definitively say "NotBound").
    let bound = verify_vba_digital_signature_bound(&signed)
        .expect("bound verify")
        .expect("signature should be present");
    match bound.binding {
        VbaProjectBindingVerification::BoundUnknown(info) => {
            assert_eq!(info.signed_digest.as_deref(), Some(wrong_digest.as_slice()));
            assert_eq!(
                info.computed_digest.as_deref(),
                Some(expected_content_hash.as_slice())
            );
        }
        other => panic!("expected BoundUnknown, got {other:?}"),
    }

    // And the signature-part binding helper should behave consistently.
    let binding =
        verify_vba_project_signature_binding(&signed, &sig_stream).expect("binding verification");
    match binding {
        VbaProjectBindingVerification::BoundUnknown(info) => {
            assert_eq!(info.signed_digest.as_deref(), Some(wrong_digest.as_slice()));
            assert_eq!(
                info.computed_digest.as_deref(),
                Some(expected_content_hash.as_slice())
            );
        }
        other => panic!("expected BoundUnknown, got {other:?}"),
    }
}

#[test]
fn binding_is_bound_when_content_hash_matches_even_if_forms_normalized_data_is_unavailable() {
    let unsigned = build_project_with_missing_designer_storage(None);
    assert!(
        forms_normalized_data(&unsigned).is_err(),
        "expected FormsNormalizedData computation to fail for a missing designer storage"
    );

    // Compute the legacy Content Hash (MD5(ContentNormalizedData)).
    let content = content_normalized_data(&unsigned).expect("ContentNormalizedData");
    let digest: [u8; 16] = Md5::digest(&content).into();

    let sig_stream = make_signature_stream_detached(&digest);
    let signed = build_project_with_missing_designer_storage(Some(&sig_stream));

    let sig = verify_vba_digital_signature(&signed)
        .expect("signature inspection should succeed")
        .expect("signature should be present");
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);

    let bound = verify_vba_digital_signature_bound(&signed)
        .expect("bound verify")
        .expect("signature should be present");
    match bound.binding {
        VbaProjectBindingVerification::BoundVerified(info) => {
            assert_eq!(info.signed_digest.as_deref(), Some(digest.as_slice()));
            assert_eq!(info.computed_digest.as_deref(), Some(digest.as_slice()));
        }
        other => panic!("expected BoundVerified, got {other:?}"),
    }

    let binding =
        verify_vba_project_signature_binding(&signed, &sig_stream).expect("binding verification");
    match binding {
        VbaProjectBindingVerification::BoundVerified(info) => {
            assert_eq!(info.signed_digest.as_deref(), Some(digest.as_slice()));
            assert_eq!(info.computed_digest.as_deref(), Some(digest.as_slice()));
        }
        other => panic!("expected BoundVerified, got {other:?}"),
    }
}
