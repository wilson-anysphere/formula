#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, extract_vba_signature_signed_digest,
    forms_normalized_data, verify_vba_digital_signature, verify_vba_project_signature_binding,
    verify_vba_signature_binding, VbaProjectBindingVerification, VbaSignatureBinding,
    VbaSignatureVerification,
};
use md5::{Digest as _, Md5};

mod signature_test_utils;

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn md5_digest(bytes: &[u8]) -> Vec<u8> {
    Md5::digest(bytes).to_vec()
}

fn md5_content_plus_forms(content: &[u8], forms: &[u8]) -> Vec<u8> {
    let mut hasher = Md5::new();
    hasher.update(content);
    hasher.update(forms);
    hasher.finalize().to_vec()
}

fn build_project_with_designer(
    signature_stream_path: &str,
    signature_blob: Option<&[u8]>,
) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // `forms_normalized_data` identifies designer storages via `BaseClass=` lines in the `PROJECT`
    // stream. Keep this minimal but spec-ish.
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

    // Minimal modules (one regular module and one designer module).
    let module1_source = b"Sub Hello()\r\nEnd Sub\r\n";
    let module1_container = compress_container(module1_source);

    let userform_source = b"Sub FormInit()\r\nEnd Sub\r\n";
    let userform_container = compress_container(userform_source);

    // Minimal `dir` stream (decompressed form) with two modules.
    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTNAME
        push_record(&mut out, 0x0004, b"VBAProject");
        // ---- Module1 ----
        push_record(&mut out, 0x0019, b"Module1");
        // MODULESTREAMNAME + reserved u16
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // MODULETYPE (standard)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        // MODULETEXTOFFSET
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());

        // ---- UserForm1 ----
        push_record(&mut out, 0x0019, b"UserForm1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // MODULETYPE (userform)
        push_record(&mut out, 0x0021, &3u16.to_le_bytes());
        // MODULETEXTOFFSET
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());
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

    // Root-level designer storage with non-empty stream data so FormsNormalizedData is non-empty.
    ole.create_storage("UserForm1")
        .expect("create designer storage");
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("create designer stream");
        s.write_all(b"FORMDATA").expect("write designer stream");
    }

    if let Some(sig) = signature_blob {
        let mut s = ole
            .create_stream(signature_stream_path)
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
fn signature_binding_accepts_content_only_or_content_plus_forms_transcripts() {
    // Build the unsigned project first so we can compute both candidate digests.
    let unsigned = build_project_with_designer("\u{0005}DigitalSignature", None);
    let content = content_normalized_data(&unsigned).expect("ContentNormalizedData");
    let forms = forms_normalized_data(&unsigned).expect("FormsNormalizedData");
    assert!(
        !forms.is_empty(),
        "expected FormsNormalizedData to be non-empty for a project with a designer storage"
    );

    let digest_content_only = md5_digest(&content);
    let digest_content_plus_forms = md5_content_plus_forms(&content, &forms);
    assert_ne!(
        digest_content_only, digest_content_plus_forms,
        "expected digest variants to differ when FormsNormalizedData is non-empty"
    );

    // ---- Case 1: content-only digest ----
    let sig_stream = make_signature_stream_detached(&digest_content_only);
    let signed = build_project_with_designer("\u{0005}DigitalSignature", Some(&sig_stream));
    let sig = verify_vba_digital_signature(&signed)
        .expect("signature inspection should succeed")
        .expect("signature should be present");
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
    assert_eq!(
        verify_vba_signature_binding(&signed, &sig_stream),
        VbaSignatureBinding::Bound
    );
    match verify_vba_project_signature_binding(&signed, &sig_stream).expect("binding verification") {
        VbaProjectBindingVerification::BoundVerified(info) => {
            assert_eq!(info.signed_digest.as_deref(), Some(digest_content_only.as_slice()));
            assert_eq!(
                info.computed_digest.as_deref(),
                Some(digest_content_only.as_slice())
            );
        }
        other => panic!("expected BoundVerified, got {other:?}"),
    }

    // ---- Case 2: content + forms digest ----
    let sig_stream = make_signature_stream_detached(&digest_content_plus_forms);
    let signed = build_project_with_designer("\u{0005}DigitalSignatureEx", Some(&sig_stream));
    let sig = verify_vba_digital_signature(&signed)
        .expect("signature inspection should succeed")
        .expect("signature should be present");
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    let signed_digest = extract_vba_signature_signed_digest(&sig.signature)
        .expect("signed digest parse")
        .expect("signed digest present");
    assert_eq!(signed_digest.digest, digest_content_plus_forms);

    let content2 = content_normalized_data(&signed).expect("ContentNormalizedData");
    let forms2 = forms_normalized_data(&signed).expect("FormsNormalizedData");
    assert_eq!(content, content2);
    assert_eq!(forms, forms2);
    let computed_digest = md5_content_plus_forms(&content2, &forms2);
    assert_eq!(computed_digest, digest_content_plus_forms);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
    assert_eq!(
        verify_vba_signature_binding(&signed, &sig_stream),
        VbaSignatureBinding::Bound
    );
    match verify_vba_project_signature_binding(&signed, &sig_stream).expect("binding verification") {
        VbaProjectBindingVerification::BoundVerified(info) => {
            assert_eq!(
                info.signed_digest.as_deref(),
                Some(digest_content_plus_forms.as_slice())
            );
            assert_eq!(
                info.computed_digest.as_deref(),
                Some(digest_content_plus_forms.as_slice())
            );
        }
        other => panic!("expected BoundVerified, got {other:?}"),
    }

    // ---- Case 3: wrong digest ----
    let mut wrong_digest = digest_content_plus_forms.clone();
    wrong_digest[0] ^= 0xFF;
    let sig_stream = make_signature_stream_detached(&wrong_digest);
    let signed = build_project_with_designer("\u{0005}DigitalSignatureEx", Some(&sig_stream));
    let sig = verify_vba_digital_signature(&signed)
        .expect("signature inspection should succeed")
        .expect("signature should be present");
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::NotBound);
    assert_eq!(
        verify_vba_signature_binding(&signed, &sig_stream),
        VbaSignatureBinding::Unknown
    );
    match verify_vba_project_signature_binding(&signed, &sig_stream).expect("binding verification") {
        VbaProjectBindingVerification::BoundUnknown(info) => {
            assert_eq!(info.signed_digest.as_deref(), Some(wrong_digest.as_slice()));
            assert!(
                matches!(info.computed_digest.as_deref(), Some(d) if d == digest_content_only.as_slice()),
                "expected computed_digest to be the legacy ContentHash digest"
            );
        }
        other => panic!("expected BoundUnknown, got {other:?}"),
    }
}
