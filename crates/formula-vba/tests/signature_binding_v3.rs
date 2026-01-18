#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    agile_content_hash_md5, compress_container, compute_vba_project_digest_v3, content_hash_md5,
    contents_hash_v3, verify_vba_digital_signature, verify_vba_digital_signature_bound,
    verify_vba_project_signature_binding, verify_vba_signature_binding,
    verify_vba_signature_binding_with_stream_path, DigestAlg, VbaProjectBindingVerification,
    VbaSignatureBinding, VbaSignatureStreamKind, VbaSignatureVerification,
};

mod signature_test_utils;

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_minimal_vba_project_bin_v3(
    signature_blob: Option<&[u8]>,
    designer_payload: &[u8],
) -> Vec<u8> {
    let module_source = b"Sub Hello()\r\nEnd Sub";
    let module_container = compress_container(module_source);
    let userform_source = b"Sub FormHello()\r\nEnd Sub";
    let userform_container = compress_container(userform_source);

    // Minimal `dir` stream (decompressed form) with:
    // - one standard module, and
    // - one UserForm module so FormsNormalizedData is non-empty.
    let dir_decompressed = {
        let mut out = Vec::new();

        // PROJECTMODULES (count = 2) and PROJECTCOOKIE (arbitrary).
        push_record(&mut out, 0x000F, &2u16.to_le_bytes());
        push_record(&mut out, 0x0013, &0xFFFFu16.to_le_bytes());

        // MODULENAME (standard module)
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
        // MODULEREADONLY / MODULEPRIVATE / MODULE terminator
        push_record(&mut out, 0x0025, b"");
        push_record(&mut out, 0x0028, b"");
        push_record(&mut out, 0x002B, b"");

        // MODULENAME (UserForm/designer module referenced from PROJECT by BaseClass=)
        push_record(&mut out, 0x0019, b"UserForm1");
        // MODULESTREAMNAME + reserved u16
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // MODULETYPE = UserForm (0x0003 per MS-OVBA).
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes());
        // MODULETEXTOFFSET
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());
        push_record(&mut out, 0x0025, b"");
        push_record(&mut out, 0x0028, b"");
        push_record(&mut out, 0x002B, b"");

        // dir terminator
        push_record(&mut out, 0x0010, b"");
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    ole.create_storage("UserForm1").expect("designer storage");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\nModule=Module1\r\nBaseClass=\"UserForm1\"\r\n")
            .expect("write PROJECT");
    }

    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }
    {
        let mut s = ole
            .create_stream("VBA/UserForm1")
            .expect("userform module stream");
        s.write_all(&userform_container)
            .expect("write userform module");
    }

    // Designer payload so FormsNormalizedData is non-empty (and therefore bound by v3 digest).
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("designer stream");
        s.write_all(designer_payload)
            .expect("write designer payload");
    }

    if let Some(sig) = signature_blob {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignatureExt")
            .expect("signature stream");
        s.write_all(sig).expect("write signature");
    }

    ole.into_inner().into_inner()
}

fn build_single_userform_vba_project_bin_v3(
    signature_blob: Option<&[u8]>,
    designer_payload: &[u8],
) -> Vec<u8> {
    let userform_source = b"Sub FormHello()\r\nEnd Sub\r\n";
    let userform_container = compress_container(userform_source);

    // `PROJECT` must reference the designer module via `BaseClass=` so `FormsNormalizedData` is
    // non-empty.
    let project_stream = b"Name=\"VBAProject\"\r\nBaseClass=\"UserForm1\"\r\n";

    // Minimal decompressed `VBA/dir` stream describing a single UserForm module.
    let dir_decompressed = {
        let mut out = Vec::new();

        // MODULENAME (UserForm/designer module)
        push_record(&mut out, 0x0019, b"UserForm1");
        // MODULESTREAMNAME + reserved u16
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // MODULETYPE = UserForm (0x0003 per MS-OVBA).
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes());
        // MODULETEXTOFFSET
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());
        out
    };
    let dir_container = compress_container(&dir_decompressed);

    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    ole.create_storage("VBA").expect("VBA storage");
    ole.create_storage("UserForm1").expect("designer storage");

    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(project_stream).expect("write PROJECT");
    }

    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole
            .create_stream("VBA/UserForm1")
            .expect("userform module stream");
        s.write_all(&userform_container)
            .expect("write userform module");
    }

    // Designer payload so FormsNormalizedData is non-empty.
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("designer stream");
        s.write_all(designer_payload)
            .expect("write designer payload");
    }

    if let Some(sig) = signature_blob {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignatureExt")
            .expect("signature stream");
        s.write_all(sig).expect("write signature");
    }

    ole.into_inner().into_inner()
}

fn build_signature_part_ole(signature_stream_payload: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignatureExt")
            .expect("signature stream");
        s.write_all(signature_stream_payload)
            .expect("write signature");
    }
    ole.into_inner().into_inner()
}

fn der_len(len: usize) -> Vec<u8> {
    if len < 0x80 {
        return vec![len as u8];
    }
    let mut buf = Vec::new();
    let mut n = len;
    while n > 0 {
        buf.push((n & 0xFF) as u8);
        n >>= 8;
    }
    buf.reverse();
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(1usize.saturating_add(buf.len()));
    out.push(0x80 | (buf.len() as u8));
    out.extend_from_slice(&buf);
    out
}

fn der_tlv(tag: u8, content: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(tag);
    out.extend_from_slice(&der_len(content.len()));
    out.extend_from_slice(content);
    out
}

fn der_sequence(items: &[Vec<u8>]) -> Vec<u8> {
    let mut content = Vec::new();
    for item in items {
        content.extend_from_slice(item);
    }
    der_tlv(0x30, &content)
}

fn der_oid(oid_content: &[u8]) -> Vec<u8> {
    der_tlv(0x06, oid_content)
}

fn der_null() -> Vec<u8> {
    vec![0x05, 0x00]
}

fn der_octet_string(bytes: &[u8]) -> Vec<u8> {
    der_tlv(0x04, bytes)
}

fn build_spc_indirect_data_content_sha256(project_digest: &[u8]) -> Vec<u8> {
    // SHA-256 OID: 2.16.840.1.101.3.4.2.1
    let sha256_oid = der_oid(&[0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01]);
    let alg_id = der_sequence(&[sha256_oid, der_null()]);

    // DigestInfo ::= SEQUENCE { digestAlgorithm AlgorithmIdentifier, digest OCTET STRING }
    let digest_info = der_sequence(&[alg_id, der_octet_string(project_digest)]);

    // SpcIndirectDataContent ::= SEQUENCE { data, messageDigest }
    // `data` is ignored by our parser; use NULL.
    der_sequence(&[der_null(), digest_info])
}

fn build_spc_indirect_data_content_md5(project_digest: &[u8]) -> Vec<u8> {
    // MD5 OID: 1.2.840.113549.2.5
    let md5_oid = der_oid(&[0x2A, 0x86, 0x48, 0x86, 0xF7, 0x0D, 0x02, 0x05]);
    let alg_id = der_sequence(&[md5_oid, der_null()]);

    // DigestInfo ::= SEQUENCE { digestAlgorithm AlgorithmIdentifier, digest OCTET STRING }
    let digest_info = der_sequence(&[alg_id, der_octet_string(project_digest)]);

    // SpcIndirectDataContent ::= SEQUENCE { data, messageDigest }
    // `data` is ignored by our parser; use NULL.
    der_sequence(&[der_null(), digest_info])
}

#[test]
fn verify_v3_md5_digest_does_not_bind_when_stream_kind_is_unknown() {
    let project_ole = build_minimal_vba_project_bin_v3(None, b"ABC");
    let digest = compute_vba_project_digest_v3(&project_ole, DigestAlg::Md5).expect("digest v3");
    assert_eq!(digest.len(), 16, "MD5 digest must be 16 bytes");

    // Ensure this test is actually exercising a non-legacy digest value.
    let legacy_content = content_hash_md5(&project_ole).expect("Content hash MD5");
    let legacy_agile = agile_content_hash_md5(&project_ole)
        .expect("Agile content hash MD5")
        .expect("designer present");
    assert_ne!(digest.as_slice(), legacy_content.as_ref());
    assert_ne!(digest.as_slice(), legacy_agile.as_ref());

    let signed_content = build_spc_indirect_data_content_md5(&digest);
    let pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&signed_content);
    let mut signature_stream_payload = signed_content.clone();
    signature_stream_payload.extend_from_slice(&pkcs7);

    // When the signature bytes are provided without an OLE stream name (for example, raw PKCS#7
    // bytes from an external signature part), we still want best-effort binding verification.
    //
    // In that situation we use a heuristic based on digest length:
    // - 16-byte digests are treated as legacy v1/v2 candidates (MD5)
    // - 32-byte digests are treated as v3 candidates (`contents_hash_v3`, currently SHA-256)
    //
    // An out-of-spec MD5 digest over the v3 transcript should therefore not bind.
    let binding = verify_vba_signature_binding(&project_ole, &signature_stream_payload);
    assert_eq!(binding, VbaSignatureBinding::NotBound);

    let binding2 =
        verify_vba_signature_binding_with_stream_path(&project_ole, "", &signature_stream_payload);
    assert_eq!(binding2, VbaSignatureBinding::NotBound);

    let binding3 = verify_vba_project_signature_binding(&project_ole, &signature_stream_payload)
        .expect("binding");
    match binding3 {
        VbaProjectBindingVerification::BoundMismatch(debug) => {
            assert_eq!(debug.signed_digest.as_deref(), Some(digest.as_slice()));
            assert_eq!(debug.computed_digest.as_deref(), Some(legacy_content.as_ref()));
        }
        other => panic!("expected BoundMismatch, got {other:?}"),
    }
}

#[test]
fn digital_signature_ext_uses_v3_project_digest_for_binding() {
    let unsigned = build_minimal_vba_project_bin_v3(None, b"ABC");
    let digest = contents_hash_v3(&unsigned).expect("ContentsHashV3");
    assert_eq!(digest.len(), 32, "SHA-256 digest must be 32 bytes");

    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&signed_content);
    let mut signature_stream = signed_content.clone();
    signature_stream.extend_from_slice(&pkcs7);

    let signed = build_minimal_vba_project_bin_v3(Some(&signature_stream), b"ABC");
    let sig = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
    assert_eq!(
        sig.stream_kind,
        VbaSignatureStreamKind::DigitalSignatureExt,
        "expected DigitalSignatureExt stream, got {}",
        sig.stream_path
    );

    // `formula-xlsx` prefixes OLE stream paths with the source part name.
    let prefixed_path = format!("xl/vbaProjectSignature.bin:{}", sig.stream_path);
    let binding = verify_vba_signature_binding_with_stream_path(
        &signed,
        &prefixed_path,
        &signature_stream,
    );
    assert_eq!(
        binding,
        VbaSignatureBinding::Bound,
        "expected binding to remain Bound for prefixed stream path {prefixed_path}"
    );

    // If callers don't know the original OLE stream name, we still attempt v3 binding
    // best-effort by comparing against ContentsHashV3.
    let binding2 = verify_vba_signature_binding(&signed, &signature_stream);
    assert_eq!(binding2, VbaSignatureBinding::Bound);

    let bound = verify_vba_digital_signature_bound(&signed)
        .expect("bound verify")
        .expect("signature present");
    assert!(matches!(
        bound.binding,
        VbaProjectBindingVerification::BoundVerified(_)
    ));
}

#[test]
fn digital_signature_ext_does_not_bind_md5_digest_bytes_even_when_oid_is_sha256() {
    // ---- 1) Build a minimal V3 project with non-empty FormsNormalizedData ----
    let unsigned = build_single_userform_vba_project_bin_v3(None, b"ABC");

    // ---- 2) Compute an out-of-spec MD5 digest over the v3 transcript (16 bytes) ----
    let digest_md5 = compute_vba_project_digest_v3(&unsigned, DigestAlg::Md5).expect("digest v3");
    assert_eq!(digest_md5.len(), 16, "MD5 digest must be 16 bytes");

    // ---- 3) Construct SpcIndirectDataContent with SHA-256 OID but MD5 digest bytes ----
    //
    // For `\x05DigitalSignatureExt`, binding uses the v3 digest (`contents_hash_v3` in this crate,
    // currently SHA-256). This signature therefore must *not* be considered bound.
    let signed_content = build_spc_indirect_data_content_sha256(&digest_md5);

    // ---- 4) Sign and store in \x05DigitalSignatureExt ----
    let pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&signed_content);
    let mut signature_stream = signed_content.clone();
    signature_stream.extend_from_slice(&pkcs7);

    let signed = build_single_userform_vba_project_bin_v3(Some(&signature_stream), b"ABC");

    // ---- 5) Verify ----
    let sig = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(
        sig.binding,
        VbaSignatureBinding::NotBound,
        "expected DigitalSignatureExt binding to be NotBound when digest bytes are MD5"
    );
}

#[test]
fn verify_vba_project_signature_binding_supports_v3_signature_part() {
    let project_ole = build_minimal_vba_project_bin_v3(None, b"ABC");
    let digest = contents_hash_v3(&project_ole).expect("ContentsHashV3");

    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&signed_content);
    let mut signature_stream_payload = signed_content.clone();
    signature_stream_payload.extend_from_slice(&pkcs7);

    let signature_part = build_signature_part_ole(&signature_stream_payload);

    let binding =
        verify_vba_project_signature_binding(&project_ole, &signature_part).expect("binding");
    match binding {
        VbaProjectBindingVerification::BoundVerified(debug) => {
            assert_eq!(
                debug.hash_algorithm_oid.as_deref(),
                Some("2.16.840.1.101.3.4.2.1")
            );
            assert_eq!(debug.hash_algorithm_name.as_deref(), Some("SHA-256"));
            assert_eq!(debug.signed_digest.as_deref(), Some(digest.as_slice()));
            assert_eq!(debug.computed_digest.as_deref(), Some(digest.as_slice()));
        }
        other => panic!("expected BoundVerified, got {other:?}"),
    }

    // Tamper with the project bytes (designer payload) and ensure binding mismatch is detected.
    let tampered_project = build_minimal_vba_project_bin_v3(None, b"ABD");
    let tampered_digest = contents_hash_v3(&tampered_project).expect("ContentsHashV3(tampered)");

    let binding2 =
        verify_vba_project_signature_binding(&tampered_project, &signature_part).expect("binding");
    match binding2 {
        VbaProjectBindingVerification::BoundMismatch(debug) => {
            assert_eq!(debug.signed_digest.as_deref(), Some(digest.as_slice()));
            assert_eq!(
                debug.computed_digest.as_deref(),
                Some(tampered_digest.as_slice())
            );
        }
        other => panic!("expected BoundMismatch, got {other:?}"),
    }
}

#[test]
fn verify_vba_project_signature_binding_supports_v3_signature_part_even_when_oid_is_md5() {
    let project_ole = build_minimal_vba_project_bin_v3(None, b"ABC");

    // Ensure this test actually distinguishes v3 (DigitalSignatureExt) from legacy binding.
    let legacy_content = content_hash_md5(&project_ole).expect("Content hash MD5");
    let legacy_agile = agile_content_hash_md5(&project_ole)
        .expect("Agile content hash MD5")
        .expect("designer present");

    let digest = contents_hash_v3(&project_ole).expect("ContentsHashV3");
    assert_eq!(digest.len(), 32, "SHA-256 digest must be 32 bytes");
    assert_ne!(digest.as_slice(), legacy_content.as_ref());
    assert_ne!(digest.as_slice(), legacy_agile.as_ref());

    let signed_content = build_spc_indirect_data_content_md5(&digest);
    let pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&signed_content);
    let mut signature_stream_payload = signed_content.clone();
    signature_stream_payload.extend_from_slice(&pkcs7);

    let signature_part = build_signature_part_ole(&signature_stream_payload);

    let binding =
        verify_vba_project_signature_binding(&project_ole, &signature_part).expect("binding");
    match binding {
        VbaProjectBindingVerification::BoundVerified(debug) => {
            assert_eq!(
                debug.hash_algorithm_oid.as_deref(),
                Some("1.2.840.113549.2.5")
            );
            assert_eq!(debug.hash_algorithm_name.as_deref(), Some("MD5"));
            assert_eq!(debug.signed_digest.as_deref(), Some(digest.as_slice()));
            assert_eq!(debug.computed_digest.as_deref(), Some(digest.as_slice()));
        }
        other => panic!("expected BoundVerified, got {other:?}"),
    }

    // Also ensure binding works when callers provide only the raw signature payload bytes (no OLE
    // stream name). This exercises the "try all candidate digests" logic for 16-byte digests.
    let binding2 = verify_vba_signature_binding(&project_ole, &signature_stream_payload);
    assert_eq!(binding2, VbaSignatureBinding::Bound);
}

#[test]
fn verify_vba_project_signature_binding_v3_reports_mismatch_when_digest_len_is_md5() {
    let project_ole = build_minimal_vba_project_bin_v3(None, b"ABC");

    // Construct an out-of-spec `DigitalSignatureExt` binding digest (MD5-sized) and ensure the
    // v3 binder does not treat it as a valid v3 binding: `DigitalSignatureExt` binds against the
    // v3 digest (`contents_hash_v3` in this crate, currently a 32-byte SHA-256).
    let digest = compute_vba_project_digest_v3(&project_ole, DigestAlg::Md5).expect("digest v3");
    assert_eq!(digest.len(), 16, "MD5 digest must be 16 bytes");
    let expected = contents_hash_v3(&project_ole).expect("ContentsHashV3");
    assert_eq!(expected.len(), 32, "SHA-256 digest must be 32 bytes");

    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&signed_content);
    let mut signature_stream_payload = signed_content.clone();
    signature_stream_payload.extend_from_slice(&pkcs7);

    let signature_part = build_signature_part_ole(&signature_stream_payload);

    let binding =
        verify_vba_project_signature_binding(&project_ole, &signature_part).expect("binding");
    match binding {
        VbaProjectBindingVerification::BoundMismatch(debug) => {
            assert_eq!(
                debug.hash_algorithm_oid.as_deref(),
                Some("2.16.840.1.101.3.4.2.1")
            );
            assert_eq!(debug.hash_algorithm_name.as_deref(), Some("SHA-256"));
            assert_eq!(debug.signed_digest.as_deref(), Some(digest.as_slice()));
            assert_eq!(debug.computed_digest.as_deref(), Some(expected.as_slice()));
        }
        other => panic!("expected BoundMismatch, got {other:?}"),
    }
}

#[test]
fn verify_vba_digital_signature_bound_v3_reports_mismatch_when_digest_len_is_md5() {
    let unsigned = build_minimal_vba_project_bin_v3(None, b"ABC");

    // Compute an out-of-spec MD5 v3 digest and wrap it in a DigestInfo that *claims* to be SHA-256.
    // For `\x05DigitalSignatureExt`, binding is against the v3 digest (`contents_hash_v3` in this
    // crate, currently a 32-byte SHA-256), so this must produce a binding mismatch.
    let digest = compute_vba_project_digest_v3(&unsigned, DigestAlg::Md5).expect("digest v3");
    assert_eq!(digest.len(), 16, "MD5 digest must be 16 bytes");
    let expected = contents_hash_v3(&unsigned).expect("ContentsHashV3");
    assert_eq!(expected.len(), 32, "SHA-256 digest must be 32 bytes");

    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&signed_content);
    let mut signature_stream_payload = signed_content.clone();
    signature_stream_payload.extend_from_slice(&pkcs7);

    let signed = build_minimal_vba_project_bin_v3(Some(&signature_stream_payload), b"ABC");
    let bound = verify_vba_digital_signature_bound(&signed)
        .expect("bound verify")
        .expect("signature present");

    assert_eq!(bound.signature.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(bound.signature.binding, VbaSignatureBinding::NotBound);

    match bound.binding {
        VbaProjectBindingVerification::BoundMismatch(debug) => {
            assert_eq!(
                debug.hash_algorithm_oid.as_deref(),
                Some("2.16.840.1.101.3.4.2.1")
            );
            assert_eq!(debug.hash_algorithm_name.as_deref(), Some("SHA-256"));
            assert_eq!(debug.signed_digest.as_deref(), Some(digest.as_slice()));
            assert_eq!(debug.computed_digest.as_deref(), Some(expected.as_slice()));
        }
        other => panic!("expected BoundMismatch, got {other:?}"),
    }
}

#[test]
fn verify_vba_project_signature_binding_reports_mismatch_for_md5_digests_when_stream_kind_is_unknown(
) {
    let project_ole = build_minimal_vba_project_bin_v3(None, b"ABC");

    // This is an out-of-spec v3 project digest (MD5, 16 bytes).
    // Ensure we don't accidentally match legacy binding digests.
    let legacy_content = content_hash_md5(&project_ole).expect("Content hash MD5");
    let legacy_agile = agile_content_hash_md5(&project_ole)
        .expect("Agile content hash MD5")
        .expect("designer present");

    let digest = compute_vba_project_digest_v3(&project_ole, DigestAlg::Md5).expect("digest v3");
    assert_eq!(digest.len(), 16, "MD5 digest must be 16 bytes");
    assert_ne!(digest.as_slice(), legacy_content.as_ref());
    assert_ne!(digest.as_slice(), legacy_agile.as_ref());

    // Construct a raw signature stream payload (`signed_content || pkcs7`) without any surrounding
    // CFB container or stream-path metadata.
    let signed_content = build_spc_indirect_data_content_md5(&digest);
    let pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&signed_content);
    let mut signature_stream_payload = signed_content.clone();
    signature_stream_payload.extend_from_slice(&pkcs7);

    // `verify_vba_project_signature_binding` should be able to recover by attempting both legacy
    // and v3 comparisons when the stream kind is unknown.
    let binding = verify_vba_project_signature_binding(&project_ole, &signature_stream_payload)
        .expect("binding");
    match binding {
        VbaProjectBindingVerification::BoundMismatch(debug) => {
            assert_eq!(
                debug.hash_algorithm_oid.as_deref(),
                Some("1.2.840.113549.2.5")
            );
            assert_eq!(debug.hash_algorithm_name.as_deref(), Some("MD5"));
            assert_eq!(debug.signed_digest.as_deref(), Some(digest.as_slice()));
            assert_eq!(debug.computed_digest.as_deref(), Some(legacy_content.as_ref()));
        }
        other => panic!("expected BoundMismatch, got {other:?}"),
    }

    // Also cover the simpler binding helper (no debug info).
    assert_eq!(
        verify_vba_signature_binding(&project_ole, &signature_stream_payload),
        VbaSignatureBinding::NotBound
    );
}
