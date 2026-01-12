#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, compute_vba_project_digest_v3, content_normalized_data,
    forms_normalized_data, list_vba_digital_signatures, verify_vba_digital_signature,
    verify_vba_digital_signature_with_project, verify_vba_signature_binding_with_stream_path,
    DigestAlg, VbaSignatureBinding, VbaSignatureStreamKind, VbaSignatureVerification,
};
use md5::{Digest as _, Md5};

mod signature_test_utils;

fn build_minimal_vba_project_bin_with_signature_streams(
    module1_code: &[u8],
    signature_streams: &[(&str, &[u8])],
) -> Vec<u8> {
    let module_container = compress_container(module1_code);

    let dir_decompressed = {
        let mut out = Vec::new();
        // PROJECTNAME + PROJECTCONSTANTS.
        push_record(&mut out, 0x0004, b"VBAProject");
        push_record(&mut out, 0x000C, b"");

        // Minimal module record group.
        push_record(&mut out, 0x0019, b"Module1"); // MODULENAME
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name); // MODULESTREAMNAME
        push_record(&mut out, 0x0021, &0u16.to_le_bytes()); // MODULETYPE (standard)
        push_record(&mut out, 0x0031, &0u32.to_le_bytes()); // MODULETEXTOFFSET
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

    for (path, bytes) in signature_streams {
        if let Some((parent, _)) = path.rsplit_once('/') {
            ole.create_storage_all(parent)
                .expect("create signature parent storage");
        }
        let mut s = ole.create_stream(path).expect("signature stream");
        s.write_all(bytes).expect("write signature");
    }

    ole.into_inner().into_inner()
}

fn build_minimal_vba_project_bin_v3_with_signature_streams(
    signature_streams: &[(&str, &[u8])],
    designer_payload: &[u8],
) -> Vec<u8> {
    let module_source = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_source);
    let userform_source = b"Sub FormHello()\r\nEnd Sub\r\n";
    let userform_container = compress_container(userform_source);

    // Minimal `dir` stream (decompressed form) with:
    // - a v3-specific reference record,
    // - one standard module, and
    // - one UserForm module so FormsNormalizedData is non-empty.
    let dir_decompressed = {
        let mut out = Vec::new();

        // Include a v3-specific reference record type so the transcript depends on it.
        let libid_twiddled = b"REFCTRL-V3";
        let reserved1: u32 = 0;
        let reserved2: u16 = 0;
        let mut reference_control = Vec::new();
        reference_control.extend_from_slice(&(libid_twiddled.len() as u32).to_le_bytes());
        reference_control.extend_from_slice(libid_twiddled);
        reference_control.extend_from_slice(&reserved1.to_le_bytes());
        reference_control.extend_from_slice(&reserved2.to_le_bytes());
        push_record(&mut out, 0x002F, &reference_control);

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

    for (path, bytes) in signature_streams {
        if let Some((parent, _)) = path.rsplit_once('/') {
            ole.create_storage_all(parent)
                .expect("create signature parent storage");
        }
        let mut s = ole.create_stream(path).expect("signature stream");
        s.write_all(bytes).expect("write signature");
    }

    ole.into_inner().into_inner()
}

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
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

#[test]
fn prefers_bound_verified_signature_stream_over_unbound_verified_candidate() {
    let module1 = b"Sub A()\r\nEnd Sub\r\n";

    // Build an unsigned project first to compute the digest that Office would sign.
    let unsigned = build_minimal_vba_project_bin_with_signature_streams(module1, &[]);
    let normalized = content_normalized_data(&unsigned).expect("content normalized data");
    let digest: [u8; 16] = Md5::digest(&normalized).into();

    // Create a bound signature stream (digest matches the project).
    let bound_content = build_spc_indirect_data_content_sha256(&digest);
    let bound_pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&bound_content);
    let mut bound_stream = bound_content.clone();
    bound_stream.extend_from_slice(&bound_pkcs7);

    // Create an unbound signature stream that is still cryptographically valid, but whose signed
    // digest does not match the current project.
    let mut wrong_digest = digest.clone();
    wrong_digest[0] ^= 0xFF;
    let unbound_content = build_spc_indirect_data_content_sha256(&wrong_digest);
    let unbound_pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&unbound_content);
    let mut unbound_stream = unbound_content.clone();
    unbound_stream.extend_from_slice(&unbound_pkcs7);

    // Include both signature streams; Excel-like stream-name ordering will consider
    // `DigitalSignatureExt` before `DigitalSignatureEx`, so without the bound-selection logic we'd
    // pick the unbound stream.
    let streams = [
        ("\u{0005}DigitalSignatureExt", unbound_stream.as_slice()),
        ("\u{0005}DigitalSignatureEx", bound_stream.as_slice()),
    ];
    let signed = build_minimal_vba_project_bin_with_signature_streams(module1, &streams);

    let sig = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
    assert_eq!(
        sig.stream_kind,
        VbaSignatureStreamKind::DigitalSignatureEx,
        "expected bound DigitalSignatureEx stream to be selected, got {}",
        sig.stream_path
    );
}

#[test]
fn prefers_digital_signature_ex_over_legacy_when_both_verified_and_bound() {
    let module1 = b"Sub A()\r\nEnd Sub\r\n";

    // Build an unsigned project first to compute the digest that Office would sign.
    let unsigned = build_minimal_vba_project_bin_with_signature_streams(module1, &[]);
    let normalized = content_normalized_data(&unsigned).expect("content normalized data");
    let digest: [u8; 16] = Md5::digest(&normalized).into();

    // Create a signature payload that is bound to the project.
    let content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&content);
    let mut signature_stream = content.clone();
    signature_stream.extend_from_slice(&pkcs7);

    // Include both legacy and Ex signature streams. When both are verified+bound, selection should
    // still prefer `DigitalSignatureEx` per Excel-like stream precedence rules.
    let signed = build_minimal_vba_project_bin_with_signature_streams(
        module1,
        &[
            ("\u{0005}DigitalSignature", signature_stream.as_slice()),
            ("\u{0005}DigitalSignatureEx", signature_stream.as_slice()),
        ],
    );

    let sig = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
    assert_eq!(
        sig.stream_kind,
        VbaSignatureStreamKind::DigitalSignatureEx,
        "expected bound DigitalSignatureEx stream to be selected, got {}",
        sig.stream_path
    );
}

#[test]
fn prefers_bound_verified_digital_signature_ext_over_unbound_verified_ex_candidate() {
    // Build an unsigned v3-capable project (includes a v3 reference record and a designer storage)
    // so `compute_vba_project_digest_v3` can derive a v3 digest transcript.
    let unsigned = build_minimal_vba_project_bin_v3_with_signature_streams(&[], b"ABC");
    let digest = compute_vba_project_digest_v3(&unsigned, DigestAlg::Sha256).expect("digest v3");
    assert_eq!(digest.len(), 32, "SHA-256 digest must be 32 bytes");

    // Create a bound `\x05DigitalSignatureExt` stream (digest matches the project).
    let bound_content = build_spc_indirect_data_content_sha256(&digest);
    let bound_pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&bound_content);
    let mut bound_stream = bound_content.clone();
    bound_stream.extend_from_slice(&bound_pkcs7);

    // Create an unbound `\x05DigitalSignatureEx` stream that is still cryptographically valid, but
    // whose signed digest does not match the current project.
    //
    // Note: For legacy `DigitalSignatureEx` signatures, Office stores an MD5 digest (16 bytes) in
    // `DigestInfo.digest` even when the DigestInfo algorithm OID is SHA-256 (MS-OSHARED §4.3). Use a
    // wrong 16-byte digest to guarantee the signature is unbound under legacy binding rules.
    let content_normalized = content_normalized_data(&unsigned).expect("content normalized data");
    let content_hash_md5: [u8; 16] = Md5::digest(&content_normalized).into();
    let forms = forms_normalized_data(&unsigned).expect("forms normalized data");
    assert!(
        !forms.is_empty(),
        "expected FormsNormalizedData to be non-empty (designer payload should contribute)"
    );
    let mut h = Md5::new();
    h.update(&content_normalized);
    h.update(&forms);
    let agile_hash_md5: [u8; 16] = h.finalize().into();
    assert_ne!(
        content_hash_md5, agile_hash_md5,
        "expected designer payload to affect the legacy Agile Content Hash transcript"
    );

    let mut wrong_md5 = content_hash_md5;
    wrong_md5[0] = wrong_md5[0].wrapping_add(1);
    if wrong_md5 == content_hash_md5 || wrong_md5 == agile_hash_md5 {
        wrong_md5[1] ^= 0xFF;
    }
    // Avoid a leading 0x30 that could make the digest look like DER.
    if wrong_md5[0] == 0x30 {
        wrong_md5[0] = 0x31;
    }

    let unbound_content = build_spc_indirect_data_content_sha256(&wrong_md5);
    let unbound_pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&unbound_content);
    let mut unbound_stream = unbound_content.clone();
    unbound_stream.extend_from_slice(&unbound_pkcs7);

    let streams = [
        ("\u{0005}DigitalSignatureEx", unbound_stream.as_slice()),
        ("\u{0005}DigitalSignatureExt", bound_stream.as_slice()),
    ];
    let signed = build_minimal_vba_project_bin_v3_with_signature_streams(&streams, b"ABC");

    // Sanity-check: both signature streams are internally verified.
    let listed = list_vba_digital_signatures(&signed).expect("list signatures");
    assert_eq!(listed.len(), 2, "expected exactly two signature streams");
    for sig in &listed {
        assert_eq!(
            sig.verification,
            VbaSignatureVerification::SignedVerified,
            "expected {} to be cryptographically verified",
            sig.stream_path
        );
    }

    // Ensure the intended binding split: Ext is bound via v3 digest, Ex is not bound.
    let ext = listed
        .iter()
        .find(|s| s.stream_kind == VbaSignatureStreamKind::DigitalSignatureExt)
        .expect("DigitalSignatureExt present");
    let ext_binding = verify_vba_signature_binding_with_stream_path(&signed, &ext.stream_path, &ext.signature);
    assert_eq!(ext_binding, VbaSignatureBinding::Bound);

    let ex = listed
        .iter()
        .find(|s| s.stream_kind == VbaSignatureStreamKind::DigitalSignatureEx)
        .expect("DigitalSignatureEx present");
    let ex_binding = verify_vba_signature_binding_with_stream_path(&signed, &ex.stream_path, &ex.signature);
    assert_eq!(ex_binding, VbaSignatureBinding::NotBound);

    // Finally, assert the selection logic returns the verified+bound `DigitalSignatureExt` stream.
    let chosen = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(chosen.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(chosen.binding, VbaSignatureBinding::Bound);
    assert_eq!(
        chosen.stream_kind,
        VbaSignatureStreamKind::DigitalSignatureExt,
        "expected bound DigitalSignatureExt stream to be selected, got {}",
        chosen.stream_path
    );
}

#[test]
fn prefers_bound_verified_digital_signature_ext_when_signatures_are_in_separate_container() {
    // Some XLSM producers store signature streams in a dedicated OLE part
    // (`xl/vbaProjectSignature.bin`) instead of embedding them in `vbaProject.bin`. Ensure the
    // selection logic remains consistent in that configuration.
    let project = build_minimal_vba_project_bin_v3_with_signature_streams(&[], b"ABC");
    let digest = compute_vba_project_digest_v3(&project, DigestAlg::Sha256).expect("digest v3");

    // Bound `DigitalSignatureExt` stream (v3 digest matches the project).
    let bound_content = build_spc_indirect_data_content_sha256(&digest);
    let bound_pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&bound_content);
    let mut bound_stream = bound_content.clone();
    bound_stream.extend_from_slice(&bound_pkcs7);

    // Verified but unbound `DigitalSignatureEx` stream (wrong legacy MD5 digest).
    let content_normalized = content_normalized_data(&project).expect("content normalized data");
    let content_hash_md5: [u8; 16] = Md5::digest(&content_normalized).into();
    let forms = forms_normalized_data(&project).expect("forms normalized data");
    assert!(
        !forms.is_empty(),
        "expected FormsNormalizedData to be non-empty (designer payload should contribute)"
    );
    let mut h = Md5::new();
    h.update(&content_normalized);
    h.update(&forms);
    let agile_hash_md5: [u8; 16] = h.finalize().into();
    assert_ne!(
        content_hash_md5, agile_hash_md5,
        "expected designer payload to affect the legacy Agile Content Hash transcript"
    );

    let mut wrong_md5 = content_hash_md5;
    wrong_md5[0] = wrong_md5[0].wrapping_add(1);
    if wrong_md5 == content_hash_md5 || wrong_md5 == agile_hash_md5 {
        wrong_md5[1] ^= 0xFF;
    }
    if wrong_md5[0] == 0x30 {
        wrong_md5[0] = 0x31;
    }

    let unbound_content = build_spc_indirect_data_content_sha256(&wrong_md5);
    let unbound_pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&unbound_content);
    let mut unbound_stream = unbound_content.clone();
    unbound_stream.extend_from_slice(&unbound_pkcs7);

    let signature_container = signature_test_utils::build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignatureEx", unbound_stream.as_slice()),
        ("\u{0005}DigitalSignatureExt", bound_stream.as_slice()),
    ]);

    let chosen = verify_vba_digital_signature_with_project(&project, &signature_container)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(chosen.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(chosen.binding, VbaSignatureBinding::Bound);
    assert_eq!(chosen.stream_kind, VbaSignatureStreamKind::DigitalSignatureExt);
}

#[test]
fn prefers_bound_verified_digital_signature_ext_when_signature_stream_is_nested_in_separate_container() {
    // Some XLSM producers store signatures in a dedicated OLE part, and some producers also nest
    // signature streams under a storage (e.g. `\x05DigitalSignatureExt/sig`). Ensure selection
    // works in that combined configuration.
    let project = build_minimal_vba_project_bin_v3_with_signature_streams(&[], b"ABC");
    let digest = compute_vba_project_digest_v3(&project, DigestAlg::Sha256).expect("digest v3");

    // Bound `DigitalSignatureExt` stream (v3 digest matches the project).
    let bound_content = build_spc_indirect_data_content_sha256(&digest);
    let bound_pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&bound_content);
    let mut bound_stream = bound_content.clone();
    bound_stream.extend_from_slice(&bound_pkcs7);

    // Verified but unbound `DigitalSignatureEx` stream (wrong legacy MD5 digest).
    let content_normalized = content_normalized_data(&project).expect("content normalized data");
    let content_hash_md5: [u8; 16] = Md5::digest(&content_normalized).into();
    let forms = forms_normalized_data(&project).expect("forms normalized data");
    assert!(
        !forms.is_empty(),
        "expected FormsNormalizedData to be non-empty (designer payload should contribute)"
    );
    let mut h = Md5::new();
    h.update(&content_normalized);
    h.update(&forms);
    let agile_hash_md5: [u8; 16] = h.finalize().into();
    assert_ne!(
        content_hash_md5, agile_hash_md5,
        "expected designer payload to affect the legacy Agile Content Hash transcript"
    );

    let mut wrong_md5 = content_hash_md5;
    wrong_md5[0] = wrong_md5[0].wrapping_add(1);
    if wrong_md5 == content_hash_md5 || wrong_md5 == agile_hash_md5 {
        wrong_md5[1] ^= 0xFF;
    }
    if wrong_md5[0] == 0x30 {
        wrong_md5[0] = 0x31;
    }

    let unbound_content = build_spc_indirect_data_content_sha256(&wrong_md5);
    let unbound_pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&unbound_content);
    let mut unbound_stream = unbound_content.clone();
    unbound_stream.extend_from_slice(&unbound_pkcs7);

    let signature_container = signature_test_utils::build_vba_project_bin_with_signature_streams(&[
        ("\u{0005}DigitalSignatureEx/sig", unbound_stream.as_slice()),
        ("\u{0005}DigitalSignatureExt/sig", bound_stream.as_slice()),
    ]);

    let chosen = verify_vba_digital_signature_with_project(&project, &signature_container)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(chosen.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(chosen.binding, VbaSignatureBinding::Bound);
    assert_eq!(chosen.stream_kind, VbaSignatureStreamKind::DigitalSignatureExt);
    assert_eq!(chosen.stream_path, "\u{0005}DigitalSignatureExt/sig");
}

#[test]
fn prefers_bound_verified_digital_signature_ext_when_signature_stream_is_nested_under_storage() {
    // Some producers store signature streams inside a storage (e.g. `\x05DigitalSignatureExt/sig`).
    // Ensure the selection logic (including v3 binding for `DigitalSignatureExt`) remains correct.
    let project = build_minimal_vba_project_bin_v3_with_signature_streams(&[], b"ABC");
    let digest = compute_vba_project_digest_v3(&project, DigestAlg::Sha256).expect("digest v3");

    // Bound v3 `DigitalSignatureExt` stream.
    let bound_content = build_spc_indirect_data_content_sha256(&digest);
    let bound_pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&bound_content);
    let mut bound_stream = bound_content.clone();
    bound_stream.extend_from_slice(&bound_pkcs7);

    // Verified but unbound legacy `DigitalSignatureEx` stream (wrong MD5 digest).
    let content_normalized = content_normalized_data(&project).expect("content normalized data");
    let content_hash_md5: [u8; 16] = Md5::digest(&content_normalized).into();
    let forms = forms_normalized_data(&project).expect("forms normalized data");
    assert!(
        !forms.is_empty(),
        "expected FormsNormalizedData to be non-empty (designer payload should contribute)"
    );
    let mut h = Md5::new();
    h.update(&content_normalized);
    h.update(&forms);
    let agile_hash_md5: [u8; 16] = h.finalize().into();
    assert_ne!(
        content_hash_md5, agile_hash_md5,
        "expected designer payload to affect the legacy Agile Content Hash transcript"
    );

    let mut wrong_md5 = content_hash_md5;
    wrong_md5[0] = wrong_md5[0].wrapping_add(1);
    if wrong_md5 == content_hash_md5 || wrong_md5 == agile_hash_md5 {
        wrong_md5[1] ^= 0xFF;
    }
    if wrong_md5[0] == 0x30 {
        wrong_md5[0] = 0x31;
    }

    let unbound_content = build_spc_indirect_data_content_sha256(&wrong_md5);
    let unbound_pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&unbound_content);
    let mut unbound_stream = unbound_content.clone();
    unbound_stream.extend_from_slice(&unbound_pkcs7);

    let signed = build_minimal_vba_project_bin_v3_with_signature_streams(
        &[
            ("\u{0005}DigitalSignatureEx/sig", unbound_stream.as_slice()),
            ("\u{0005}DigitalSignatureExt/sig", bound_stream.as_slice()),
        ],
        b"ABC",
    );

    let chosen = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(chosen.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(chosen.binding, VbaSignatureBinding::Bound);
    assert_eq!(chosen.stream_kind, VbaSignatureStreamKind::DigitalSignatureExt);
    assert_eq!(chosen.stream_path, "\u{0005}DigitalSignatureExt/sig");
}

#[test]
fn prefers_digital_signature_ext_over_ex_when_both_verified_and_bound() {
    // When multiple signature streams are present and *both* are verified+bound, we should still
    // prefer the newest stream name (`DigitalSignatureExt`) per Excel-like precedence rules.
    let project = build_minimal_vba_project_bin_v3_with_signature_streams(&[], b"ABC");

    // Bound `DigitalSignatureExt` stream (v3 digest).
    let digest_v3 = compute_vba_project_digest_v3(&project, DigestAlg::Sha256).expect("digest v3");
    let ext_content = build_spc_indirect_data_content_sha256(&digest_v3);
    let ext_pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&ext_content);
    let mut ext_stream = ext_content.clone();
    ext_stream.extend_from_slice(&ext_pkcs7);

    // Bound `DigitalSignatureEx` stream (legacy binding via Agile Content Hash MD5).
    let content_normalized = content_normalized_data(&project).expect("content normalized data");
    let forms = forms_normalized_data(&project).expect("forms normalized data");
    assert!(
        !forms.is_empty(),
        "expected FormsNormalizedData to be non-empty (designer payload should contribute)"
    );
    let mut h = Md5::new();
    h.update(&content_normalized);
    h.update(&forms);
    let agile_hash_md5: [u8; 16] = h.finalize().into();

    let ex_content = build_spc_indirect_data_content_sha256(&agile_hash_md5);
    let ex_pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&ex_content);
    let mut ex_stream = ex_content.clone();
    ex_stream.extend_from_slice(&ex_pkcs7);

    let signed = build_minimal_vba_project_bin_v3_with_signature_streams(
        &[
            ("\u{0005}DigitalSignatureEx", ex_stream.as_slice()),
            ("\u{0005}DigitalSignatureExt", ext_stream.as_slice()),
        ],
        b"ABC",
    );

    // Sanity-check: both signature streams are verified and bound.
    let listed = list_vba_digital_signatures(&signed).expect("list signatures");
    assert_eq!(listed.len(), 2, "expected exactly two signature streams");
    for sig in &listed {
        assert_eq!(
            sig.verification,
            VbaSignatureVerification::SignedVerified,
            "expected {} to be cryptographically verified",
            sig.stream_path
        );
        let binding =
            verify_vba_signature_binding_with_stream_path(&signed, &sig.stream_path, &sig.signature);
        assert_eq!(
            binding,
            VbaSignatureBinding::Bound,
            "expected {} to be bound",
            sig.stream_path
        );
    }

    // Assert selection prefers DigitalSignatureExt.
    let chosen = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");
    assert_eq!(chosen.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(chosen.binding, VbaSignatureBinding::Bound);
    assert_eq!(chosen.stream_kind, VbaSignatureStreamKind::DigitalSignatureExt);
}
