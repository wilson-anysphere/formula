#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, contents_hash_v3, extract_vba_signature_signed_digest,
    project_normalized_data_v3_transcript, v3_content_normalized_data, verify_vba_digital_signature,
    VbaSignatureBinding, VbaSignatureVerification,
};

mod signature_test_utils;

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

fn build_minimal_vba_project_bin_with_designer(
    module_source: &[u8],
    designer_bytes: &[u8],
    signature_stream: Option<&[u8]>,
) -> Vec<u8> {
    let userform_source = b"Sub FormHello()\r\nEnd Sub\r\n";

    // ---- 1) Build the `VBA/dir` stream (decompressed form) describing a module and a UserForm. ----
    let dir_decompressed = {
        let mut out = Vec::new();

        // Include a v3-specific reference record type so the transcript depends on it.
        //
        // REFERENCECONTROL (0x002F) has a structured payload; build enough of it for our
        // normalizer to accept it.
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

        // MODULESTREAMNAME + reserved u16.
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);

        // MODULETYPE (standard)
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());

        // MODULETEXTOFFSET: the module stream is just the compressed container.
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());

        // MODULENAME (UserForm/designer module referenced from PROJECT by BaseClass=)
        push_record(&mut out, 0x0019, b"UserForm1");
        // MODULESTREAMNAME + reserved u16.
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

    // ---- 2) Build a compressed module stream. ----
    let module_container = compress_container(module_source);
    let userform_container = compress_container(userform_source);

    // ---- 3) Construct the OLE/CFB container. ----
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");

    // Root-level streams used by real VBA projects.
    {
        let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
        s.write_all(b"Name=\"VBAProject\"\r\nModule=Module1\r\nBaseClass=\"UserForm1\"\r\n")
            .expect("write PROJECT");
    }

    // VBA storage.
    ole.create_storage("VBA").expect("VBA storage");
    {
        let mut s = ole.create_stream("VBA/dir").expect("dir stream");
        s.write_all(&dir_container).expect("write dir");
    }
    {
        let mut s = ole.create_stream("VBA/Module1").expect("module stream");
        s.write_all(&module_container).expect("write module");
    }

    // UserForm module stream + designer storage (root-level, non-VBA), so `FormsNormalizedData` is
    // non-empty and `ContentsHashV3` changes when designers are tampered with.
    {
        let mut s = ole
            .create_stream("VBA/UserForm1")
            .expect("userform module stream");
        s.write_all(&userform_container)
            .expect("write userform module");
    }

    ole.create_storage("UserForm1")
        .expect("create designer storage");
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("designer stream");
        s.write_all(designer_bytes).expect("write designer bytes");
    }

    // Signature stream: `\x05DigitalSignatureExt`.
    if let Some(sig) = signature_stream {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignatureExt")
            .expect("signature stream");
        s.write_all(sig).expect("write signature");
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

fn build_spc_indirect_data_content_sha256(contents_hash_v3: &[u8]) -> Vec<u8> {
    // SHA-256 OID: 2.16.840.1.101.3.4.2.1
    let sha256_oid = [0x60, 0x86, 0x48, 0x01, 0x65, 0x03, 0x04, 0x02, 0x01];

    let mut alg_id = Vec::new();
    alg_id.extend_from_slice(&der_oid_raw(&sha256_oid));
    alg_id.extend_from_slice(&der_null());
    let alg_id = der_sequence(&alg_id);

    let mut digest_info = Vec::new();
    digest_info.extend_from_slice(&alg_id);
    digest_info.extend_from_slice(&der_octet_string(contents_hash_v3));
    let digest_info = der_sequence(&digest_info);

    let mut spc = Vec::new();
    // `data` (ignored by our parser) – use NULL.
    spc.extend_from_slice(&der_null());
    spc.extend_from_slice(&digest_info);
    der_sequence(&spc)
}

#[test]
fn digital_signature_ext_binds_using_contents_hash_v3() {
    let module_source = b"Sub Hello()\r\nEnd Sub\r\n";
    let designer_bytes = b"DESIGNER";

    // ---- 1) Compute ContentsHashV3 over an unsigned project. ----
    let unsigned =
        build_minimal_vba_project_bin_with_designer(module_source, designer_bytes, None);
    let digest = contents_hash_v3(&unsigned).expect("ContentsHashV3");
    assert_eq!(digest.len(), 32, "expected SHA-256 digest bytes");

    // ---- 2) Build SpcIndirectDataContent and sign it. ----
    let signed_content = build_spc_indirect_data_content_sha256(&digest);
    let pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&signed_content);

    let mut signature_stream = signed_content.clone();
    signature_stream.extend_from_slice(&pkcs7);

    // ---- 3) Store the signature in `\x05DigitalSignatureExt`. ----
    let signed = build_minimal_vba_project_bin_with_designer(
        module_source,
        designer_bytes,
        Some(&signature_stream),
    );

    // The digest should be stable between the unsigned and signed container: signature streams are
    // excluded from the transcript to avoid recursion.
    assert_eq!(
        project_normalized_data_v3_transcript(&unsigned).expect("ProjectNormalizedDataV3(unsigned)"),
        project_normalized_data_v3_transcript(&signed).expect("ProjectNormalizedDataV3(signed)")
    );
    assert_eq!(
        v3_content_normalized_data(&unsigned).expect("V3ContentNormalizedData(unsigned)"),
        v3_content_normalized_data(&signed).expect("V3ContentNormalizedData(signed)")
    );

    let sig = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);

    // Sanity-check the plumbing: the signed digest extracted from the signature stream should
    // match `ContentsHashV3(signed_project)`.
    let extracted = extract_vba_signature_signed_digest(&sig.signature)
        .expect("extract signed digest")
        .expect("signed digest present");
    assert_eq!(extracted.digest, digest);
    assert_eq!(
        contents_hash_v3(&signed).expect("ContentsHashV3(signed)"),
        digest
    );

    assert_eq!(sig.binding, VbaSignatureBinding::Bound);

    // ---- 4) Tamper with the designer stream: PKCS#7 remains valid, binding must break. ----
    let tampered = build_minimal_vba_project_bin_with_designer(
        module_source,
        b"DESIGNER!",
        Some(&signature_stream),
    );
    let sig2 = verify_vba_digital_signature(&tampered)
        .expect("tampered signature verification should succeed")
        .expect("signature should be present");
    assert_eq!(sig2.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig2.binding, VbaSignatureBinding::NotBound);
}
