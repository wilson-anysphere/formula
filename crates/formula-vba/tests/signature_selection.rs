#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, verify_vba_digital_signature, VbaSignatureBinding,
    VbaSignatureVerification,
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
    assert!(
        sig.stream_path.contains("DigitalSignatureEx"),
        "expected bound DigitalSignatureEx stream to be selected, got {}",
        sig.stream_path
    );
}
