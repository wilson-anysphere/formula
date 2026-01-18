#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, contents_hash_v3, verify_vba_digital_signature, VbaSignatureBinding,
    VbaSignatureStreamKind, VbaSignatureVerification,
};

mod signature_test_utils;

fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
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

fn der_null() -> Vec<u8> {
    vec![0x05, 0x00]
}

fn der_oid_raw(oid: &[u8]) -> Vec<u8> {
    der_tlv(0x06, oid)
}

fn der_octet_string(bytes: &[u8]) -> Vec<u8> {
    der_tlv(0x04, bytes)
}

fn build_spc_indirect_data_content_sha1_oid_with_digest(digest: &[u8]) -> Vec<u8> {
    // SHA-1 OID: 1.3.14.3.2.26
    let sha1_oid = [0x2B, 0x0E, 0x03, 0x02, 0x1A];
    let alg_id = der_sequence(&[der_oid_raw(&sha1_oid), der_null()]);
    let digest_info = der_sequence(&[alg_id, der_octet_string(digest)]);

    // SpcIndirectDataContent ::= SEQUENCE { data, messageDigest }
    der_sequence(&[der_null(), digest_info])
}

fn build_vba_project_bin(signature_stream_payload: Option<&[u8]>) -> Vec<u8> {
    let module_source = b"Sub Hello()\r\nEnd Sub\r\n";
    let module_container = compress_container(module_source);
    let userform_source = b"Sub FormHello()\r\nEnd Sub\r\n";
    let userform_container = compress_container(userform_source);

    // Minimal `dir` stream (decompressed form) with:
    // - one standard module, and
    // - one designer module so FormsNormalizedData is non-empty.
    let dir_decompressed = {
        let mut out = Vec::new();

        // Standard module.
        push_record(&mut out, 0x0019, b"Module1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"Module1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        push_record(&mut out, 0x0021, &0u16.to_le_bytes());
        push_record(&mut out, 0x0031, &0u32.to_le_bytes());

        // UserForm (designer) module referenced from PROJECT by BaseClass=.
        push_record(&mut out, 0x0019, b"UserForm1");
        let mut stream_name = Vec::new();
        stream_name.extend_from_slice(b"UserForm1");
        stream_name.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut out, 0x001A, &stream_name);
        // Use a non-zero reserved/module-type value so the transcript depends on it.
        push_record(&mut out, 0x0021, &0x0003u16.to_le_bytes());
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
        s.write_all(b"Name=\"VBAProject\"\r\nBaseClass=\"UserForm1\"\r\n")
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
    {
        let mut s = ole
            .create_stream("UserForm1/Payload")
            .expect("designer stream");
        s.write_all(b"PAYLOAD").expect("write designer payload");
    }

    if let Some(sig) = signature_stream_payload {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignatureExt")
            .expect("signature stream");
        s.write_all(sig).expect("write signature");
    }

    ole.into_inner().into_inner()
}

#[test]
fn digital_signature_ext_binds_even_when_digestinfo_oid_is_sha1() {
    // Build an unsigned project, compute the v3 binding digest bytes, then embed them in a
    // DigestInfo whose algorithm OID is SHA-1. The binding logic should compare digest bytes to
    // `ContentsHashV3` and ignore the (sometimes inconsistent) OID.
    let unsigned = build_vba_project_bin(None);
    let digest = contents_hash_v3(&unsigned).expect("ContentsHashV3");
    assert_eq!(digest.len(), 32);

    let signed_content = build_spc_indirect_data_content_sha1_oid_with_digest(&digest);
    let pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&signed_content);
    let mut signature_stream_payload = signed_content.clone();
    signature_stream_payload.extend_from_slice(&pkcs7);

    let signed = build_vba_project_bin(Some(&signature_stream_payload));

    let sig = verify_vba_digital_signature(&signed)
        .expect("verify signature")
        .expect("signature present");

    assert_eq!(sig.stream_kind, VbaSignatureStreamKind::DigitalSignatureExt);
    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);
}
