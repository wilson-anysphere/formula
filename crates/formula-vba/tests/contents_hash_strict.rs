#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{compress_container, content_normalized_data, verify_vba_digital_signature, VbaSignatureBinding, VbaSignatureVerification};
use md5::{Digest as _, Md5};

mod signature_test_utils;

fn push_u16(out: &mut Vec<u8>, v: u16) {
    out.extend_from_slice(&v.to_le_bytes());
}

fn push_u32(out: &mut Vec<u8>, v: u32) {
    out.extend_from_slice(&v.to_le_bytes());
}

fn push_utf16le(out: &mut Vec<u8>, s: &str) {
    for unit in s.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
}

fn build_dir_decompressed_spec(project_name: &str, project_constants: &str, module_name: &str) -> Vec<u8> {
    let project_name_bytes = project_name.as_bytes();
    let constants_bytes = project_constants.as_bytes();

    let mut out = Vec::new();

    // --- PROJECTINFORMATION (MS-OVBA §2.3.4.2.1) ---
    //
    // PROJECTSYSKIND
    push_u16(&mut out, 0x0001);
    push_u32(&mut out, 0x0000_0004);
    push_u32(&mut out, 0x0000_0003); // SysKind: 64-bit Windows

    // PROJECTLCID
    push_u16(&mut out, 0x0002);
    push_u32(&mut out, 0x0000_0004);
    push_u32(&mut out, 0x0000_0409); // en-US

    // PROJECTLCIDINVOKE
    push_u16(&mut out, 0x0014);
    push_u32(&mut out, 0x0000_0004);
    push_u32(&mut out, 0x0000_0409);

    // PROJECTCODEPAGE
    push_u16(&mut out, 0x0003);
    push_u32(&mut out, 0x0000_0002);
    push_u16(&mut out, 1252);

    // PROJECTNAME
    push_u16(&mut out, 0x0004);
    push_u32(&mut out, project_name_bytes.len() as u32);
    out.extend_from_slice(project_name_bytes);

    // PROJECTDOCSTRING (empty)
    push_u16(&mut out, 0x0005);
    push_u32(&mut out, 0);
    push_u16(&mut out, 0x0040);
    push_u32(&mut out, 0);

    // PROJECTHELPFILEPATH (empty)
    push_u16(&mut out, 0x0006);
    push_u32(&mut out, 0);
    push_u16(&mut out, 0x003D);
    push_u32(&mut out, 0);

    // PROJECTHELPCONTEXT
    push_u16(&mut out, 0x0007);
    push_u32(&mut out, 0x0000_0004);
    push_u32(&mut out, 0);

    // PROJECTLIBFLAGS
    push_u16(&mut out, 0x0008);
    push_u32(&mut out, 0x0000_0004);
    push_u32(&mut out, 0);

    // PROJECTVERSION
    push_u16(&mut out, 0x0009);
    push_u32(&mut out, 0); // Reserved
    push_u32(&mut out, 1); // VersionMajor
    push_u16(&mut out, 0); // VersionMinor

    // PROJECTCONSTANTS (MBCS + Unicode)
    push_u16(&mut out, 0x000C);
    push_u32(&mut out, constants_bytes.len() as u32);
    out.extend_from_slice(constants_bytes);
    push_u16(&mut out, 0x003C);
    let mut constants_unicode = Vec::new();
    push_utf16le(&mut constants_unicode, project_constants);
    push_u32(&mut out, constants_unicode.len() as u32);
    out.extend_from_slice(&constants_unicode);

    // --- PROJECTREFERENCES (empty) ---
    // Directly start PROJECTMODULES (0x000F).

    // --- PROJECTMODULES (MS-OVBA §2.3.4.2.3) ---
    push_u16(&mut out, 0x000F);
    push_u32(&mut out, 0x0000_0002); // Size of Count
    push_u16(&mut out, 1); // Count (one module)

    // PROJECTCOOKIE (0x0013)
    push_u16(&mut out, 0x0013);
    push_u32(&mut out, 0x0000_0002);
    push_u16(&mut out, 0xFFFF);

    // --- MODULE record (MS-OVBA §2.3.4.2.3.2) ---
    //
    // MODULENAME
    push_u16(&mut out, 0x0019);
    push_u32(&mut out, module_name.len() as u32);
    out.extend_from_slice(module_name.as_bytes());

    // MODULESTREAMNAME
    push_u16(&mut out, 0x001A);
    push_u32(&mut out, module_name.len() as u32);
    out.extend_from_slice(module_name.as_bytes());
    push_u16(&mut out, 0x0032);
    let mut stream_name_unicode = Vec::new();
    push_utf16le(&mut stream_name_unicode, module_name);
    push_u32(&mut out, stream_name_unicode.len() as u32);
    out.extend_from_slice(&stream_name_unicode);

    // MODULEDOCSTRING (empty)
    push_u16(&mut out, 0x001C);
    push_u32(&mut out, 0);
    push_u16(&mut out, 0x0048);
    push_u32(&mut out, 0);

    // MODULEOFFSET (TextOffset = 0)
    push_u16(&mut out, 0x0031);
    push_u32(&mut out, 0x0000_0004);
    push_u32(&mut out, 0);

    // MODULEHELPCONTEXT
    push_u16(&mut out, 0x001E);
    push_u32(&mut out, 0x0000_0004);
    push_u32(&mut out, 0);

    // MODULECOOKIE
    push_u16(&mut out, 0x002C);
    push_u32(&mut out, 0x0000_0002);
    push_u16(&mut out, 0xFFFF);

    // MODULETYPE (procedural module)
    push_u16(&mut out, 0x0021);
    push_u32(&mut out, 0);

    // Terminator + Reserved
    push_u16(&mut out, 0x002B);
    push_u32(&mut out, 0);

    // --- dir stream terminator ---
    push_u16(&mut out, 0x0010);
    push_u32(&mut out, 0);

    out
}

fn build_vba_project_bin_spec(module_source: &[u8], signature_blob: Option<&[u8]>) -> Vec<u8> {
    let project_name = "VBAProject";
    let module_name = "Module1";
    let project_constants = "Answer=42";

    let dir_decompressed = build_dir_decompressed_spec(project_name, project_constants, module_name);
    let dir_container = compress_container(&dir_decompressed);

    let module_container = compress_container(module_source);

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
        // MODULEOFFSET.TextOffset is 0, so write a compressed container at offset 0.
        s.write_all(&module_container).expect("write module bytes");
    }

    if let Some(sig) = signature_blob {
        let mut s = ole
            .create_stream("\u{0005}DigitalSignature")
            .expect("signature stream");
        s.write_all(sig).expect("write signature bytes");
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
    let mut out = Vec::with_capacity(1 + buf.len());
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

fn build_spc_indirect_data_content_sha1(project_digest: &[u8]) -> Vec<u8> {
    // SHA-1 OID: 1.3.14.3.2.26
    let sha1_oid = der_oid(&[0x2B, 0x0E, 0x03, 0x02, 0x1A]);
    let alg_id = der_sequence(&[sha1_oid, der_null()]);

    // DigestInfo ::= SEQUENCE { digestAlgorithm AlgorithmIdentifier, digest OCTET STRING }
    let digest_info = der_sequence(&[alg_id, der_octet_string(project_digest)]);

    // SpcIndirectDataContent ::= SEQUENCE { data, messageDigest }
    // `data` is ignored by our parser; use NULL.
    der_sequence(&[der_null(), digest_info])
}

#[test]
fn content_normalized_data_parses_spec_dir_stream() {
    let module_source = concat!(
        "Attribute VB_Name = \"Module1\"\r\n",
        "Option Explicit\r",
        "Print \"Attribute\"\n",
        "Sub Foo()\r\n",
        "End Sub",
    )
    .as_bytes()
    .to_vec();

    let vba_project_bin = build_vba_project_bin_spec(&module_source, None);
    let normalized = content_normalized_data(&vba_project_bin).expect("ContentNormalizedData");

    // Spec (MS-OVBA §2.4.2.1) appends line bytes without preserving newline delimiters.
    let expected_module_normalized = b"Option ExplicitPrint \"Attribute\"Sub Foo()End Sub".to_vec();
    let expected = [b"VBAProject".as_slice(), b"Answer=42".as_slice(), expected_module_normalized.as_slice()].concat();

    assert_eq!(normalized, expected);
}

#[test]
fn signature_binding_is_bound_for_spec_dir_stream() {
    let module_source = b"Sub Hello()\r\nEnd Sub\r\n";
    let unsigned = build_vba_project_bin_spec(module_source, None);
    let normalized = content_normalized_data(&unsigned).expect("ContentNormalizedData");
    let digest: [u8; 16] = Md5::digest(&normalized).into();

    let signed_content = build_spc_indirect_data_content_sha1(&digest);
    let pkcs7 = signature_test_utils::make_pkcs7_detached_signature(&signed_content);

    let mut signature_stream = signed_content.clone();
    signature_stream.extend_from_slice(&pkcs7);

    let signed = build_vba_project_bin_spec(module_source, Some(&signature_stream));
    let sig = verify_vba_digital_signature(&signed)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig.binding, VbaSignatureBinding::Bound);

    // Tamper module source bytes: PKCS#7 should still verify (detached signature over signed_content),
    // but binding must fail.
    let tampered_module_source = b"Sub Hello()\r\nMsgBox \"tampered\"\r\nEnd Sub\r\n";
    let tampered = build_vba_project_bin_spec(tampered_module_source, Some(&signature_stream));
    let sig2 = verify_vba_digital_signature(&tampered)
        .expect("signature verification should succeed")
        .expect("signature should be present");

    assert_eq!(sig2.verification, VbaSignatureVerification::SignedVerified);
    assert_eq!(sig2.binding, VbaSignatureBinding::NotBound);
}

