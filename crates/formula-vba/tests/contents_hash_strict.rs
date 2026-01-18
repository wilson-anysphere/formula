#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write};

use formula_vba::{
    compress_container, content_normalized_data, project_normalized_data,
    project_normalized_data_v3_dir_records, v3_content_normalized_data,
    verify_vba_digital_signature, VBAProject, VbaSignatureBinding, VbaSignatureVerification,
};
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

fn push_referencename_record(out: &mut Vec<u8>, name: &[u8]) {
    // REFERENCENAME (0x0016) + Unicode marker (0x003E)
    //
    // This record does not contribute to v1/v2 ContentNormalizedData, but spec-compliant `VBA/dir`
    // streams can include it before many reference records.
    push_u16(out, 0x0016);
    push_u32(out, name.len() as u32);
    out.extend_from_slice(name);
    // Reserved marker + Unicode bytes (empty for this synthetic fixture).
    push_u16(out, 0x003E);
    push_u32(out, 0);
}

fn build_dir_decompressed_spec_with_references(
    project_name: &str,
    project_constants: &str,
    module_name: &str,
    projectversion_reserved: u32,
    references: &[u8],
) -> Vec<u8> {
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
    push_u32(&mut out, projectversion_reserved); // Reserved
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

    // --- PROJECTREFERENCES ---
    out.extend_from_slice(references);

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

fn build_dir_decompressed_spec_with_references_and_stream_name(
    project_name: &str,
    project_constants: &str,
    module_name: &str,
    module_stream_name_ansi: &str,
    module_stream_name_unicode_bytes: &[u8],
    projectversion_reserved: u32,
    references: &[u8],
) -> Vec<u8> {
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
    push_u32(&mut out, projectversion_reserved); // Reserved
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

    // --- PROJECTREFERENCES ---
    out.extend_from_slice(references);

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

    // MODULESTREAMNAME (with Reserved=0x0032 + StreamNameUnicode)
    push_u16(&mut out, 0x001A);
    push_u32(&mut out, module_stream_name_ansi.len() as u32);
    out.extend_from_slice(module_stream_name_ansi.as_bytes());
    push_u16(&mut out, 0x0032);
    push_u32(&mut out, module_stream_name_unicode_bytes.len() as u32);
    out.extend_from_slice(module_stream_name_unicode_bytes);

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

fn build_dir_decompressed_spec_with_alternate_unicode_record_ids(
    project_name: &str,
    project_constants: &str,
    module_name: &str,
    module_stream_name_ansi: &str,
    module_stream_name_unicode: &str,
) -> Vec<u8> {
    let project_name_bytes = project_name.as_bytes();
    let constants_bytes = project_constants.as_bytes();

    let mut out = Vec::new();

    // --- PROJECTINFORMATION (MS-OVBA §2.3.4.2.1) ---
    //
    // Use alternate IDs for some Unicode sub-record markers to exercise permissive parsing.
    // These variants are seen in some real-world files.
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

    // PROJECTDOCSTRING (empty) + alternate Unicode marker 0x0041.
    push_u16(&mut out, 0x0005);
    push_u32(&mut out, 0);
    push_u16(&mut out, 0x0041);
    push_u32(&mut out, 0);

    // PROJECTHELPFILEPATH (empty) + alternate second path marker 0x0042.
    push_u16(&mut out, 0x0006);
    push_u32(&mut out, 0);
    push_u16(&mut out, 0x0042);
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

    // PROJECTCONSTANTS (MBCS + alternate Unicode marker 0x0043).
    push_u16(&mut out, 0x000C);
    push_u32(&mut out, constants_bytes.len() as u32);
    out.extend_from_slice(constants_bytes);
    push_u16(&mut out, 0x0043);
    let mut constants_unicode = Vec::new();
    push_utf16le(&mut constants_unicode, project_constants);
    push_u32(&mut out, constants_unicode.len() as u32);
    out.extend_from_slice(&constants_unicode);

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

    // MODULESTREAMNAME + MODULESTREAMNAMEUNICODE encoded as a distinct record ID (0x0048).
    // The ANSI stream name is deliberately wrong so resolution depends on the Unicode field.
    push_u16(&mut out, 0x001A);
    push_u32(&mut out, module_stream_name_ansi.len() as u32);
    out.extend_from_slice(module_stream_name_ansi.as_bytes());

    push_u16(&mut out, 0x0048);
    let mut stream_name_unicode = Vec::new();
    push_utf16le(&mut stream_name_unicode, module_stream_name_unicode);
    push_u32(&mut out, stream_name_unicode.len() as u32);
    out.extend_from_slice(&stream_name_unicode);

    // MODULEDOCSTRING (empty) using alternate id 0x001B + Unicode marker 0x0049.
    push_u16(&mut out, 0x001B);
    push_u32(&mut out, 0);
    push_u16(&mut out, 0x0049);
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

fn build_dir_decompressed_spec(project_name: &str, project_constants: &str, module_name: &str) -> Vec<u8> {
    build_dir_decompressed_spec_with_references(project_name, project_constants, module_name, 0, &[])
}

fn build_vba_project_bin_spec_with_dir(
    dir_decompressed: &[u8],
    module_source: &[u8],
    signature_blob: Option<&[u8]>,
) -> Vec<u8> {
    let dir_container = compress_container(dir_decompressed);
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

fn build_vba_project_bin_spec(module_source: &[u8], signature_blob: Option<&[u8]>) -> Vec<u8> {
    let project_name = "VBAProject";
    let module_name = "Module1";
    let project_constants = "Answer=42";

    let dir_decompressed =
        build_dir_decompressed_spec(project_name, project_constants, module_name);
    build_vba_project_bin_spec_with_dir(&dir_decompressed, module_source, signature_blob)
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

    // Spec (MS-OVBA §2.4.2.1) strips Attribute lines and normalizes line endings to CRLF.
    let expected_module_normalized = concat!(
        "Option Explicit\r\n",
        "Print \"Attribute\"\r\n",
        "Sub Foo()\r\n",
        "End Sub\r\n",
    )
    .as_bytes()
    .to_vec();
    let expected = [
        b"VBAProject".as_slice(),
        b"Answer=42".as_slice(),
        expected_module_normalized.as_slice(),
    ]
    .concat();

    assert_eq!(normalized, expected);
}

#[test]
fn content_normalized_data_parses_spec_dir_stream_with_fixed_length_projectversion_reserved_4() {
    // Regression: some real-world projects store PROJECTVERSION (0x0009) using the fixed-length
    // layout with Reserved(u32)=4. The strict dir stream parser must still scan to the module
    // records and compute ContentNormalizedData correctly.
    let module_source = concat!(
        "Attribute VB_Name = \"Module1\"\r\n",
        "Option Explicit\r",
        "Print \"Attribute\"\n",
        "Sub Foo()\r\n",
        "End Sub",
    )
    .as_bytes()
    .to_vec();

    let project_name = "VBAProject";
    let module_name = "Module1";
    let project_constants = "Answer=42";
    let dir_decompressed =
        build_dir_decompressed_spec_with_references(project_name, project_constants, module_name, 4, &[]);
    let vba_project_bin =
        build_vba_project_bin_spec_with_dir(&dir_decompressed, &module_source, None);

    let normalized = content_normalized_data(&vba_project_bin).expect("ContentNormalizedData");

    // Spec (MS-OVBA §2.4.2.1) strips Attribute lines and normalizes line endings to CRLF.
    let expected_module_normalized = concat!(
        "Option Explicit\r\n",
        "Print \"Attribute\"\r\n",
        "Sub Foo()\r\n",
        "End Sub\r\n",
    )
    .as_bytes()
    .to_vec();
    let expected = [
        b"VBAProject".as_slice(),
        b"Answer=42".as_slice(),
        expected_module_normalized.as_slice(),
    ]
    .concat();

    assert_eq!(normalized, expected);
}

#[test]
fn content_normalized_data_parses_spec_dir_stream_with_unicode_module_stream_name() {
    // Ensure the strict MS-OVBA dir-stream parser can resolve module streams when the
    // MODULESTREAMNAME Unicode field is required (non-ASCII OLE stream name).
    let project_name = "VBAProject";
    let module_name = "Module1";
    let project_constants = "Answer=42";

    let module_stream_name_unicode = "МодульПоток"; // non-ASCII

    // StreamNameUnicode bytes: `u32 byte_len || utf16le_bytes || trailing_nul`.
    // This matches a pattern seen in real-world files and should still decode correctly.
    let mut stream_name_utf16 = Vec::new();
    push_utf16le(&mut stream_name_utf16, module_stream_name_unicode);
    stream_name_utf16.extend_from_slice(&0u16.to_le_bytes()); // NUL terminator (defensive)
    let mut module_stream_name_unicode_bytes =
        (stream_name_utf16.len() as u32).to_le_bytes().to_vec();
    module_stream_name_unicode_bytes.extend_from_slice(&stream_name_utf16);

    let dir_decompressed = build_dir_decompressed_spec_with_references_and_stream_name(
        project_name,
        project_constants,
        module_name,
        "WrongStreamName",
        &module_stream_name_unicode_bytes,
        0,
        &[],
    );

    let module_source = b"Sub Foo()\r\nEnd Sub\r\n";
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
        let stream_path = format!("VBA/{module_stream_name_unicode}");
        let mut s = ole.create_stream(&stream_path).expect("module stream");
        // MODULEOFFSET.TextOffset is 0, so write a compressed container at offset 0.
        s.write_all(&module_container).expect("write module bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();
    let normalized = content_normalized_data(&vba_project_bin).expect("ContentNormalizedData");

    let expected_module_normalized = b"Sub Foo()\r\nEnd Sub\r\n".as_slice();
    let expected = [
        project_name.as_bytes(),
        project_constants.as_bytes(),
        expected_module_normalized,
    ]
    .concat();
    assert_eq!(normalized, expected);
}

#[test]
fn content_normalized_data_parses_spec_dir_stream_with_alternate_unicode_record_ids() {
    // Regression/robustness: some real-world `VBA/dir` streams use alternate IDs for Unicode
    // sub-record markers (e.g. 0x0041/0x0042/0x0043) and/or store MODULESTREAMNAMEUNICODE as a
    // separate record (0x0048).
    //
    // Ensure our strict MS-OVBA parser stays aligned and can still resolve a Unicode-only module
    // stream name for lookup.
    let project_name = "VBAProject";
    let module_name = "Module1";
    let project_constants = "Answer=42";

    let module_stream_name_unicode = "МодульПоток"; // non-ASCII

    let dir_decompressed = build_dir_decompressed_spec_with_alternate_unicode_record_ids(
        project_name,
        project_constants,
        module_name,
        "WrongStreamName",
        module_stream_name_unicode,
    );

    let module_source = b"Sub Foo()\r\nEnd Sub\r\n";
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
        let stream_path = format!("VBA/{module_stream_name_unicode}");
        let mut s = ole.create_stream(&stream_path).expect("module stream");
        // MODULEOFFSET.TextOffset is 0, so write a compressed container at offset 0.
        s.write_all(&module_container).expect("write module bytes");
    }

    let vba_project_bin = ole.into_inner().into_inner();

    // ContentNormalizedData must be computed successfully and include the normalized module source.
    let normalized = content_normalized_data(&vba_project_bin).expect("ContentNormalizedData");
    let expected_module_normalized = b"Sub Foo()\r\nEnd Sub\r\n".as_slice();
    let expected = [
        project_name.as_bytes(),
        project_constants.as_bytes(),
        expected_module_normalized,
    ]
    .concat();
    assert_eq!(normalized, expected);

    // Ensure the general VBA project parser can also resolve the Unicode module stream name record.
    let project = VBAProject::parse(&vba_project_bin).expect("parse VBA project");
    assert_eq!(project.modules.len(), 1);
    let module = &project.modules[0];
    assert_eq!(module.name, "Module1");
    assert_eq!(module.stream_name, module_stream_name_unicode);
    assert!(module.code.contains("Sub Foo"));

    // V3 content transcript should also be able to locate the module stream.
    let v3 = v3_content_normalized_data(&vba_project_bin).expect("V3ContentNormalizedData");
    assert!(
        v3.windows(b"Sub Foo()".len()).any(|w| w == b"Sub Foo()"),
        "expected v3 transcript to include module source"
    );
}

#[test]
fn vba_project_parse_accepts_spec_dir_stream() {
    let module_source = b"Attribute VB_Name = \"Module1\"\r\nSub Foo()\r\nEnd Sub\r\n";
    let vba_project_bin = build_vba_project_bin_spec(module_source, None);
    let project = VBAProject::parse(&vba_project_bin).expect("parse VBA project");

    assert_eq!(project.name.as_deref(), Some("VBAProject"));
    assert_eq!(project.modules.len(), 1);
    let module = &project.modules[0];
    assert_eq!(module.name, "Module1");
    assert_eq!(module.stream_name, "Module1");
    assert!(module.code.contains("Sub Foo"));
    assert_eq!(
        module.attributes.get("VB_Name").map(String::as_str),
        Some("Module1")
    );
}

#[test]
fn project_normalized_data_v3_dir_records_accepts_spec_dir_stream() {
    // Ensure the v3 dir-record-only transcript helper can scan spec-compliant `VBA/dir` streams
    // that include the fixed-length PROJECTVERSION (0x0009) record layout.
    let module_source = b"Attribute VB_Name = \"Module1\"\r\nSub Foo()\r\nEnd Sub\r\n";
    let vba_project_bin = build_vba_project_bin_spec(module_source, None);

    let normalized = project_normalized_data_v3_dir_records(&vba_project_bin)
        .expect("ProjectNormalizedDataV3 dir-record transcript");

    let mut expected = Vec::new();
    // PROJECTSYSKIND.SysKind
    expected.extend_from_slice(&0x0000_0003u32.to_le_bytes());
    // PROJECTLCID.Lcid
    expected.extend_from_slice(&0x0000_0409u32.to_le_bytes());
    // PROJECTLCIDINVOKE.LcidInvoke
    expected.extend_from_slice(&0x0000_0409u32.to_le_bytes());
    // PROJECTCODEPAGE.CodePage (u16)
    expected.extend_from_slice(&1252u16.to_le_bytes());
    // PROJECTNAME.ProjectName
    expected.extend_from_slice(b"VBAProject");
    // PROJECTHELPCONTEXT.HelpContext
    expected.extend_from_slice(&0u32.to_le_bytes());
    // PROJECTLIBFLAGS.ProjectLibFlags
    expected.extend_from_slice(&0u32.to_le_bytes());
    // PROJECTVERSION: Reserved(u32) || VersionMajor(u32) || VersionMinor(u16)
    expected.extend_from_slice(&0u32.to_le_bytes());
    expected.extend_from_slice(&1u32.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());

    // PROJECTCONSTANTSUNICODE payload bytes ("Answer=42" as UTF-16LE).
    let mut constants_unicode = Vec::new();
    push_utf16le(&mut constants_unicode, "Answer=42");
    expected.extend_from_slice(&constants_unicode);

    // Module group:
    // - MODULENAME bytes ("Module1")
    expected.extend_from_slice(b"Module1");
    // - MODULESTREAMNAMEUNICODE payload bytes ("Module1" as UTF-16LE)
    let mut stream_name_unicode = Vec::new();
    push_utf16le(&mut stream_name_unicode, "Module1");
    expected.extend_from_slice(&stream_name_unicode);
    // - MODULEHELPCONTEXT (u32)
    expected.extend_from_slice(&0u32.to_le_bytes());

    assert_eq!(normalized, expected);
}

#[test]
fn project_normalized_data_accepts_spec_dir_stream_with_fixed_length_projectversion_record() {
    // Ensure the legacy ProjectNormalizedData helper can scan spec-compliant `VBA/dir` streams that
    // use the fixed-length PROJECTVERSION (0x0009) record layout, while still honoring the
    // ProjectProperties token rules from the `PROJECT` stream.
    let module_source = b"Attribute VB_Name = \"Module1\"\r\nSub Foo()\r\nEnd Sub\r\n";
    let vba_project_bin = build_vba_project_bin_spec(module_source, None);

    let normalized = project_normalized_data(&vba_project_bin).expect("ProjectNormalizedData");

    let mut expected = Vec::new();
    // Selected ProjectInformation record data bytes.
    expected.extend_from_slice(&0x0000_0003u32.to_le_bytes()); // PROJECTSYSKIND.SysKind
    expected.extend_from_slice(&0x0000_0409u32.to_le_bytes()); // PROJECTLCID.Lcid
    expected.extend_from_slice(&0x0000_0409u32.to_le_bytes()); // PROJECTLCIDINVOKE.LcidInvoke
    expected.extend_from_slice(&1252u16.to_le_bytes()); // PROJECTCODEPAGE.CodePage
    expected.extend_from_slice(b"VBAProject"); // PROJECTNAME.ProjectName
    expected.extend_from_slice(&0u32.to_le_bytes()); // PROJECTHELPCONTEXT.HelpContext
    expected.extend_from_slice(&0u32.to_le_bytes()); // PROJECTLIBFLAGS.ProjectLibFlags
                                                     // PROJECTVERSION: Reserved(u32) || VersionMajor(u32) || VersionMinor(u16)
    expected.extend_from_slice(&0u32.to_le_bytes());
    expected.extend_from_slice(&1u32.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());

    // PROJECTCONSTANTSUNICODE payload bytes ("Answer=42" as UTF-16LE).
    let mut constants_unicode = Vec::new();
    push_utf16le(&mut constants_unicode, "Answer=42");
    expected.extend_from_slice(&constants_unicode);

    // No designer modules, so FormsNormalizedData is empty.
    // ProjectProperties token bytes from the `PROJECT` stream.
    expected.extend_from_slice(b"NameVBAProjectModuleModule1");

    assert_eq!(normalized, expected);
}

#[test]
fn project_normalized_data_v3_dir_records_accepts_projectversion_reserved_4() {
    // Some producers emit the MS-OVBA fixed-length PROJECTVERSION record as:
    //   Id(u16) || Reserved(u32=4) || VersionMajor(u32) || VersionMinor(u16)
    //
    // Ensure the v3 dir-record-only transcript helper handles it without mis-parsing it as a TLV
    // record with Size=4.
    let project_name = "VBAProject";
    let module_name = "Module1";
    let project_constants = "Answer=42";

    let dir_decompressed = build_dir_decompressed_spec_with_references(
        project_name,
        project_constants,
        module_name,
        4,
        &[],
    );
    let module_source = b"Attribute VB_Name = \"Module1\"\r\nSub Foo()\r\nEnd Sub\r\n";
    let vba_project_bin = build_vba_project_bin_spec_with_dir(&dir_decompressed, module_source, None);

    let normalized = project_normalized_data_v3_dir_records(&vba_project_bin)
        .expect("ProjectNormalizedDataV3 dir-record transcript");

    let mut expected = Vec::new();
    // PROJECTSYSKIND.SysKind
    expected.extend_from_slice(&0x0000_0003u32.to_le_bytes());
    // PROJECTLCID.Lcid
    expected.extend_from_slice(&0x0000_0409u32.to_le_bytes());
    // PROJECTLCIDINVOKE.LcidInvoke
    expected.extend_from_slice(&0x0000_0409u32.to_le_bytes());
    // PROJECTCODEPAGE.CodePage (u16)
    expected.extend_from_slice(&1252u16.to_le_bytes());
    // PROJECTNAME.ProjectName
    expected.extend_from_slice(b"VBAProject");
    // PROJECTHELPCONTEXT.HelpContext
    expected.extend_from_slice(&0u32.to_le_bytes());
    // PROJECTLIBFLAGS.ProjectLibFlags
    expected.extend_from_slice(&0u32.to_le_bytes());
    // PROJECTVERSION: Reserved(u32) || VersionMajor(u32) || VersionMinor(u16)
    expected.extend_from_slice(&4u32.to_le_bytes());
    expected.extend_from_slice(&1u32.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());

    // PROJECTCONSTANTSUNICODE payload bytes ("Answer=42" as UTF-16LE).
    let mut constants_unicode = Vec::new();
    push_utf16le(&mut constants_unicode, "Answer=42");
    expected.extend_from_slice(&constants_unicode);

    // Module group:
    // - MODULENAME bytes ("Module1")
    expected.extend_from_slice(b"Module1");
    // - MODULESTREAMNAMEUNICODE payload bytes ("Module1" as UTF-16LE)
    let mut stream_name_unicode = Vec::new();
    push_utf16le(&mut stream_name_unicode, "Module1");
    expected.extend_from_slice(&stream_name_unicode);
    // - MODULEHELPCONTEXT (u32)
    expected.extend_from_slice(&0u32.to_le_bytes());

    assert_eq!(normalized, expected);
}

#[test]
fn project_normalized_data_accepts_spec_dir_stream_with_fixed_length_projectversion_reserved_4() {
    // Same as the previous regression, but uses the MS-OVBA mandated Reserved=4 value for the
    // fixed-length PROJECTVERSION record.
    let module_source = b"Attribute VB_Name = \"Module1\"\r\nSub Foo()\r\nEnd Sub\r\n";

    let project_name = "VBAProject";
    let module_name = "Module1";
    let project_constants = "Answer=42";
    let dir_decompressed =
        build_dir_decompressed_spec_with_references(project_name, project_constants, module_name, 4, &[]);
    let vba_project_bin =
        build_vba_project_bin_spec_with_dir(&dir_decompressed, module_source, None);

    let normalized = project_normalized_data(&vba_project_bin).expect("ProjectNormalizedData");

    let mut expected = Vec::new();
    // Selected ProjectInformation record data bytes.
    expected.extend_from_slice(&0x0000_0003u32.to_le_bytes()); // PROJECTSYSKIND.SysKind
    expected.extend_from_slice(&0x0000_0409u32.to_le_bytes()); // PROJECTLCID.Lcid
    expected.extend_from_slice(&0x0000_0409u32.to_le_bytes()); // PROJECTLCIDINVOKE.LcidInvoke
    expected.extend_from_slice(&1252u16.to_le_bytes()); // PROJECTCODEPAGE.CodePage
    expected.extend_from_slice(b"VBAProject"); // PROJECTNAME.ProjectName
    expected.extend_from_slice(&0u32.to_le_bytes()); // PROJECTHELPCONTEXT.HelpContext
    expected.extend_from_slice(&0u32.to_le_bytes()); // PROJECTLIBFLAGS.ProjectLibFlags
    // PROJECTVERSION: Reserved(u32) || VersionMajor(u32) || VersionMinor(u16)
    expected.extend_from_slice(&4u32.to_le_bytes());
    expected.extend_from_slice(&1u32.to_le_bytes());
    expected.extend_from_slice(&0u16.to_le_bytes());
    // PROJECTCONSTANTSUNICODE payload bytes ("Answer=42" as UTF-16LE).
    let mut constants_unicode = Vec::new();
    push_utf16le(&mut constants_unicode, "Answer=42");
    expected.extend_from_slice(&constants_unicode);

    // No designer modules, so FormsNormalizedData is empty.
    // ProjectProperties token bytes from the `PROJECT` stream.
    expected.extend_from_slice(b"NameVBAProjectModuleModule1");

    assert_eq!(normalized, expected);
}

#[test]
fn content_normalized_data_parses_spec_dir_stream_with_reference_records() {
    let project_name = "VBAProject";
    let module_name = "Module1";
    let project_constants = "Answer=42";

    let module_source = b"Attribute VB_Name = \"Module1\"\r\nSub Foo()\r\nEnd Sub\r\n";

    let references = {
        let mut out = Vec::new();

        // Excluded record: REFERENCENAME (0x0016). Must not affect output.
        push_referencename_record(&mut out, b"EXCLUDED_REF_NAME");

        // Included record: REFERENCEREGISTERED (0x000D).
        push_u16(&mut out, 0x000D);
        push_u32(&mut out, 5);
        out.extend_from_slice(b"{REG}");

        // Included record: REFERENCECONTROL (0x002F), plus embedded REFERENCEEXTENDED (0x0030).
        //
        // Control record data:
        // - u32 len + bytes (LibidTwiddled)
        // - Reserved1 (u32)
        // - Reserved2 (u16)
        let libid_twiddled = b"CtrlLib";
        let reserved1: u32 = 1; // 01 00 00 00
        let reserved2: u16 = 0;
        let mut control_data = Vec::new();
        control_data.extend_from_slice(&(libid_twiddled.len() as u32).to_le_bytes());
        control_data.extend_from_slice(libid_twiddled);
        control_data.extend_from_slice(&reserved1.to_le_bytes());
        control_data.extend_from_slice(&reserved2.to_le_bytes());

        // Optional NameRecordExtended (excluded from v1/v2 transcript).
        push_u16(&mut out, 0x002F);
        push_u32(&mut out, control_data.len() as u32);
        out.extend_from_slice(&control_data);
        push_referencename_record(&mut out, b"CONTROL_NAME_EXT");

        // Embedded REFERENCEEXTENDED (0x0030) bytes are included verbatim.
        push_u16(&mut out, 0x0030);
        push_u32(&mut out, 3);
        out.extend_from_slice(b"EXT");

        // Included record: REFERENCEPROJECT (0x000E).
        //
        // Choose major=1 so the copy-until-NUL logic stops after copying 0x01.
        let libid_absolute = b"ProjLib";
        let libid_relative = b"";
        let major: u32 = 1; // 01 00 00 00
        let minor: u16 = 0;
        let trailing = b"TRAIL";
        let size_total =
            4 + libid_absolute.len() + 4 + libid_relative.len() + 4 + 2 + trailing.len();
        push_u16(&mut out, 0x000E);
        push_u32(&mut out, size_total as u32);
        push_u32(&mut out, libid_absolute.len() as u32);
        out.extend_from_slice(libid_absolute);
        push_u32(&mut out, libid_relative.len() as u32);
        out.extend_from_slice(libid_relative);
        push_u32(&mut out, major);
        push_u16(&mut out, minor);
        out.extend_from_slice(trailing);

        // Included record: REFERENCEORIGINAL (0x0033), with an embedded REFERENCECONTROL that must
        // be skipped (it is part of the REFERENCEORIGINAL structure).
        let libid_original = b"OrigLib";
        push_u16(&mut out, 0x0033);
        push_u32(&mut out, libid_original.len() as u32);
        out.extend_from_slice(libid_original);

        // Embedded REFERENCECONTROL (0x002F) + REFERENCEEXTENDED (0x0030) that must be skipped.
        let nested_libid_twiddled = b"SHOULD_NOT_APPEAR";
        let nested_reserved1: u32 = 1;
        let nested_reserved2: u16 = 0;
        let mut nested_control_data = Vec::new();
        nested_control_data.extend_from_slice(&(nested_libid_twiddled.len() as u32).to_le_bytes());
        nested_control_data.extend_from_slice(nested_libid_twiddled);
        nested_control_data.extend_from_slice(&nested_reserved1.to_le_bytes());
        nested_control_data.extend_from_slice(&nested_reserved2.to_le_bytes());
        push_u16(&mut out, 0x002F);
        push_u32(&mut out, nested_control_data.len() as u32);
        out.extend_from_slice(&nested_control_data);

        push_u16(&mut out, 0x0030);
        let nested_extended = b"SKIP_EXTENDED";
        push_u32(&mut out, nested_extended.len() as u32);
        out.extend_from_slice(nested_extended);

        out
    };

    let dir_decompressed = build_dir_decompressed_spec_with_references(
        project_name,
        project_constants,
        module_name,
        0,
        &references,
    );
    let vba_project_bin =
        build_vba_project_bin_spec_with_dir(&dir_decompressed, module_source, None);
    let normalized = content_normalized_data(&vba_project_bin).expect("ContentNormalizedData");

    let expected_module_normalized = b"Sub Foo()\r\nEnd Sub\r\n".as_slice();
    let expected = [
        project_name.as_bytes(),
        project_constants.as_bytes(),
        b"{REG}".as_slice(),
        b"CtrlLib\x01".as_slice(),
        b"EXT".as_slice(),
        b"ProjLib\x01".as_slice(),
        b"OrigLib".as_slice(),
        expected_module_normalized,
    ]
    .concat();

    assert_eq!(normalized, expected);
    assert!(
        !normalized
            .windows(b"SHOULD_NOT_APPEAR".len())
            .any(|w| w == b"SHOULD_NOT_APPEAR"),
        "embedded REFERENCECONTROL inside REFERENCEORIGINAL must not contribute"
    );
    assert!(
        !normalized
            .windows(b"SKIP_EXTENDED".len())
            .any(|w| w == b"SKIP_EXTENDED"),
        "embedded REFERENCEEXTENDED inside REFERENCEORIGINAL must not contribute"
    );
    assert!(
        !normalized
            .windows(b"EXCLUDED_REF_NAME".len())
            .any(|w| w == b"EXCLUDED_REF_NAME"),
        "REFERENCENAME (0x0016) must not contribute to ContentNormalizedData"
    );
    assert!(
        !normalized
            .windows(b"CONTROL_NAME_EXT".len())
            .any(|w| w == b"CONTROL_NAME_EXT"),
        "NameRecordExtended inside REFERENCECONTROL must not contribute to ContentNormalizedData"
    );
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
