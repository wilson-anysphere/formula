//! `formula-vba` provides lightweight parsing utilities for the VBA project
//! embedded in macro-enabled Excel workbooks (`vbaProject.bin`).
//!
//! Current scope (per `docs/08-macro-compatibility.md`):
//! - L1: Preserve `vbaProject.bin` bytes (handled by the XLSX layer).
//! - L2: Parse enough of `vbaProject.bin` to enumerate modules and show source.
//! - Inspect and (best-effort) cryptographically verify VBA project digital signatures
//!   (PKCS#7/CMS) for desktop Trust Center policy.

mod authenticode;
mod compression;
mod contents_hash;
mod dir;
mod normalized_data;
mod offcrypto;
mod offcrypto_rc4;
mod ole;
mod project_normalized_data;
mod project_digest;
mod rc4_cryptoapi;
mod signature;

pub use authenticode::{
    extract_vba_signature_signed_digest, VbaSignatureSignedDigestError, VbaSignedDigest,
};
pub use compression::{compress_container, decompress_container, CompressionError};
pub use contents_hash::{
    agile_content_hash_md5, content_hash_md5, content_normalized_data, contents_hash_v3,
    project_normalized_data, project_normalized_data_v3_transcript, v3_content_normalized_data,
};
pub use dir::{DirParseError, DirStream, ModuleRecord, ModuleType};
pub use normalized_data::forms_normalized_data;
pub use ole::{OleError, OleFile};
pub use project_digest::{compute_vba_project_digest, compute_vba_project_digest_v3, DigestAlg};
pub use project_normalized_data::{
    project_normalized_data_v3, project_normalized_data_v3_dir_records,
};
pub use signature::{
    extract_signer_certificate_info, list_vba_digital_signatures,
    parse_and_verify_vba_signature_blob, parse_vba_digital_signature, verify_vba_digital_signature,
    verify_vba_digital_signature_bound, verify_vba_digital_signature_with_project,
    verify_vba_digital_signature_with_trust, verify_vba_project_signature_binding,
    verify_vba_signature_binding, verify_vba_signature_binding_with_stream_path,
    verify_vba_signature_blob, verify_vba_signature_certificate_trust, SignatureError,
    VbaCertificateTrust, VbaDigitalSignature, VbaDigitalSignatureBound, VbaDigitalSignatureStream,
    VbaDigitalSignatureTrusted, VbaProjectBindingVerification, VbaProjectDigestDebugInfo,
    VbaSignatureBinding, VbaSignatureBlobInfo, VbaSignatureStreamKind, VbaSignatureTrustOptions,
    VbaSignatureVerification, VbaSignerCertificateInfo,
};

use std::collections::BTreeMap;

use encoding_rs::{
    Encoding, BIG5, EUC_KR, GBK, SHIFT_JIS, UTF_16LE, UTF_8, WINDOWS_1250, WINDOWS_1251,
    WINDOWS_1252, WINDOWS_1253, WINDOWS_1254, WINDOWS_1255, WINDOWS_1256, WINDOWS_1257,
    WINDOWS_1258, WINDOWS_874,
};
use thiserror::Error;

/// Parsed representation of a VBA project. This model is intentionally minimal
/// and geared towards UI display (module tree + code viewer).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VBAProject {
    pub name: Option<String>,
    pub constants: Option<String>,
    pub references: Vec<VBAReference>,
    pub modules: Vec<VBAModule>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VBAModule {
    pub name: String,
    pub stream_name: String,
    pub module_type: ModuleType,
    pub code: String,
    pub attributes: BTreeMap<String, String>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VBAReference {
    pub name: Option<String>,
    pub guid: Option<String>,
    pub major: Option<u16>,
    pub minor: Option<u16>,
    pub path: Option<String>,
    pub raw: String,
}

#[derive(Debug, Error)]
pub enum ParseError {
    #[error("OLE error: {0}")]
    Ole(#[from] OleError),
    #[error("VBA compression error: {0}")]
    Compression(#[from] CompressionError),
    #[error("dir stream parse error: {0}")]
    Dir(#[from] DirParseError),
    #[error("missing required stream {0}")]
    MissingStream(&'static str),
    #[error("missing required storage {0}")]
    MissingStorage(String),
}

impl VBAProject {
    /// Parse a `vbaProject.bin` OLE compound file into a [`VBAProject`].
    ///
    /// The goal is to be permissive: we try to recover modules even if some
    /// optional metadata is missing.
    pub fn parse(vba_project_bin: &[u8]) -> Result<Self, ParseError> {
        let mut ole = OleFile::open(vba_project_bin)?;

        let project_stream_bytes = ole.read_stream_opt("PROJECT")?;
        let mut references = Vec::new();
        let mut name_from_project_stream = None;
        let dir_bytes = ole
            .read_stream_opt("VBA/dir")?
            .ok_or(ParseError::MissingStream("VBA/dir"))?;
        let dir_decompressed = decompress_container(&dir_bytes)?;

        let encoding = project_stream_bytes
            .as_deref()
            .and_then(detect_project_codepage)
            .or_else(|| {
                DirStream::detect_codepage(&dir_decompressed)
                    .map(|cp| cp as u32)
                    .map(encoding_for_codepage)
            })
            .unwrap_or(WINDOWS_1252);

        if let Some(project_stream_bytes) = project_stream_bytes.as_deref() {
            let text = decode_with_encoding(project_stream_bytes, encoding);
            for line in split_crlf_lines(&text) {
                if let Some(rest) = line.strip_prefix("Name=") {
                    name_from_project_stream = Some(rest.trim_matches('"').to_owned());
                } else if let Some(rest) = line.strip_prefix("Reference=") {
                    references.push(parse_reference(rest));
                }
            }
        }

        let dir_stream = DirStream::parse_with_encoding(&dir_decompressed, encoding)?;

        let mut modules = Vec::new();
        for module in &dir_stream.modules {
            let stream_path = format!("VBA/{}", module.stream_name);
            let module_stream = ole
                .read_stream_opt(&stream_path)?
                .ok_or(ParseError::MissingStream("module stream"))?;

            let text_offset = module
                .text_offset
                .unwrap_or_else(|| guess_text_offset(&module_stream));
            let text_offset = text_offset.min(module_stream.len());
            let source_container = &module_stream[text_offset..];
            let source_bytes = decompress_container(source_container)?;
            let code = decode_with_encoding(&source_bytes, encoding);
            let attributes = parse_attributes(&code);

            modules.push(VBAModule {
                name: module.name.clone(),
                stream_name: module.stream_name.clone(),
                module_type: module.module_type,
                code,
                attributes,
            });
        }

        Ok(Self {
            name: dir_stream.project_name.clone().or(name_from_project_stream),
            constants: dir_stream.constants.clone(),
            references,
            modules,
        })
    }
}

fn decode_with_encoding(bytes: &[u8], encoding: &'static Encoding) -> String {
    // Module source is commonly stored as MBCS in the project codepage. If it
    // looks like UTF-16LE, decode as UTF-16LE instead (some producers emit it).
    if bytes.len() >= 2 && bytes.len().is_multiple_of(2) {
        // Same heuristic as `dir` strings: if many high bytes are NUL, treat as UTF-16LE.
        let total = bytes.len() / 2;
        let nul_high = bytes.iter().skip(1).step_by(2).filter(|&&b| b == 0).count();
        // Use a ceiling half threshold. For very short inputs (e.g. 2 bytes), `total / 2` is 0 and
        // would incorrectly classify any 2-byte MBCS string as UTF-16LE.
        if nul_high * 2 >= total {
            let (cow, _) = UTF_16LE.decode_without_bom_handling(bytes);
            return cow.into_owned();
        }
    }

    let (cow, _, _) = encoding.decode(bytes);
    cow.into_owned()
}

fn detect_project_codepage(project_stream_bytes: &[u8]) -> Option<&'static Encoding> {
    // The `PROJECT` stream is plain text, but it may be encoded in the project's codepage. The
    // `CodePage=...` directive itself is ASCII, so we can parse it directly from bytes without
    // decoding the entire stream.
    //
    // Be permissive: accept case-insensitive `CodePage` with optional whitespace around `=`.
    for mut line in project_stream_bytes.split(|&b| b == b'\n' || b == b'\r') {
        if line.is_empty() {
            continue;
        }

        // Some producers may include a UTF-8 BOM at the start of the stream.
        if line.starts_with(&[0xEF, 0xBB, 0xBF]) {
            line = &line[3..];
        }

        // Trim leading ASCII whitespace.
        let mut start = 0usize;
        while start < line.len() && matches!(line[start], b' ' | b'\t') {
            start += 1;
        }
        let mut line = &line[start..];

        // Trim trailing ASCII whitespace.
        let mut end = line.len();
        while end > 0 && matches!(line[end - 1], b' ' | b'\t') {
            end -= 1;
        }
        line = &line[..end];

        // MS-OVBA `ProjectProperties` ends at the first section header like `[Host Extender Info]`
        // or `[Workspace]`. The CodePage directive is a ProjectProperty, so stop scanning once we've
        // reached a section header.
        if line.starts_with(b"[") && line.ends_with(b"]") {
            break;
        }

        // Parse `CodePage` key.
        const KEY: &[u8] = b"CodePage";
        let Some(prefix) = line.get(..KEY.len()) else {
            continue;
        };
        if !prefix.eq_ignore_ascii_case(KEY) {
            continue;
        }
        line = &line[KEY.len()..];

        // Optional whitespace then '='.
        let mut start = 0usize;
        while start < line.len() && matches!(line[start], b' ' | b'\t') {
            start += 1;
        }
        line = &line[start..];
        if !line.starts_with(b"=") {
            continue;
        }
        line = &line[1..];

        // Optional whitespace then digits.
        let mut start = 0usize;
        while start < line.len() && matches!(line[start], b' ' | b'\t') {
            start += 1;
        }
        line = &line[start..];

        let mut end = 0usize;
        while end < line.len() && line[end].is_ascii_digit() {
            end += 1;
        }
        if end == 0 {
            continue;
        }
        let digits = &line[..end];
        let Ok(n) = std::str::from_utf8(digits).ok()?.parse::<u32>() else {
            // Be robust against malformed/overflowing values; continue scanning other lines.
            continue;
        };
        return Some(encoding_for_codepage(n));
    }

    None
}

fn split_crlf_lines(text: &str) -> impl Iterator<Item = &str> {
    text.split(['\n', '\r'])
        .map(str::trim)
        .filter(|line| !line.is_empty())
}

fn encoding_for_codepage(codepage: u32) -> &'static Encoding {
    match codepage {
        874 => WINDOWS_874,
        932 => SHIFT_JIS,
        936 => GBK,
        949 => EUC_KR,
        950 => BIG5,
        1250 => WINDOWS_1250,
        1251 => WINDOWS_1251,
        1252 => WINDOWS_1252,
        1253 => WINDOWS_1253,
        1254 => WINDOWS_1254,
        1255 => WINDOWS_1255,
        1256 => WINDOWS_1256,
        1257 => WINDOWS_1257,
        1258 => WINDOWS_1258,
        65001 => UTF_8,
        _ => WINDOWS_1252,
    }
}

fn parse_reference(raw: &str) -> VBAReference {
    // References in the `PROJECT` stream are stored as a single string with
    // `#`-delimited fields. We parse a few common patterns (type library refs)
    // and keep the original raw value as a fallback.
    //
    // Example (typelib):
    //   *\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\SysWOW64\\stdole2.tlb#OLE Automation
    let mut reference = VBAReference {
        name: None,
        guid: None,
        major: None,
        minor: None,
        path: None,
        raw: raw.to_owned(),
    };

    let parts: Vec<&str> = raw.split('#').collect();
    if parts.is_empty() {
        return reference;
    }

    // Extract GUID from the first part if present.
    if let Some(first) = parts.first() {
        if let Some(start) = first.find('{') {
            if let Some(begin) = start.checked_add(1) {
                if let Some(after_start) = first.get(begin..) {
                    if let Some(end_rel) = after_start.find('}') {
                        if let Some(end) = begin.checked_add(end_rel) {
                            if let Some(guid) = first.get(begin..end) {
                                if !guid.is_empty() {
                                    reference.guid = Some(guid.to_owned());
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Version major.minor in the second part.
    if let Some(version) = parts.get(1) {
        if let Some((major, minor)) = version.split_once('.') {
            reference.major = major.parse::<u16>().ok();
            reference.minor = minor.parse::<u16>().ok();
        }
    }

    // Path is commonly the 4th field.
    if let Some(path) = parts.get(3) {
        if !path.is_empty() {
            reference.path = Some((*path).to_owned());
        }
    }

    // Human-readable name/description is typically the last field.
    if let Some(name) = parts.last() {
        if !name.is_empty() {
            reference.name = Some((*name).to_owned());
        }
    }

    reference
}

fn parse_attributes(code: &str) -> BTreeMap<String, String> {
    let mut attrs = BTreeMap::new();
    for line in code.lines() {
        let line = line.trim_end_matches('\r').trim();
        let Some(rest) = line.strip_prefix("Attribute ") else {
            continue;
        };

        let Some((key, value)) = rest.split_once('=') else {
            continue;
        };

        let key = key.trim().to_owned();
        let value = value.trim();
        let value = value
            .strip_prefix('"')
            .and_then(|v| v.strip_suffix('"'))
            .unwrap_or(value)
            .to_owned();
        attrs.insert(key, value);
    }
    attrs
}

/// Try to find the start of the compressed source container in a module stream
/// when the `dir` stream doesn't give us a text offset.
///
/// This is a best-effort heuristic: scan for the first 0x01 byte that looks
/// like an MS-OVBA compressed container signature.
fn guess_text_offset(module_stream: &[u8]) -> usize {
    // CompressedContainer starts with 0x01, followed by a chunk header whose signature bits
    // must be 0b011 (MS-OVBA 2.4.1.3.5). This avoids false positives where 0x01 appears in the
    // module stream header.
    for idx in 0..module_stream.len().saturating_sub(3) {
        if module_stream[idx] != 0x01 {
            continue;
        }
        let Some(bytes) = module_stream.get(idx..).and_then(|rest| rest.get(..3)) else {
            continue;
        };
        let header = u16::from_le_bytes([bytes[1], bytes[2]]);
        let signature_bits = (header & 0x7000) >> 12;
        if signature_bits == 0b011 {
            // Best-effort validation: module streams can contain header bytes before the compressed
            // source container, and those bytes can occasionally look like a container signature.
            // Attempt decompression; if it fails, keep scanning for a later candidate.
            if decompress_container(&module_stream[idx..]).is_ok() {
                return idx;
            }
        }
    }
    0
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Read;

    fn load_fixture_vba_bin() -> Vec<u8> {
        let fixture_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../fixtures/xlsx/macros/basic.xlsm"
        );
        let data = std::fs::read(fixture_path).expect("fixture xlsm exists");
        let reader = std::io::Cursor::new(data);
        let mut zip = zip::ZipArchive::new(reader).expect("valid zip");
        let mut file = zip
            .by_name("xl/vbaProject.bin")
            .expect("vbaProject.bin in fixture");
        let mut buf = Vec::new();
        file.read_to_end(&mut buf).unwrap();
        buf
    }

    #[test]
    fn parses_modules_and_code_from_fixture() {
        let vba_bin = load_fixture_vba_bin();
        let project = VBAProject::parse(&vba_bin).expect("parse VBA project");
        assert_eq!(project.name.as_deref(), Some("VBAProject"));
        assert!(!project.modules.is_empty());

        let module = project
            .modules
            .iter()
            .find(|m| m.name == "Module1")
            .expect("Module1 present");
        assert_eq!(
            module.attributes.get("VB_Name").map(String::as_str),
            Some("Module1")
        );
        assert!(module.code.contains("Sub Hello"));
        assert!(module.code.contains("MsgBox"));
    }

    #[test]
    fn decompresses_copy_tokens() {
        // Compressed container that expands "ABCABCABC" using a copy token.
        //
        // We build one compressed chunk:
        // - First 3 literals: A B C
        // - Then a copy token: offset=3, length=6 -> repeats ABCABC
        //
        // Encoding details follow MS-OVBA:
        // copy_token_bit_count at output_len=3 is 4 (minimum).
        // length bits = 12, offset bits = 4.
        // token layout: (offset-1)<<12 | (length-3)
        // offset-1 = 2, length-3 = 3 => 0x2000 | 0x0003 = 0x2003
        let mut chunk_data = Vec::new();
        // Flag byte: bits 0-2 literal, bit3 copy token.
        chunk_data.push(0b0000_1000);
        chunk_data.extend_from_slice(b"ABC");
        chunk_data.extend_from_slice(&0x2003u16.to_le_bytes());

        // Container: signature + chunk header + chunk data
        // header:
        // - size field = chunk_data_len - 1
        // - signature bits = 0b011 at bits 12..14 => 0x3000
        // - compressed flag at bit 15 => 0x8000
        let size_field = (chunk_data.len() - 1) as u16;
        let header = 0xB000u16 | size_field;
        let mut container = Vec::new();
        container.push(0x01);
        container.extend_from_slice(&header.to_le_bytes());
        container.extend_from_slice(&chunk_data);

        let out = decompress_container(&container).expect("decompress");
        assert_eq!(&out, b"ABCABCABC");
    }

    #[test]
    fn compresses_and_decompresses_roundtrip() {
        let data = b"ABCABCABCABCABCABCABCABCABCABC";
        let compressed = compress_container(data);
        // The first chunk should be marked as compressed for this highly repetitive input.
        let header = u16::from_le_bytes([compressed[1], compressed[2]]);
        assert_ne!(header & 0x8000, 0);

        let out = decompress_container(&compressed).expect("decompress");
        assert_eq!(&out, data);
    }

    #[test]
    fn detects_project_codepage_with_whitespace_and_case() {
        let bytes = b"Name=\"VBAProject\"\r\ncodepage = 1251\r\n";
        let enc = detect_project_codepage(bytes).expect("detect CodePage");
        assert_eq!(enc, WINDOWS_1251);
    }

    #[test]
    fn detects_project_codepage_with_cr_line_endings() {
        // Some producers might use CR-only line endings. We should still find the CodePage line.
        let bytes = b"ID={123}\rcodepage=1251\rName=VBAProject\r";
        let enc = detect_project_codepage(bytes).expect("detect CodePage");
        assert_eq!(enc, WINDOWS_1251);
    }

    #[test]
    fn detects_project_codepage_with_utf8_bom() {
        let bytes = b"\xEF\xBB\xBFCodePage=1251\r\n";
        let enc = detect_project_codepage(bytes).expect("detect CodePage");
        assert_eq!(enc, WINDOWS_1251);
    }

    #[test]
    fn detects_project_codepage_ignores_overflowing_values() {
        // A malicious/invalid CodePage line shouldn't prevent us from finding a later valid one.
        let bytes = b"CodePage=99999999999999999999999999999999\r\nCodePage=1251\r\n";
        let enc = detect_project_codepage(bytes).expect("detect CodePage");
        assert_eq!(enc, WINDOWS_1251);
    }

    #[test]
    fn detects_project_codepage_stops_at_section_headers() {
        // MS-OVBA ProjectProperties ends at the first section header like `[Host Extender Info]` or
        // `[Workspace]`. CodePage is a ProjectProperty and should not be read from later sections.
        let bytes = b"CodePage=1252\r\n[Workspace]\r\nCodePage=1251\r\n";
        let enc = detect_project_codepage(bytes).expect("detect CodePage");
        assert_eq!(enc, WINDOWS_1252);
    }

    #[test]
    fn parses_project_stream_with_cr_line_endings() {
        use std::io::{Cursor, Write};

        let module_code = "Sub Hello()\r\nEnd Sub\r\n";
        let module_container = compress_container(module_code.as_bytes());

        let dir_decompressed = {
            let mut out = Vec::new();
            // Omit PROJECTNAME (0x0004) so project name comes from PROJECT stream parsing.

            // MODULENAME / MODULESTREAMNAME
            push_record(&mut out, 0x0019, b"Module1");
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(b"Module1");
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name);

            // MODULETYPE (standard) + MODULETEXTOFFSET (0)
            push_record(&mut out, 0x0021, &0u16.to_le_bytes());
            push_record(&mut out, 0x0031, &0u32.to_le_bytes());
            out
        };
        let dir_container = compress_container(&dir_decompressed);

        // CR-only PROJECT stream lines.
        let project_stream = b"CodePage=1252\rName=\"MyProj\"\rModule=Module1\r";

        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        {
            let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
            s.write_all(project_stream).expect("write PROJECT");
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

        let vba_bin = ole.into_inner().into_inner();
        let project = VBAProject::parse(&vba_bin).expect("parse");
        assert_eq!(project.name.as_deref(), Some("MyProj"));

        let module = project.modules.iter().find(|m| m.name == "Module1").unwrap();
        assert!(module.code.contains("Sub Hello"));
    }

    #[test]
    fn respects_dir_codepage_for_module_source() {
        use std::io::{Cursor, Write};

        let comment = "привет"; // "hello" in Russian
        let code_utf8 = format!("Sub Hello()\r\n'{}\r\nEnd Sub\r\n", comment);
        let (code_bytes, _, _) = WINDOWS_1251.encode(&code_utf8);

        let module_container = compress_container(code_bytes.as_ref());

        let dir_decompressed = {
            let mut out = Vec::new();
            // PROJECTCODEPAGE (u16 LE)
            push_record(&mut out, 0x0003, &1251u16.to_le_bytes());
            // PROJECTNAME
            let (proj_name_bytes, _, _) = WINDOWS_1251.encode("Проект");
            push_record(&mut out, 0x0004, proj_name_bytes.as_ref());

            // MODULENAME
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
            out
        };
        let dir_container = compress_container(&dir_decompressed);

        let project_stream_text = "Name=\"VBAProject\"\r\nModule=Module1\r\n";
        let (project_stream_bytes, _, _) = WINDOWS_1251.encode(project_stream_text);

        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        {
            let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
            s.write_all(project_stream_bytes.as_ref())
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

        let vba_bin = ole.into_inner().into_inner();
        let project = VBAProject::parse(&vba_bin).expect("parse");
        assert_eq!(project.name.as_deref(), Some("Проект"));

        let module = project
            .modules
            .iter()
            .find(|m| m.name == "Module1")
            .unwrap();
        assert!(module.code.contains(comment));
    }

    #[test]
    fn parses_module_with_two_byte_ascii_name() {
        use std::io::{Cursor, Write};

        // Regression test for the UTF-16LE detection heuristic: a 2-byte ASCII/MBCS name like "AB"
        // must not be treated as UTF-16LE, otherwise module stream lookup breaks (e.g. "AB" would
        // decode to a different single UTF-16 code unit).
        let module_code = "Sub Hello()\r\nEnd Sub\r\n";
        let module_container = compress_container(module_code.as_bytes());

        let dir_decompressed = {
            let mut out = Vec::new();
            // PROJECTCODEPAGE (u16 LE)
            push_record(&mut out, 0x0003, &1252u16.to_le_bytes());

            // MODULENAME / MODULESTREAMNAME
            push_record(&mut out, 0x0019, b"AB");
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(b"AB");
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name);

            // MODULETYPE (standard) + MODULETEXTOFFSET (0)
            push_record(&mut out, 0x0021, &0u16.to_le_bytes());
            push_record(&mut out, 0x0031, &0u32.to_le_bytes());
            out
        };
        let dir_container = compress_container(&dir_decompressed);

        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        {
            let mut s = ole.create_stream("PROJECT").expect("PROJECT stream");
            s.write_all(b"CodePage=1252\r\n").expect("write PROJECT");
        }
        ole.create_storage("VBA").expect("VBA storage");
        {
            let mut s = ole.create_stream("VBA/dir").expect("dir stream");
            s.write_all(&dir_container).expect("write dir");
        }
        {
            let mut s = ole.create_stream("VBA/AB").expect("module stream");
            s.write_all(&module_container).expect("write module");
        }

        let vba_bin = ole.into_inner().into_inner();
        let project = VBAProject::parse(&vba_bin).expect("parse");
        let module = project.modules.iter().find(|m| m.name == "AB").unwrap();
        assert_eq!(module.stream_name, "AB");
        assert!(module.code.contains("Sub Hello"));
    }

    #[test]
    fn parses_module_using_modulestreamnameunicode_record() {
        use std::io::{Cursor, Write};

        // Some `VBA/dir` streams include a separate MODULESTREAMNAMEUNICODE (0x0032) record that
        // provides the UTF-16LE module stream name. Ensure we honor it for OLE lookup.
        let module_stream_name = "Модуль1";
        let mut stream_name_unicode = Vec::new();
        for unit in module_stream_name.encode_utf16() {
            stream_name_unicode.extend_from_slice(&unit.to_le_bytes());
        }

        let module_code = "Sub Hello()\r\nEnd Sub\r\n";
        let module_container = compress_container(module_code.as_bytes());

        let dir_decompressed = {
            let mut out = Vec::new();

            // PROJECTCODEPAGE (u16 LE) so the project encoding resolves deterministically.
            push_record(&mut out, 0x0003, &1251u16.to_le_bytes());

            // MODULENAME is the module identifier; keep it ASCII.
            push_record(&mut out, 0x0019, b"Module1");

            // Deliberately wrong MODULESTREAMNAME to prove we prefer MODULESTREAMNAMEUNICODE.
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(b"WrongName");
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name);

            // Correct Unicode stream name.
            push_record(&mut out, 0x0032, &stream_name_unicode);

            // MODULETYPE (standard) + MODULETEXTOFFSET (0).
            push_record(&mut out, 0x0021, &0u16.to_le_bytes());
            push_record(&mut out, 0x0031, &0u32.to_le_bytes());
            out
        };
        let dir_container = compress_container(&dir_decompressed);

        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        ole.create_storage("VBA").expect("VBA storage");
        {
            let mut s = ole.create_stream("VBA/dir").expect("dir stream");
            s.write_all(&dir_container).expect("write dir");
        }
        {
            let path = format!("VBA/{module_stream_name}");
            let mut s = ole.create_stream(&path).expect("module stream");
            s.write_all(&module_container).expect("write module");
        }

        let vba_bin = ole.into_inner().into_inner();
        let project = VBAProject::parse(&vba_bin).expect("parse");
        let module = project
            .modules
            .iter()
            .find(|m| m.name == "Module1")
            .expect("Module1 present");
        assert_eq!(module.stream_name, module_stream_name);
        assert!(module.code.contains("Sub Hello"));
    }

    #[test]
    fn parses_module_using_unicode_only_name_and_stream_records() {
        use std::io::{Cursor, Write};

        // Some `VBA/dir` encodings omit MODULENAME (0x0019) and emit only MODULENAMEUNICODE (0x0047),
        // along with MODULESTREAMNAMEUNICODE (0x0032). Ensure we can still enumerate and open the
        // module stream.
        let module_name = "Модуль1"; // non-ASCII

        let mut name_unicode = Vec::new();
        for unit in module_name.encode_utf16() {
            name_unicode.extend_from_slice(&unit.to_le_bytes());
        }

        let module_code = "Sub Hello()\r\nEnd Sub\r\n";
        let module_container = compress_container(module_code.as_bytes());

        let dir_decompressed = {
            let mut out = Vec::new();

            // PROJECTCODEPAGE (u16 LE): present in most real projects; not required for Unicode
            // names, but keeps the overall encoding resolution deterministic.
            push_record(&mut out, 0x0003, &1251u16.to_le_bytes());

            // Unicode-only module identifier + stream name.
            push_record(&mut out, 0x0047, &name_unicode); // MODULENAMEUNICODE
            push_record(&mut out, 0x0032, &name_unicode); // MODULESTREAMNAMEUNICODE

            // MODULETYPE (standard) + MODULETEXTOFFSET (0).
            push_record(&mut out, 0x0021, &0u16.to_le_bytes());
            push_record(&mut out, 0x0031, &0u32.to_le_bytes());
            out
        };
        let dir_container = compress_container(&dir_decompressed);

        let cursor = Cursor::new(Vec::new());
        let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
        ole.create_storage("VBA").expect("VBA storage");
        {
            let mut s = ole.create_stream("VBA/dir").expect("dir stream");
            s.write_all(&dir_container).expect("write dir");
        }
        {
            let path = format!("VBA/{module_name}");
            let mut s = ole.create_stream(&path).expect("module stream");
            s.write_all(&module_container).expect("write module");
        }

        let vba_bin = ole.into_inner().into_inner();
        let project = VBAProject::parse(&vba_bin).expect("parse");
        let module = project
            .modules
            .iter()
            .find(|m| m.name == module_name)
            .expect("module present");
        assert_eq!(module.stream_name, module_name);
        assert!(module.code.contains("Sub Hello"));
    }

    #[test]
    fn parses_module_without_text_offset_using_signature_scan() {
        use std::io::{Cursor, Write};

        let module_code = "Sub Hello()\r\nEnd Sub\r\n";
        let container = compress_container(module_code.as_bytes());

        let dir_decompressed = {
            let mut out = Vec::new();
            push_record(&mut out, 0x0004, b"VBAProject");
            push_record(&mut out, 0x0019, b"Module1");
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(b"Module1");
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name);
            push_record(&mut out, 0x0021, &0u16.to_le_bytes());
            // Intentionally omit MODULETEXTOFFSET (0x0031) so we exercise the signature scan.
            out
        };
        let dir_container = compress_container(&dir_decompressed);

        let mut module_stream = vec![0x01, 0x00, 0x00, 0x99, 0x99, 0x88, 0x77];
        module_stream.extend_from_slice(&container);

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
            s.write_all(&module_stream).expect("write module");
        }

        let vba_bin = ole.into_inner().into_inner();
        let project = VBAProject::parse(&vba_bin).expect("parse");
        let module = project
            .modules
            .iter()
            .find(|m| m.name == "Module1")
            .unwrap();
        assert!(module.code.contains("Sub Hello"));
    }

    #[test]
    fn parses_module_without_text_offset_skipping_invalid_container_candidate() {
        // Regression test for `guess_text_offset()`: some module stream header bytes can resemble an
        // MS-OVBA CompressedContainer signature (0x01 + chunk header signature bits 0b011), but still
        // fail decompression. The heuristic should keep scanning until it finds a *valid*
        // CompressedContainer.
        use std::io::{Cursor, Write};

        let module_code = "Sub Hello()\r\nEnd Sub\r\n";
        let container = compress_container(module_code.as_bytes());

        let dir_decompressed = {
            let mut out = Vec::new();
            push_record(&mut out, 0x0004, b"VBAProject");
            push_record(&mut out, 0x0019, b"Module1");
            let mut stream_name = Vec::new();
            stream_name.extend_from_slice(b"Module1");
            stream_name.extend_from_slice(&0u16.to_le_bytes());
            push_record(&mut out, 0x001A, &stream_name);
            push_record(&mut out, 0x0021, &0u16.to_le_bytes());
            // Intentionally omit MODULETEXTOFFSET (0x0031) so we exercise the signature scan.
            out
        };
        let dir_container = compress_container(&dir_decompressed);

        // Fake CompressedContainer signature at the start:
        // 0x01 + chunk header 0x3FFF (uncompressed, size_field=0xFFF => expects 4096 bytes of data),
        // but we don't include enough bytes for that chunk, so decompression should fail.
        let mut module_stream = vec![0x01, 0xFF, 0x3F, 0xAA, 0xAA, 0xAA, 0xAA];
        module_stream.extend_from_slice(&container);

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
            s.write_all(&module_stream).expect("write module");
        }

        let vba_bin = ole.into_inner().into_inner();
        let project = VBAProject::parse(&vba_bin).expect("parse");
        let module = project.modules.iter().find(|m| m.name == "Module1").unwrap();
        assert!(module.code.contains("Sub Hello"));
    }

    fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(data.len() as u32).to_le_bytes());
        out.extend_from_slice(data);
    }
}
