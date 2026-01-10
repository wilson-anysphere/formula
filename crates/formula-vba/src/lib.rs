//! `formula-vba` provides lightweight parsing utilities for the VBA project
//! embedded in macro-enabled Excel workbooks (`vbaProject.bin`).
//!
//! Current scope (per `docs/08-macro-compatibility.md`):
//! - L1: Preserve `vbaProject.bin` bytes (handled by the XLSX layer).
//! - L2: Parse enough of `vbaProject.bin` to enumerate modules and show source.

mod compression;
mod dir;
mod ole;

pub use compression::{compress_container, decompress_container, CompressionError};
pub use dir::{DirParseError, DirStream, ModuleRecord, ModuleType};
pub use ole::{OleError, OleFile};

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
            .or_else(|| DirStream::detect_codepage(&dir_decompressed).map(|cp| cp as u32).map(encoding_for_codepage))
            .unwrap_or(WINDOWS_1252);

        if let Some(project_stream_bytes) = project_stream_bytes.as_deref() {
            let text = decode_with_encoding(project_stream_bytes, encoding);
            for line in text.lines() {
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
            name: dir_stream
                .project_name
                .clone()
                .or(name_from_project_stream),
            constants: dir_stream.constants.clone(),
            references,
            modules,
        })
    }
}

fn decode_with_encoding(bytes: &[u8], encoding: &'static Encoding) -> String {
    // Module source is commonly stored as MBCS in the project codepage. If it
    // looks like UTF-16LE, decode as UTF-16LE instead (some producers emit it).
    if bytes.len() >= 2 && bytes.len() % 2 == 0 {
        // Same heuristic as `dir` strings: if many high bytes are NUL, treat as UTF-16LE.
        let total = bytes.len() / 2;
        let nul_high = bytes.iter().skip(1).step_by(2).filter(|&&b| b == 0).count();
        if nul_high >= total / 2 {
            let (cow, _) = UTF_16LE.decode_without_bom_handling(bytes);
            return cow.into_owned();
        }
    }

    let (cow, _, _) = encoding.decode(bytes);
    cow.into_owned()
}

fn detect_project_codepage(project_stream_bytes: &[u8]) -> Option<&'static Encoding> {
    // The `PROJECT` stream is plain text; we can find the codepage by scanning
    // the raw bytes for the ASCII `CodePage=` line.
    let mut haystack = project_stream_bytes;
    while let Some(idx) = find_subslice(haystack, b"CodePage=") {
        let after = &haystack[idx + "CodePage=".len()..];
        let mut digits = Vec::new();
        for &b in after {
            if b.is_ascii_digit() {
                digits.push(b);
            } else {
                break;
            }
        }
        if let Ok(n) = std::str::from_utf8(&digits).ok()?.parse::<u32>() {
            return Some(encoding_for_codepage(n));
        }
        haystack = after;
    }
    None
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

fn find_subslice(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() {
        return Some(0);
    }
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
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
            if let Some(end_rel) = first[start + 1..].find('}') {
                let guid = &first[start + 1..start + 1 + end_rel];
                if !guid.is_empty() {
                    reference.guid = Some(guid.to_owned());
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
    module_stream
        .iter()
        .position(|&b| b == 0x01)
        .unwrap_or(0)
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
        assert_eq!(module.attributes.get("VB_Name").map(String::as_str), Some("Module1"));
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

        let module = project.modules.iter().find(|m| m.name == "Module1").unwrap();
        assert!(module.code.contains(comment));
    }

    fn push_record(out: &mut Vec<u8>, id: u16, data: &[u8]) {
        out.extend_from_slice(&id.to_le_bytes());
        out.extend_from_slice(&(data.len() as u32).to_le_bytes());
        out.extend_from_slice(data);
    }
}
