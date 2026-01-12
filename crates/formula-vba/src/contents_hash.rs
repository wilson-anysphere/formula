use crate::{decompress_container, DirParseError, OleFile, ParseError};

#[derive(Debug, Clone, Default)]
struct ModuleInfo {
    stream_name: String,
    text_offset: Option<usize>,
}

/// Build the MS-OVBA "ContentNormalizedData" byte sequence for a VBA project.
///
/// This is a building block used by MS-OVBA when computing the VBA project digest that a
/// `\x05DigitalSignature*` stream signs.
///
/// This implementation is intentionally focused on correctness for the trickier normalization
/// rules that are easy to regress:
///
/// - **Module ordering** comes from the stored order in `VBA/dir` (`PROJECTMODULES.Modules` order),
///   not alphabetical sorting and not OLE directory enumeration order.
/// - **Module source normalization** treats CR and lone-LF as line breaks, ignores the LF of CRLF,
///   and strips `Attribute ...` lines (case-insensitive, start-of-line match).
/// - **Reference records** are incorporated for a subset of record types (e.g. registered/project
///   references), matching the MS-OVBA §2.4.2.1 pseudocode.
///
/// Spec reference: MS-OVBA §2.4.2.1 "Content Normalized Data".
pub fn content_normalized_data(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    let mut ole = OleFile::open(vba_project_bin)?;

    let dir_bytes = ole
        .read_stream_opt("VBA/dir")?
        .ok_or(ParseError::MissingStream("VBA/dir"))?;
    let dir_decompressed = decompress_container(&dir_bytes)?;

    let mut out = Vec::new();
    let mut modules: Vec<ModuleInfo> = Vec::new();
    let mut current_module: Option<ModuleInfo> = None;

    let mut offset = 0usize;
    while offset < dir_decompressed.len() {
        if offset + 6 > dir_decompressed.len() {
            return Err(DirParseError::Truncated.into());
        }

        let id = u16::from_le_bytes([dir_decompressed[offset], dir_decompressed[offset + 1]]);
        let len = u32::from_le_bytes([
            dir_decompressed[offset + 2],
            dir_decompressed[offset + 3],
            dir_decompressed[offset + 4],
            dir_decompressed[offset + 5],
        ]) as usize;
        offset += 6;
        if offset + len > dir_decompressed.len() {
            return Err(DirParseError::BadRecordLength { id, len }.into());
        }
        let data = &dir_decompressed[offset..offset + len];
        offset += len;

        match id {
            // REFERENCEREGISTERED
            0x000D => {
                out.extend_from_slice(data);
            }

            // REFERENCEPROJECT
            0x000E => {
                out.extend_from_slice(&normalize_reference_project(data)?);
            }

            // MODULENAME: start a new module record group.
            0x0019 => {
                if let Some(m) = current_module.take() {
                    modules.push(m);
                }
                current_module = Some(ModuleInfo {
                    stream_name: String::from_utf8_lossy(data).into_owned(),
                    text_offset: None,
                });
            }

            // MODULESTREAMNAME. Some files include a reserved u16 at the end.
            0x001A => {
                if let Some(m) = current_module.as_mut() {
                    m.stream_name = String::from_utf8_lossy(trim_reserved_u16(data)).into_owned();
                }
            }

            // MODULETEXTOFFSET (u32 LE).
            0x0031 => {
                if let Some(m) = current_module.as_mut() {
                    if data.len() >= 4 {
                        let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
                        m.text_offset = Some(n);
                    }
                }
            }

            _ => {}
        }
    }

    if let Some(m) = current_module.take() {
        modules.push(m);
    }

    for module in modules {
        let stream_path = format!("VBA/{}", module.stream_name);
        let module_stream = ole
            .read_stream_opt(&stream_path)?
            .ok_or(ParseError::MissingStream("module stream"))?;

        let text_offset = module
            .text_offset
            .unwrap_or_else(|| guess_text_offset(&module_stream));
        let text_offset = text_offset.min(module_stream.len());
        let source_container = &module_stream[text_offset..];
        let source = decompress_container(source_container)?;
        out.extend_from_slice(&normalize_module_source(&source));
    }

    Ok(out)
}

fn normalize_reference_project(data: &[u8]) -> Result<Vec<u8>, ParseError> {
    // Minimal parser for the fields used by the MS-OVBA normalization pseudocode.
    //
    // REFERENCEPROJECT (0x000E) contains two u32-len-prefixed strings followed by version integers.
    // The exact MS-OVBA `ContentNormalizedData` logic for this record type is subtle: it builds a
    // temporary byte buffer and then copies bytes until the first NUL byte.
    let mut cur = data;

    let libid_absolute = read_u32_len_prefixed_bytes(&mut cur)?;
    let libid_relative = read_u32_len_prefixed_bytes(&mut cur)?;
    if cur.len() < 6 {
        return Err(DirParseError::Truncated.into());
    }
    let major = u32::from_le_bytes([cur[0], cur[1], cur[2], cur[3]]);
    let minor = u16::from_le_bytes([cur[4], cur[5]]);

    // TempBuffer = LibidAbsolute || LibidRelative || MajorVersion(u32le) || MinorVersion(u16le)
    // Then copy bytes until NUL.
    let mut temp = Vec::new();
    temp.extend_from_slice(&libid_absolute);
    temp.extend_from_slice(&libid_relative);
    temp.extend_from_slice(&major.to_le_bytes());
    temp.extend_from_slice(&minor.to_le_bytes());

    let mut out = Vec::new();
    for b in temp {
        if b == 0x00 {
            break;
        }
        out.push(b);
    }
    Ok(out)
}

fn read_u32_len_prefixed_bytes<'a>(cur: &mut &'a [u8]) -> Result<Vec<u8>, ParseError> {
    if cur.len() < 4 {
        return Err(DirParseError::Truncated.into());
    }
    let len = u32::from_le_bytes([cur[0], cur[1], cur[2], cur[3]]) as usize;
    *cur = &cur[4..];
    if cur.len() < len {
        return Err(DirParseError::Truncated.into());
    }
    let out = cur[..len].to_vec();
    *cur = &cur[len..];
    Ok(out)
}

fn trim_reserved_u16(bytes: &[u8]) -> &[u8] {
    if bytes.len() >= 2 && bytes[bytes.len() - 2..] == [0x00, 0x00] {
        &bytes[..bytes.len() - 2]
    } else {
        bytes
    }
}

fn guess_text_offset(module_stream: &[u8]) -> usize {
    // CompressedContainer starts with 0x01, followed by a chunk header whose signature bits
    // must be 0b011 (MS-OVBA 2.4.1.3.5). This avoids false positives where 0x01 appears in the
    // module stream header.
    for idx in 0..module_stream.len().saturating_sub(3) {
        if module_stream[idx] != 0x01 {
            continue;
        }
        let header = u16::from_le_bytes([module_stream[idx + 1], module_stream[idx + 2]]);
        let signature_bits = (header & 0x7000) >> 12;
        if signature_bits == 0b011 {
            return idx;
        }
    }
    0
}

fn normalize_module_source(bytes: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(bytes.len());

    let mut line_start = 0usize;
    let mut i = 0usize;
    while i < bytes.len() {
        match bytes[i] {
            b'\r' => {
                append_module_line(&bytes[line_start..i], &mut out);
                i += 1;
                if i < bytes.len() && bytes[i] == b'\n' {
                    // Ignore the LF of CRLF.
                    i += 1;
                }
                line_start = i;
            }
            b'\n' => {
                append_module_line(&bytes[line_start..i], &mut out);
                i += 1;
                line_start = i;
            }
            _ => i += 1,
        }
    }

    // If the module does not end with a newline, process the trailing line.
    if line_start < bytes.len() {
        append_module_line(&bytes[line_start..], &mut out);
    }

    out
}

fn append_module_line(line: &[u8], out: &mut Vec<u8>) {
    if is_attribute_line(line) {
        return;
    }
    out.extend_from_slice(line);
    out.extend_from_slice(b"\r\n");
}

fn is_attribute_line(line: &[u8]) -> bool {
    let keyword = b"attribute";
    if line.len() < keyword.len() {
        return false;
    }
    for (a, b) in line[..keyword.len()].iter().zip(keyword.iter()) {
        if a.to_ascii_lowercase() != *b {
            return false;
        }
    }
    if line.len() == keyword.len() {
        return true;
    }
    matches!(line[keyword.len()], b' ' | b'\t')
}

