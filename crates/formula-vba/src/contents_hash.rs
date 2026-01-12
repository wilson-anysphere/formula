use encoding_rs::{
    Encoding, BIG5, EUC_KR, GBK, SHIFT_JIS, UTF_16LE, UTF_8, WINDOWS_1250, WINDOWS_1251,
    WINDOWS_1252, WINDOWS_1253, WINDOWS_1254, WINDOWS_1255, WINDOWS_1256, WINDOWS_1257,
    WINDOWS_1258, WINDOWS_874,
};
use md5::Md5;
use sha2::Digest as _;
use sha2::Sha256;

use crate::forms_normalized_data;
use crate::{decompress_container, DirParseError, DirStream, OleFile, ParseError};

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
/// - **Project name and constants** are incorporated by appending the raw record payload bytes for
///   `PROJECTNAME.ProjectName` (0x0004) and `PROJECTCONSTANTS.Constants` (0x000C) in `VBA/dir`
///   record order.
/// - **Reference records** are incorporated for a subset of record types, matching the MS-OVBA
///   §2.4.2.1 pseudocode allowlist:
///   - `REFERENCEREGISTERED` (0x000D)
///   - `REFERENCEPROJECT` (0x000E)
///   - `REFERENCECONTROL` (0x002F)
///   - `REFERENCEEXTENDED` (0x0030)
///   - `REFERENCEORIGINAL` (0x0033)
///   Other reference-related records (e.g. `REFERENCENAME` (0x0016)) MUST NOT contribute.
///
/// Spec reference: MS-OVBA §2.4.2.1 "Content Normalized Data".
pub fn content_normalized_data(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    let mut ole = OleFile::open(vba_project_bin)?;

    let project_stream_bytes = ole.read_stream_opt("PROJECT")?;
    let dir_bytes = ole
        .read_stream_opt("VBA/dir")?
        .ok_or(ParseError::MissingStream("VBA/dir"))?;
    let dir_decompressed = decompress_container(&dir_bytes)?;
    let encoding = project_stream_bytes
        .as_deref()
        .and_then(detect_project_codepage)
        .or_else(|| DirStream::detect_codepage(&dir_decompressed).map(encoding_for_codepage))
        .unwrap_or(WINDOWS_1252);

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
            // PROJECTNAME.ProjectName
            0x0004 => out.extend_from_slice(data),

            // PROJECTCONSTANTS.Constants
            0x000C => out.extend_from_slice(data),

            // MS-OVBA §2.4.2.1 ContentNormalizedData reference record allowlist.
            //
            // NOTE: The spec explicitly includes only some REFERENCE* record variants.
            //
            // REFERENCEREGISTERED (0x000D)
            0x000D => out.extend_from_slice(data),

            // REFERENCEPROJECT (0x000E)
            0x000E => out.extend_from_slice(&normalize_reference_project(data)?),

            // REFERENCECONTROL (0x002F)
            0x002F => out.extend_from_slice(&normalize_reference_control(data)?),

            // REFERENCEEXTENDED (0x0030)
            0x0030 => out.extend_from_slice(data),

            // REFERENCEORIGINAL (0x0033)
            0x0033 => out.extend_from_slice(&normalize_reference_original(data)?),

            // MODULENAME: start a new module record group.
            0x0019 => {
                if let Some(m) = current_module.take() {
                    modules.push(m);
                }
                current_module = Some(ModuleInfo {
                    stream_name: decode_dir_string(data, encoding),
                    text_offset: None,
                });
            }

            // MODULESTREAMNAME. Some files include a reserved u16 at the end.
            0x001A => {
                if let Some(m) = current_module.as_mut() {
                    m.stream_name = decode_dir_string(trim_reserved_u16(data), encoding);
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

/// Compute the MS-OVBA §2.4.2.3 **Content Hash** (v1) for a VBA project.
///
/// Per MS-OSHARED §4.3, the digest bytes used for VBA signature binding are **MD5 (16 bytes)** even
/// when the PKCS#7/CMS signature uses SHA-1/SHA-256 and even when the Authenticode `DigestInfo`
/// algorithm OID indicates SHA-256.
pub fn content_hash_md5(vba_project_bin: &[u8]) -> Result<[u8; 16], ParseError> {
    let normalized = content_normalized_data(vba_project_bin)?;
    Ok(Md5::digest(&normalized).into())
}

/// Compute the MS-OVBA §2.4.2.4 **Agile Content Hash** (v2) for a VBA project, if possible.
///
/// The Agile hash extends the legacy Content Hash by incorporating `FormsNormalizedData`
/// (designer/UserForm storages).
///
/// Returns `Ok(None)` when `FormsNormalizedData` cannot be computed (missing/unparseable data).
pub fn agile_content_hash_md5(vba_project_bin: &[u8]) -> Result<Option<[u8; 16]>, ParseError> {
    let content = content_normalized_data(vba_project_bin)?;
    let forms = match crate::forms_normalized_data(vba_project_bin) {
        Ok(v) => v,
        Err(_) => return Ok(None),
    };

    let mut h = Md5::new();
    h.update(&content);
    h.update(&forms);
    Ok(Some(h.finalize().into()))
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
            if let Ok(cp) = u16::try_from(n) {
                return Some(encoding_for_codepage(cp));
            }
            return Some(WINDOWS_1252);
        }
        haystack = after;
    }
    None
}

fn find_subslice(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() {
        return Some(0);
    }
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

#[derive(Debug, Clone, Default)]
struct ModuleInfoV3 {
    stream_name: String,
    text_offset: Option<usize>,
    // Bytes contributed to V3ContentNormalizedData before the module's normalized source code.
    //
    // We keep this as raw bytes (as stored in the decompressed `VBA/dir` record payloads) to match
    // the MS-OVBA transcript semantics and avoid codepage decoding concerns.
    transcript_prefix: Vec<u8>,
}

/// Build the MS-OVBA §2.4.2 V3 "V3ContentNormalizedData" byte sequence for a VBA project.
///
/// This is the transcript used by MS-OVBA "Contents Hash" version 3, commonly associated with the
/// `\x05DigitalSignatureExt` signature stream.
///
/// Compared to [`content_normalized_data`] (v1-ish), this includes additional metadata required by
/// the v3 transcript, notably:
/// - additional reference record types (e.g. control references), and
/// - module identity/metadata record payloads (module name/stream name/type) in `VBA/dir` order.
///
/// Spec reference: MS-OVBA §2.4.2 "Contents Hash" (version 3).
pub fn v3_content_normalized_data(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    let mut ole = OleFile::open(vba_project_bin)?;

    let project_stream_bytes = ole.read_stream_opt("PROJECT")?;
    let dir_bytes = ole
        .read_stream_opt("VBA/dir")?
        .ok_or(ParseError::MissingStream("VBA/dir"))?;
    let dir_decompressed = decompress_container(&dir_bytes)?;
    let encoding = project_stream_bytes
        .as_deref()
        .and_then(detect_project_codepage)
        .or_else(|| DirStream::detect_codepage(&dir_decompressed).map(encoding_for_codepage))
        .unwrap_or(WINDOWS_1252);

    let mut out = Vec::new();
    let mut current_module: Option<ModuleInfoV3> = None;

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
            // ---- References ----
            //
            // See MS-OVBA §2.4.2 for which reference record types are incorporated in each content
            // hash version.

            // REFERENCEREGISTERED
            0x000D => {
                out.extend_from_slice(data);
            }

            // REFERENCEPROJECT
            0x000E => {
                out.extend_from_slice(&normalize_reference_project(data)?);
            }

            // REFERENCECONTROL
            0x002F => {
                out.extend_from_slice(&normalize_reference_control(data)?);
            }

            // REFERENCEEXTENDED
            0x0030 => {
                out.extend_from_slice(data);
            }

            // REFERENCEORIGINAL
            0x0033 => {
                out.extend_from_slice(&normalize_reference_original(data)?);
            }

            // ---- Modules ----

            // MODULENAME: start a new module record group.
            0x0019 => {
                if let Some(m) = current_module.take() {
                    append_v3_module(&mut out, &mut ole, &m)?;
                }
                let name = decode_dir_string(data, encoding);
                let mut transcript_prefix = Vec::new();
                transcript_prefix.extend_from_slice(data);
                current_module = Some(ModuleInfoV3 {
                    stream_name: name,
                    text_offset: None,
                    transcript_prefix,
                });
            }

            // MODULESTREAMNAME. Some files include a reserved u16 at the end.
            0x001A => {
                if let Some(m) = current_module.as_mut() {
                    let trimmed = trim_reserved_u16(data);
                    m.stream_name = decode_dir_string(trimmed, encoding);
                    m.transcript_prefix.extend_from_slice(trimmed);
                }
            }

            // MODULETYPE (u16 LE).
            0x0021 => {
                if let Some(m) = current_module.as_mut() {
                    m.transcript_prefix.extend_from_slice(data);
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
        append_v3_module(&mut out, &mut ole, &m)?;
    }

    Ok(out)
}

fn append_v3_module(
    out: &mut Vec<u8>,
    ole: &mut OleFile,
    module: &ModuleInfoV3,
) -> Result<(), ParseError> {
    out.extend_from_slice(&module.transcript_prefix);

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
    Ok(())
}

/// Build the MS-OVBA §2.4.2 v3 `ProjectNormalizedData` byte sequence for a `vbaProject.bin`.
///
/// Spec reference: MS-OVBA §2.4.2 ("Contents Hash" version 3).
pub fn project_normalized_data_v3(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    let mut out = v3_content_normalized_data(vba_project_bin)?;
    let forms = forms_normalized_data(vba_project_bin)?;
    out.extend_from_slice(&forms);
    Ok(out)
}

/// Compute the MS-OVBA §2.4.2 "Contents Hash" v3 digest bytes for a `vbaProject.bin`.
///
/// This is the digest that the newest Office signature stream (`\x05DigitalSignatureExt`) binds
/// against.
///
/// MS-OVBA v3 uses **SHA-256** over:
/// `ProjectNormalizedData = V3ContentNormalizedData || FormsNormalizedData`.
pub fn contents_hash_v3(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    let normalized = project_normalized_data_v3(vba_project_bin)?;
    Ok(Sha256::digest(&normalized).to_vec())
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

    Ok(copy_until_nul(&temp))
}

fn normalize_reference_control(data: &[u8]) -> Result<Vec<u8>, ParseError> {
    // Minimal parser for the fields used by the MS-OVBA normalization pseudocode.
    //
    // REFERENCECONTROL (0x002F) contains a u32-len-prefixed libid (LibidTwiddled) followed by
    // reserved version integers. As with REFERENCEPROJECT, the MS-OVBA normalization builds a
    // TempBuffer then copies bytes until the first NUL byte.
    let mut cur = data;

    let libid_twiddled = read_u32_len_prefixed_bytes(&mut cur)?;
    if cur.len() < 6 {
        return Err(DirParseError::Truncated.into());
    }

    // Reserved1 (u32 LE) + Reserved2 (u16 LE)
    let reserved1 = u32::from_le_bytes([cur[0], cur[1], cur[2], cur[3]]);
    let reserved2 = u16::from_le_bytes([cur[4], cur[5]]);

    // TempBuffer = LibidTwiddled || Reserved1(u32le) || Reserved2(u16le)
    // Then copy bytes until NUL.
    let mut temp = Vec::new();
    temp.extend_from_slice(&libid_twiddled);
    temp.extend_from_slice(&reserved1.to_le_bytes());
    temp.extend_from_slice(&reserved2.to_le_bytes());

    Ok(copy_until_nul(&temp))
}

fn normalize_reference_original(data: &[u8]) -> Result<Vec<u8>, ParseError> {
    // Minimal parser for MS-OVBA reference normalization.
    //
    // REFERENCEORIGINAL (0x0033) stores a u32-len-prefixed libid (LibidOriginal). The normalization
    // includes the libid bytes and stops at the first NUL byte.
    let mut cur = data;
    let libid_original = read_u32_len_prefixed_bytes(&mut cur)?;
    Ok(copy_until_nul(&libid_original))
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

fn copy_until_nul(bytes: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    for &b in bytes {
        if b == 0x00 {
            break;
        }
        out.push(b);
    }
    out
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
fn encoding_for_codepage(codepage: u16) -> &'static Encoding {
    match codepage as u32 {
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

fn decode_dir_string(bytes: &[u8], encoding: &'static Encoding) -> String {
    // MS-OVBA dir strings are generally stored using the project codepage, but some records may
    // appear in UTF-16LE form. Use the same heuristic as the main `DirStream` parser so we can
    // reliably locate module streams with non-ASCII names.
    if looks_like_utf16le(bytes) {
        let (cow, _) = UTF_16LE.decode_without_bom_handling(bytes);
        return cow.into_owned();
    }

    let (cow, _, _) = encoding.decode(bytes);
    cow.into_owned()
}

fn looks_like_utf16le(bytes: &[u8]) -> bool {
    if bytes.len() < 2 || bytes.len() % 2 != 0 {
        return false;
    }

    // If a substantial portion of the high bytes are NUL, it's probably UTF-16LE for ASCII-range
    // characters (common for simple names).
    let high_bytes = bytes.iter().skip(1).step_by(2);
    let total = bytes.len() / 2;
    let nul_count = high_bytes.filter(|&&b| b == 0).count();
    // Use a ceiling half threshold. For very short inputs (e.g. 2 bytes), `total / 2` is 0 and
    // would incorrectly classify any 2-byte MBCS string as UTF-16LE.
    nul_count * 2 >= total
}
