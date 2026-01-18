use std::io::{Cursor, Read};

use encoding_rs::{Encoding, UTF_16LE, WINDOWS_1252};
use md5::Md5;
use sha2::Digest as _;
use sha2::Sha256;

use crate::forms_normalized_data;
use crate::dir::ModuleRecord;
use crate::{decompress_container, DirParseError, DirStream, OleFile, ParseError};

#[derive(Debug, Clone, Default)]
struct ModuleInfo {
    stream_name: String,
    text_offset: Option<usize>,
    // Tracks whether we've seen any non-name module records (e.g. stream name / text offset). This
    // is used to disambiguate MODULENAMEUNICODE when both ANSI+Unicode record variants are present.
    seen_non_name_record: bool,
}

// MS-OVBA §2.4.2.5 DefaultAttributes list (byte-equality, NOT case-insensitive compare).
const V3_DEFAULT_ATTRIBUTES: [&[u8]; 7] = [
    b"Attribute VB_Base = \"0{00020820-0000-0000-C000-000000000046}\"",
    b"Attribute VB_GlobalNameSpace = False",
    b"Attribute VB_Creatable = False",
    b"Attribute VB_PredeclaredId = True",
    b"Attribute VB_Exposed = True",
    b"Attribute VB_TemplateDerived = False",
    b"Attribute VB_Customizable = True",
];

// MS-OVBA §2.4.2.5: case-insensitive prefix for skipping VB_Name lines.
const V3_VB_NAME_PREFIX: &[u8] = b"Attribute VB_Name = ";

/// Build a `ProjectNormalizedData`-like transcript for a VBA project.
///
/// Spec note: MS-OVBA §2.4.2.6 defines `ProjectNormalizedData` via `NormalizeProjectStream` over the
/// textual `PROJECT` stream. This helper additionally incorporates selected project-information
/// record payload bytes from `VBA/dir` and (best-effort) designer storage bytes
/// (`FormsNormalizedData`).
///
/// It is primarily intended for debugging/tests and should not be treated as a strict
/// spec-accurate implementation of MS-OVBA §2.4.2.6.
///
/// Like the spec pseudocode, it ignores the optional `ProjectWorkspace` / `[Workspace]` section in
/// the `PROJECT` stream (which is intended to be user/machine-local and MUST NOT influence
/// hashing/signature binding).
pub fn project_normalized_data(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    // Project information record IDs (MS-OVBA 2.3.4.2.1).
    //
    // Note: `VBA/dir` stores records as: u16 id, u32 len, `len` bytes of record data.
    const PROJECTSYSKIND: u16 = 0x0001;
    // PROJECTCOMPATVERSION (0x004A) is present in many real-world `VBA/dir` streams and must be
    // incorporated into this legacy `ProjectNormalizedData` transcript (see
    // `tests/project_normalized_data.rs`).
    //
    // We only consider it while parsing the ProjectInformation section (before the first
    // MODULENAME record). Once the first module record group begins, we stop incorporating `VBA/dir`
    // records into the project-info prefix.
    const PROJECTCOMPATVERSION: u16 = 0x004A;
    const PROJECTLCID: u16 = 0x0002;
    const PROJECTCODEPAGE: u16 = 0x0003;
    const PROJECTNAME: u16 = 0x0004;
    const PROJECTDOCSTRING: u16 = 0x0005;
    const PROJECTHELPFILEPATH: u16 = 0x0006;
    const PROJECTHELPCONTEXT: u16 = 0x0007;
    const PROJECTLIBFLAGS: u16 = 0x0008;
    const PROJECTVERSION: u16 = 0x0009;
    const PROJECTCONSTANTS: u16 = 0x000C;
    const PROJECTLCIDINVOKE: u16 = 0x0014;

    // Some producers emit a second, Unicode form record immediately following the ANSI form.
    // The MS-OVBA pseudocode prefers the Unicode form when present.
    //
    // These record IDs are not part of the minimal VBA project fixture used by this repo, but they
    // occur in real-world files.
    const PROJECTDOCSTRINGUNICODE: u16 = 0x0040;
    const PROJECTDOCSTRINGUNICODE_ALT: u16 = 0x0041;
    // PROJECTHELPFILEPATH has an optional second string record that is often Unicode (0x003D), with
    // an observed alternate ID (0x0042).
    const PROJECTHELPFILEPATH2: u16 = 0x003D;
    const PROJECTHELPFILEPATH2_ALT: u16 = 0x0042;
    const PROJECTCONSTANTSUNICODE: u16 = 0x003C;
    const PROJECTCONSTANTSUNICODE_ALT: u16 = 0x0043;

    let mut ole = OleFile::open(vba_project_bin)?;

    let dir_bytes = ole
        .read_stream_opt("VBA/dir")?
        .ok_or(ParseError::MissingStream("VBA/dir"))?;
    let dir_decompressed = decompress_container(&dir_bytes)?;

    let mut out = Vec::new();

    let mut offset = 0usize;
    let mut in_project_information = true;
    while offset < dir_decompressed.len() {
        let Some(id_end) = offset.checked_add(2) else {
            return Err(DirParseError::Truncated.into());
        };
        let Some(id_bytes) = dir_decompressed.get(offset..id_end) else {
            return Err(DirParseError::Truncated.into());
        };
        let id = u16::from_le_bytes([id_bytes[0], id_bytes[1]]);

        // PROJECTVERSION (0x0009) is fixed-length in spec-compliant `VBA/dir` streams, but some
        // synthetic fixtures encode it in a TLV form (`Id || Size || Data`). Disambiguate by
        // checking which interpretation yields a plausible next record boundary.
        if id == PROJECTVERSION {
            let Some(header_end) = offset.checked_add(6) else {
                return Err(DirParseError::Truncated.into());
            };
            let Some(header) = dir_decompressed.get(offset..header_end) else {
                return Err(DirParseError::Truncated.into());
            };
            let size_or_reserved =
                u32::from_le_bytes([header[2], header[3], header[4], header[5]]) as usize;

            let tlv_end = offset
                .checked_add(6)
                .and_then(|v| v.checked_add(size_or_reserved));
            let fixed_end = offset
                .checked_add(12)
                .ok_or_else(|| DirParseError::Truncated)?;

            let tlv_next_ok = tlv_end.is_some_and(|end| {
                looks_like_projectversion_following_record(&dir_decompressed, end)
            });
            let fixed_next_ok =
                looks_like_projectversion_following_record(&dir_decompressed, fixed_end);

            // Prefer the fixed-length interpretation when the TLV interpretation would leave us at
            // an implausible record boundary, or when the u32 field is `0` (a common reserved value
            // for fixed-length PROJECTVERSION records).
            if fixed_end <= dir_decompressed.len()
                && fixed_next_ok
                && (!tlv_next_ok || size_or_reserved == 0)
            {
                // Only incorporate PROJECTVERSION bytes while we're still in the ProjectInformation
                // section. We still parse/skip the record after module records begin so we can
                // validate record framing to EOF (strictness) without accidentally including bytes
                // from the module section.
                if in_project_information {
                    let start = offset
                        .checked_add(2)
                        .ok_or_else(|| DirParseError::Truncated)?;
                    let bytes = dir_decompressed
                        .get(start..fixed_end)
                        .ok_or_else(|| DirParseError::Truncated)?;
                    out.extend_from_slice(bytes);
                }
                offset = fixed_end;
                continue;
            }
        }

        let Some(header_end) = offset.checked_add(6) else {
            return Err(DirParseError::Truncated.into());
        };
        let Some(header) = dir_decompressed.get(offset..header_end) else {
            return Err(DirParseError::Truncated.into());
        };
        let len = u32::from_le_bytes([header[2], header[3], header[4], header[5]]) as usize;
        offset += 6;
        let Some(data_end) = offset.checked_add(len) else {
            return Err(DirParseError::BadRecordLength { id, len }.into());
        };
        if data_end > dir_decompressed.len() {
            return Err(DirParseError::BadRecordLength { id, len }.into());
        }
        let data = dir_decompressed
            .get(offset..data_end)
            .ok_or_else(|| DirParseError::Truncated)?;
        offset = data_end;

        // Stop *interpreting* records once we hit the first module record group, but keep parsing
        // to the end of the stream so truncation/length errors are still reported (strictness).
        //
        // Treat both MODULENAME (0x0019) and MODULENAMEUNICODE (0x0047) as the beginning of module
        // records, since some simplified dir encodings may omit the ANSI record and begin directly
        // with the Unicode variant.
        //
        // This avoids accidentally treating module-level records (some of which may reuse numeric
        // IDs by context) as ProjectInformation records while still validating the overall record
        // framing.
        if id == 0x0019 || id == 0x0047 {
            in_project_information = false;
            continue;
        }
        if !in_project_information {
            continue;
        }

        let next_id = peek_next_record_id(&dir_decompressed, offset);

        match id {
            PROJECTSYSKIND
            | PROJECTCOMPATVERSION
            | PROJECTLCID
            | PROJECTLCIDINVOKE
            | PROJECTCODEPAGE
            | PROJECTNAME
            | PROJECTHELPCONTEXT
            | PROJECTLIBFLAGS
            | PROJECTVERSION => {
                out.extend_from_slice(data);
            }

            PROJECTDOCSTRING => {
                if !matches!(
                    next_id,
                    Some(PROJECTDOCSTRINGUNICODE) | Some(PROJECTDOCSTRINGUNICODE_ALT)
                ) {
                    out.extend_from_slice(data);
                }
            }
            PROJECTDOCSTRINGUNICODE | PROJECTDOCSTRINGUNICODE_ALT => {
                out.extend_from_slice(trim_u32_len_prefix_unicode_string(data));
            }

            PROJECTHELPFILEPATH => {
                // Prefer the second string record when present; it is commonly a Unicode form.
                if !matches!(
                    next_id,
                    Some(PROJECTHELPFILEPATH2) | Some(PROJECTHELPFILEPATH2_ALT)
                ) {
                    out.extend_from_slice(data);
                }
            }
            PROJECTHELPFILEPATH2 | PROJECTHELPFILEPATH2_ALT => {
                out.extend_from_slice(trim_u32_len_prefix_unicode_string(data));
            }

            PROJECTCONSTANTS => {
                if !matches!(
                    next_id,
                    Some(PROJECTCONSTANTSUNICODE) | Some(PROJECTCONSTANTSUNICODE_ALT)
                ) {
                    out.extend_from_slice(data);
                }
            }
            PROJECTCONSTANTSUNICODE | PROJECTCONSTANTSUNICODE_ALT => {
                out.extend_from_slice(trim_u32_len_prefix_unicode_string(data));
            }

            _ => {}
        }
    }

    // MS-OVBA ProjectNormalizedData also incorporates specific data from the `PROJECT` stream and
    // (when present) the designer storage bytes referenced by `BaseClass=`. These are subtle to get
    // right because MS-OVBA defines `NWLN` as CRLF *or* LFCR.
    //
    // Keep this best-effort: if the project stream or designer storages are missing, we still
    // return the dir-record-derived prefix above (useful for unit tests and partial inputs).
    let mut project_properties = Vec::new();
    let mut host_extender_info = Vec::new();
    if let Some(project_stream_bytes) = ole.read_stream_opt("PROJECT")? {
        project_properties =
            project_properties_normalized_bytes(vba_project_bin, &dir_decompressed, &project_stream_bytes);
        host_extender_info = host_extender_info_normalized_bytes(&project_stream_bytes);
    }

    out.extend_from_slice(&project_properties);
    out.extend_from_slice(&host_extender_info);

    Ok(out)
}

fn project_properties_normalized_bytes(
    vba_project_bin: &[u8],
    dir_decompressed: &[u8],
    project_stream_bytes: &[u8],
) -> Vec<u8> {
    // ProjectProperties are MBCS/ASCII bytes; we must preserve them verbatim in the transcript.
    // However, resolving `BaseClass=` designer module identifiers requires decoding into a Rust
    // `String` so we can match them to `VBA/dir` module records.

    fn is_excluded_project_property_name(name: &[u8]) -> bool {
        // MS-OVBA §2.4.2.6 exclusions:
        // - ProjectId (`ID=...`)
        // - ProjectDocModule (`Document=...`)
        // - ProjectProtectionState / ProjectPassword / ProjectVisibilityState (commonly `CMG` / `DPB` / `GC`)
        //
        // Some producers also emit longer-key variants (`ProtectionState`, `Password`, `VisibilityState`) or
        // alternate doc-module keys (`DocModule`); treat them as excluded too for robustness.
        name.eq_ignore_ascii_case(b"ID")
            || name.eq_ignore_ascii_case(b"Document")
            || name.eq_ignore_ascii_case(b"DocModule")
            || name.eq_ignore_ascii_case(b"CMG")
            || name.eq_ignore_ascii_case(b"DPB")
            || name.eq_ignore_ascii_case(b"GC")
            || name.eq_ignore_ascii_case(b"ProtectionState")
            || name.eq_ignore_ascii_case(b"Password")
            || name.eq_ignore_ascii_case(b"VisibilityState")
    }

    let encoding = crate::detect_project_codepage(project_stream_bytes)
        .or_else(|| {
            DirStream::detect_codepage(dir_decompressed)
                .map(|cp| crate::encoding_for_codepage(cp as u32))
        })
        .unwrap_or(WINDOWS_1252);

    // We need `VBA/dir` module records to map BaseClass identifiers to designer storage names.
    let dir_stream = DirStream::parse_with_encoding(dir_decompressed, encoding).ok();
    let modules: &[ModuleRecord] = dir_stream.as_ref().map(|d| d.modules.as_slice()).unwrap_or(&[]);

    // To normalize designer storages we need storage enumeration, which is only exposed by the
    // underlying `cfb::CompoundFile`. Keep this best-effort: if anything goes wrong, we still
    // output the property tokens without designer bytes.
    let mut file = cfb::CompoundFile::open(Cursor::new(vba_project_bin)).ok();

    let mut out = Vec::new();
    for raw_line in split_nwln_lines(project_stream_bytes) {
        let mut line = trim_ascii_whitespace(raw_line);
        if line.is_empty() {
            continue;
        }

        // Some producers may include a UTF-8 BOM at the start of the stream. Strip it for key
        // matching and output stability.
        if line.starts_with(&[0xEF, 0xBB, 0xBF]) {
            line = trim_ascii_whitespace(&line[3..]);
            if line.is_empty() {
                continue;
            }
        }

        // Section headers are bracketed, e.g. `[Host Extender Info]` or `[Workspace]`.
        // Per MS-OVBA, `ProjectProperties` ends at the first section header.
        if line.starts_with(b"[") && line.ends_with(b"]") {
            break;
        }

        let Some(pos) = line.iter().position(|&b| b == b'=') else {
            continue;
        };

        let name = trim_ascii_whitespace(&line[..pos]);
        let Some(value_bytes) = line.get(pos.saturating_add(1)..) else {
            continue;
        };
        let mut value = trim_ascii_whitespace(value_bytes);
        value = strip_ascii_quotes(value);

        if name.is_empty() {
            continue;
        }

        if is_excluded_project_property_name(name) {
            continue;
        }

        // MS-OVBA §2.4.2.6: for ProjectDesignerModule (`BaseClass=`) append `NormalizeDesignerStorage`
        // output for the referenced designer storage *before* appending the name/value token bytes.
        if name.eq_ignore_ascii_case(b"BaseClass") {
            if let Some(file) = file.as_mut() {
                let (cow, _, _) = encoding.decode(value);
                let module_identifier = cow.trim().trim_matches('"');
                if let Some(storage_name) =
                    match_designer_module_stream_name(modules, module_identifier)
                {
                    if let Ok(bytes) = normalize_designer_storage(file, storage_name) {
                        out.extend_from_slice(&bytes);
                    }
                }
            }
        }

        // MS-OVBA pseudocode appends property name bytes then property value bytes, with no
        // separator and with any surrounding quotes removed.
        out.extend_from_slice(name);
        out.extend_from_slice(value);
    }

    out
}

fn match_designer_module_stream_name<'a>(
    modules: &'a [ModuleRecord],
    module_identifier: &str,
) -> Option<&'a str> {
    if let Some(m) = modules.iter().find(|m| m.name == module_identifier) {
        return Some(m.stream_name.as_str());
    }

    // VB(A) module identifiers are case-insensitive. The MS-OVBA pseudocode does not specify the
    // comparison algorithm, but in practice we need to match regardless of ASCII case differences
    // (common) and also for non-ASCII module names that appear in some projects (e.g. Windows-1251
    // Cyrillic). Prefer cheap ASCII-only matching first, then fall back to full Unicode lowercasing.
    if let Some(m) = modules
        .iter()
        .find(|m| m.name.eq_ignore_ascii_case(module_identifier))
    {
        return Some(m.stream_name.as_str());
    }

    let needle = module_identifier.to_lowercase();
    modules
        .iter()
        .find(|m| m.name.to_lowercase() == needle)
        .map(|m| m.stream_name.as_str())
}

fn normalize_designer_storage<F: Read + std::io::Seek>(
    file: &mut cfb::CompoundFile<F>,
    storage_name: &str,
) -> std::io::Result<Vec<u8>> {
    let entries = file.walk_storage(storage_name)?;

    // `walk_storage` includes the storage itself as the first entry.
    let mut stream_paths = Vec::new();
    let mut first = true;
    for entry in entries {
        if first {
            first = false;
            continue;
        }
        if entry.is_stream() {
            stream_paths.push(entry.path().to_string_lossy().to_string());
        }
    }

    let mut out = Vec::new();
    for path in stream_paths {
        let mut s = file.open_stream(&path)?;
        let mut buf = Vec::new();
        s.read_to_end(&mut buf)?;
        append_stream_padded_1023(&mut out, &buf);
    }

    Ok(out)
}

fn append_stream_padded_1023(out: &mut Vec<u8>, bytes: &[u8]) {
    out.extend_from_slice(bytes);
    let rem = bytes.len() % 1023;
    if rem != 0 {
        out.extend(std::iter::repeat_n(0u8, 1023 - rem));
    }
}

fn strip_ascii_quotes(bytes: &[u8]) -> &[u8] {
    if bytes.len() >= 2 && bytes[0] == b'"' && bytes[bytes.len() - 1] == b'"' {
        &bytes[1..bytes.len() - 1]
    } else {
        bytes
    }
}

fn peek_next_record_id(bytes: &[u8], offset: usize) -> Option<u16> {
    // `VBA/dir` records are stored as: u16 id, u32 len, payload bytes.
    // Only treat a next record as present when the full 6-byte header is in-bounds.
    let end = offset.checked_add(6)?;
    let hdr = bytes.get(offset..end)?;
    Some(u16::from_le_bytes([hdr[0], hdr[1]]))
}

fn looks_like_dir_record_header(bytes: &[u8], offset: usize) -> bool {
    if offset == bytes.len() {
        return true;
    }
    let id_end = match offset.checked_add(2) {
        Some(v) => v,
        None => return false,
    };
    let Some(id_bytes) = bytes.get(offset..id_end) else {
        return false;
    };
    let id = u16::from_le_bytes([id_bytes[0], id_bytes[1]]);

    // PROJECTVERSION (0x0009) is a fixed-length record (12 bytes total) in many spec-compliant dir
    // streams; v3 transcript parsing treats it as fixed-length.
    if id == 0x0009 {
        return offset
            .checked_add(12)
            .is_some_and(|end| end <= bytes.len());
    }

    let hdr_end = match offset.checked_add(6) {
        Some(v) => v,
        None => return false,
    };
    let Some(hdr) = bytes.get(offset..hdr_end) else {
        return false;
    };
    let len = u32::from_le_bytes([hdr[2], hdr[3], hdr[4], hdr[5]]) as usize;
    hdr_end
        .checked_add(len)
        .is_some_and(|end| end <= bytes.len())
}

fn looks_like_projectversion_following_record(bytes: &[u8], offset: usize) -> bool {
    if offset == bytes.len() {
        return true;
    }
    let hdr_end = match offset.checked_add(6) {
        Some(v) => v,
        None => return false,
    };
    let Some(hdr) = bytes.get(offset..hdr_end) else {
        return false;
    };
    let id = u16::from_le_bytes([hdr[0], hdr[1]]);
    // After PROJECTVERSION, we expect either PROJECTCONSTANTS (0x000C), a reference record, the
    // ProjectModules header, or (in some real-world streams) module records.
    if !matches!(
        id,
        0x000C
            | 0x003C
            | 0x0043
            | 0x003D
            | 0x0042
            | 0x0041
            | 0x004A
            | 0x000D
            | 0x000E
            | 0x0016
            | 0x002F
            | 0x0030
            | 0x0033
            | 0x000F
            | 0x0013
            | 0x0010
            | 0x0019
            | 0x0047
            | 0x001A
            | 0x0032
    ) {
        return false;
    }
    let len = u32::from_le_bytes([hdr[2], hdr[3], hdr[4], hdr[5]]) as usize;
    hdr_end
        .checked_add(len)
        .is_some_and(|end| end <= bytes.len())
}

fn host_extender_info_normalized_bytes(project_stream_bytes: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();

    let mut in_section = false;
    for raw_line in split_nwln_lines(project_stream_bytes) {
        let mut line = trim_ascii_whitespace(raw_line);

        // Some producers may include a UTF-8 BOM at the start of the stream. Strip it for
        // stable section parsing.
        if line.starts_with(&[0xEF, 0xBB, 0xBF]) {
            line = trim_ascii_whitespace(&line[3..]);
        }

        // Be permissive: match the section header case-insensitively.
        if line.eq_ignore_ascii_case(b"[Host Extender Info]") {
            in_section = true;
            out.extend_from_slice(b"Host Extender Info");
            continue;
        }

        if in_section {
            // Any new section header ends the `[Host Extender Info]` section.
            if line.starts_with(b"[") && line.ends_with(b"]") {
                break;
            }

            if starts_with_ignore_ascii_case(line, b"HostExtenderRef=") {
                // MS-OVBA pseudocode appends HostExtenderRef "without NWLN". Be robust even if
                // newline bytes slip through by stripping both CR and LF explicitly.
                out.extend(line.iter().copied().filter(|&b| b != b'\r' && b != b'\n'));
            }
        }
    }

    out
}

fn split_nwln_lines(bytes: &[u8]) -> Vec<&[u8]> {
    // MS-OVBA `NWLN` is either CRLF or LFCR. Some producers also emit lone CR or lone LF.
    let mut lines = Vec::new();
    let mut start = 0usize;
    let mut i = 0usize;
    while i < bytes.len() {
        match bytes[i] {
            b'\r' => {
                lines.push(&bytes[start..i]);
                let next = i
                    .checked_add(1)
                    .and_then(|idx| bytes.get(idx))
                    .copied();
                if matches!(next, Some(b'\n')) {
                    i += 2;
                } else {
                    i += 1;
                }
                start = i;
            }
            b'\n' => {
                lines.push(&bytes[start..i]);
                let next = i
                    .checked_add(1)
                    .and_then(|idx| bytes.get(idx))
                    .copied();
                if matches!(next, Some(b'\r')) {
                    i += 2;
                } else {
                    i += 1;
                }
                start = i;
            }
            _ => i += 1,
        }
    }

    if start < bytes.len() {
        lines.push(&bytes[start..]);
    }

    lines
}

fn trim_ascii_whitespace(bytes: &[u8]) -> &[u8] {
    let mut start = 0usize;
    let mut end = bytes.len();
    while start < end && bytes[start].is_ascii_whitespace() {
        start += 1;
    }
    while end > start && bytes[end - 1].is_ascii_whitespace() {
        end -= 1;
    }
    &bytes[start..end]
}

/// Build the MS-OVBA "ContentNormalizedData" byte sequence for a VBA project.
///
/// This is a building block used by MS-OVBA when computing the VBA signature binding digest
/// ("Contents Hash") that a `\x05DigitalSignature*` stream signs.
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
///     Other reference-related records (e.g. `REFERENCENAME` (0x0016)) MUST NOT contribute.
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
        .and_then(crate::detect_project_codepage)
        .or_else(|| {
            DirStream::detect_codepage(&dir_decompressed)
                .map(|cp| crate::encoding_for_codepage(cp as u32))
        })
        .unwrap_or(WINDOWS_1252);

    // Prefer parsing the decompressed `VBA/dir` stream using the MS-OVBA §2.3.4.2 structured record
    // layout. Many synthetic fixtures (including some in our test suite) use a simplified TLV-ish
    // encoding (`u16 id || u32 len || data`) that is not fully spec-accurate; for those we fall
    // back to the existing forgiving parser below.
    //
    // This keeps backwards compatibility for tests/fixtures while enabling correct Contents Hash
    // recomputation for real-world projects that follow the spec record layout (with module/dir
    // terminators, nested records, and reserved/unicode fields).
    if let Some(strict) = content_normalized_data_strict(&mut ole, &dir_decompressed, encoding) {
        return Ok(strict);
    }

    let mut out = Vec::new();
    let mut modules: Vec<ModuleInfo> = Vec::new();
    let mut current_module: Option<ModuleInfo> = None;
    // Some `VBA/dir` layouts encode the Unicode module stream name as a trailing sub-record after
    // MODULESTREAMNAME (0x001A). Track whether we're expecting that sub-record so we can handle
    // alternate record IDs (and avoid misinterpreting unrelated records that might share an ID).
    let mut expect_module_stream_name_unicode = false;

    let mut offset = 0usize;
    while offset < dir_decompressed.len() {
        let Some(hdr_end) = offset.checked_add(6) else {
            return Err(DirParseError::Truncated.into());
        };
        let Some(hdr) = dir_decompressed.get(offset..hdr_end) else {
            return Err(DirParseError::Truncated.into());
        };

        let id = u16::from_le_bytes([hdr[0], hdr[1]]);
        let len = u32::from_le_bytes([hdr[2], hdr[3], hdr[4], hdr[5]]) as usize;
        offset += 6;
        let Some(end) = offset.checked_add(len) else {
            return Err(DirParseError::BadRecordLength { id, len }.into());
        };
        let Some(data) = dir_decompressed.get(offset..end) else {
            return Err(DirParseError::BadRecordLength { id, len }.into());
        };
        offset = end;

        // If we just saw MODULESTREAMNAME and were expecting a Unicode stream name sub-record, any
        // other record indicates the Unicode name is absent.
        if expect_module_stream_name_unicode && !matches!(id, 0x0032 | 0x0048) {
            expect_module_stream_name_unicode = false;
        }

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
                expect_module_stream_name_unicode = false;
                current_module = Some(ModuleInfo {
                    stream_name: decode_dir_string(data, encoding),
                    text_offset: None,
                    seen_non_name_record: false,
                });
            }

            // MODULENAMEUNICODE (UTF-16LE).
            //
            // Some `VBA/dir` streams provide a Unicode module name record (0x0047). In the
            // simplified TLV-ish layouts used by some fixtures/producers, this can appear without a
            // preceding MODULENAME record, so treat it as a module-group start when we are not
            // currently in a module record group.
            0x0047 => {
                let start_new = match current_module.as_ref() {
                    None => true,
                    Some(m) => m.seen_non_name_record,
                };
                if start_new {
                    if let Some(m) = current_module.take() {
                        modules.push(m);
                    }
                    expect_module_stream_name_unicode = false;
                    current_module = Some(ModuleInfo {
                        stream_name: decode_dir_unicode_string(data),
                        text_offset: None,
                        seen_non_name_record: false,
                    });
                } else if let Some(m) = current_module.as_mut() {
                    // Update the module-name-derived default stream name; this matters when a
                    // MODULESTREAMNAME record is absent and the stream name must be inferred from
                    // the module name.
                    m.stream_name = decode_dir_unicode_string(data);
                }
            }

            // MODULESTREAMNAME. Some files include a reserved u16 at the end.
            0x001A => {
                if let Some(m) = current_module.as_mut() {
                    m.stream_name = decode_dir_string(trim_reserved_u16(data), encoding);
                    m.seen_non_name_record = true;
                    expect_module_stream_name_unicode = true;
                }
            }

            // MODULESTREAMNAMEUNICODE (UTF-16LE).
            0x0032 => {
                if let Some(m) = current_module.as_mut() {
                    m.stream_name = decode_dir_unicode_string(data);
                    m.seen_non_name_record = true;
                }
                expect_module_stream_name_unicode = false;
            }

            // Some producers use the MS-OVBA Unicode record id for MODULESTREAMNAMEUNICODE (commonly
            // 0x0048) rather than the reserved marker 0x0032. Treat it as a stream-name Unicode
            // variant only when it follows MODULESTREAMNAME.
            0x0048 if expect_module_stream_name_unicode => {
                if let Some(m) = current_module.as_mut() {
                    m.stream_name = decode_dir_unicode_string(data);
                    m.seen_non_name_record = true;
                }
                expect_module_stream_name_unicode = false;
            }

            // MODULETEXTOFFSET (u32 LE).
            0x0031 => {
                if let Some(m) = current_module.as_mut() {
                    if data.len() >= 4 {
                        let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
                        m.text_offset = Some(n);
                        m.seen_non_name_record = true;
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

#[derive(Debug, Clone)]
enum ReferenceForHash {
    /// `REFERENCEREGISTERED` (0x000D): contributes the `Libid` bytes.
    Registered { libid: Vec<u8> },
    /// `REFERENCEPROJECT` (0x000E): contributes a TempBuffer (see MS-OVBA §2.4.2.1) and then copies
    /// bytes until the first NUL byte.
    Project {
        libid_absolute: Vec<u8>,
        libid_relative: Vec<u8>,
        major_version: u32,
        minor_version: u16,
    },
    /// `REFERENCECONTROL` (0x002F): contributes a TempBuffer derived from the twiddled Libid and
    /// reserved fields (copy-until-NUL), plus the associated "extended" portion (treated as a
    /// `REFERENCEEXTENDED`-like payload, included verbatim).
    Control {
        libid_twiddled: Vec<u8>,
        reserved1: u32,
        reserved2: u16,
        extended: Vec<u8>,
    },
    /// `REFERENCEORIGINAL` (0x0033): contributes the `LibidOriginal` bytes (copy-until-NUL).
    Original { libid_original: Vec<u8> },
}

#[derive(Debug, Clone)]
struct ModuleForHash {
    stream_name: String,
    text_offset: u32,
}

#[derive(Debug, Clone)]
struct DirForHash {
    project_name: Vec<u8>,
    project_constants: Vec<u8>,
    references: Vec<ReferenceForHash>,
    modules: Vec<ModuleForHash>,
}

/// Try to parse the decompressed `VBA/dir` stream according to the MS-OVBA §2.3.4.2 structured
/// record layout and compute ContentNormalizedData (MS-OVBA §2.4.2.1).
///
/// Returns `None` when the input does not look like a spec-compliant `dir` stream (common for
/// synthetic fixtures that use a simplified TLV encoding).
fn content_normalized_data_strict(
    ole: &mut OleFile,
    dir_decompressed: &[u8],
    encoding: &'static Encoding,
) -> Option<Vec<u8>> {
    let dir = parse_dir_for_hash_strict(dir_decompressed, encoding)?;

    // MS-OVBA §2.4.2.1 ContentNormalizedData
    let mut out = Vec::new();
    out.extend_from_slice(&dir.project_name);
    out.extend_from_slice(&dir.project_constants);

    for r in &dir.references {
        match r {
            ReferenceForHash::Registered { libid } => out.extend_from_slice(libid),
            ReferenceForHash::Project {
                libid_absolute,
                libid_relative,
                major_version,
                minor_version,
            } => {
                let mut temp = Vec::new();
                temp.extend_from_slice(libid_absolute);
                temp.extend_from_slice(libid_relative);
                temp.extend_from_slice(&major_version.to_le_bytes());
                temp.extend_from_slice(&minor_version.to_le_bytes());
                out.extend_from_slice(&copy_until_nul(&temp));
            }
            ReferenceForHash::Control {
                libid_twiddled,
                reserved1,
                reserved2,
                extended,
            } => {
                let mut temp = Vec::new();
                temp.extend_from_slice(libid_twiddled);
                temp.extend_from_slice(&reserved1.to_le_bytes());
                temp.extend_from_slice(&reserved2.to_le_bytes());
                out.extend_from_slice(&copy_until_nul(&temp));
                out.extend_from_slice(extended);
            }
            ReferenceForHash::Original { libid_original } => {
                out.extend_from_slice(&copy_until_nul(libid_original));
            }
        }
    }

    for m in &dir.modules {
        let stream_path = format!("VBA/{}", m.stream_name);
        let module_stream = ole.read_stream_opt(&stream_path).ok().flatten()?;
        let text_offset = (m.text_offset as usize).min(module_stream.len());
        let source_container = &module_stream[text_offset..];
        let source = decompress_container(source_container).ok()?;
        out.extend_from_slice(&normalize_module_source_strict(&source));
    }

    Some(out)
}

#[derive(Debug, Clone, Copy)]
struct DirCursor<'a> {
    bytes: &'a [u8],
    offset: usize,
}

impl<'a> DirCursor<'a> {
    fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, offset: 0 }
    }

    fn peek_u16(&self) -> Option<u16> {
        let end = self.offset.checked_add(2)?;
        let b = self.bytes.get(self.offset..end)?;
        Some(u16::from_le_bytes([b[0], b[1]]))
    }

    fn read_u16(&mut self) -> Option<u16> {
        let v = self.peek_u16()?;
        self.offset = self.offset.checked_add(2)?;
        Some(v)
    }

    fn read_u32(&mut self) -> Option<u32> {
        let end = self.offset.checked_add(4)?;
        let b = self.bytes.get(self.offset..end)?;
        let v = u32::from_le_bytes([b[0], b[1], b[2], b[3]]);
        self.offset = end;
        Some(v)
    }

    fn take(&mut self, n: usize) -> Option<&'a [u8]> {
        let end = self.offset.checked_add(n)?;
        let out = self.bytes.get(self.offset..end)?;
        self.offset = end;
        Some(out)
    }

    fn skip(&mut self, n: usize) -> Option<()> {
        self.take(n).map(|_| ())
    }
}

fn parse_fixed_record(cur: &mut DirCursor<'_>, expected_id: u16, expected_size: u32) -> Option<()> {
    let id = cur.read_u16()?;
    let size = cur.read_u32()?;
    if id != expected_id || size != expected_size {
        return None;
    }
    cur.skip(size as usize)?;
    Some(())
}

fn parse_record_u16_u16(
    cur: &mut DirCursor<'_>,
    expected_id: u16,
    expected_size: u32,
) -> Option<u16> {
    let id = cur.read_u16()?;
    let size = cur.read_u32()?;
    if id != expected_id || size != expected_size {
        return None;
    }
    cur.read_u16()
}

fn parse_projectname_record(cur: &mut DirCursor<'_>) -> Option<Vec<u8>> {
    let id = cur.read_u16()?;
    if id != 0x0004 {
        return None;
    }
    let size = cur.read_u32()? as usize;
    let name = cur.take(size)?.to_vec();
    Some(name)
}

fn parse_projectdocstring_record(cur: &mut DirCursor<'_>) -> Option<()> {
    let id = cur.read_u16()?;
    if id != 0x0005 {
        return None;
    }
    let size = cur.read_u32()? as usize;
    cur.skip(size)?;
    // Optional Unicode sub-record:
    //   Reserved marker (commonly 0x0040) + SizeOfDocStringUnicode + DocStringUnicode
    //
    // Some producers omit the Unicode form entirely; accept that and continue parsing.
    if let Some(reserved) = cur.peek_u16() {
        if matches!(reserved, 0x0040 | 0x0041) {
            cur.read_u16()?;
            let size_unicode = cur.read_u32()? as usize;
            cur.skip(size_unicode)?;
        }
    }
    Some(())
}

fn parse_projecthelpfilepath_record(cur: &mut DirCursor<'_>) -> Option<()> {
    let id = cur.read_u16()?;
    if id != 0x0006 {
        return None;
    }
    let size1 = cur.read_u32()? as usize;
    cur.skip(size1)?;
    // Optional second path / Unicode form:
    //   Reserved marker (commonly 0x003D) + SizeOfHelpFile2 + HelpFile2
    if let Some(reserved) = cur.peek_u16() {
        if matches!(reserved, 0x003D | 0x0042) {
            cur.read_u16()?;
            let size2 = cur.read_u32()? as usize;
            cur.skip(size2)?;
        }
    }
    Some(())
}

fn parse_projectversion_record(cur: &mut DirCursor<'_>) -> Option<()> {
    let id = cur.read_u16()?;
    if id != 0x0009 {
        return None;
    }
    let _reserved = cur.read_u32()?;
    let _major = cur.read_u32()?;
    let _minor = cur.read_u16()?;
    Some(())
}

fn parse_projectconstants_record(cur: &mut DirCursor<'_>) -> Option<Vec<u8>> {
    let id = cur.read_u16()?;
    if id != 0x000C {
        return None;
    }
    let size = cur.read_u32()? as usize;
    let constants = cur.take(size)?.to_vec();
    // Optional Unicode form:
    //   Reserved marker (commonly 0x003C) + SizeOfConstantsUnicode + ConstantsUnicode
    if let Some(reserved) = cur.peek_u16() {
        if matches!(reserved, 0x003C | 0x0043) {
            cur.read_u16()?;
            let size_unicode = cur.read_u32()? as usize;
            cur.skip(size_unicode)?;
        }
    }
    Some(constants)
}

fn parse_referencename_record(cur: &mut DirCursor<'_>) -> Option<()> {
    let id = cur.read_u16()?;
    if id != 0x0016 {
        return None;
    }
    let size = cur.read_u32()? as usize;
    cur.skip(size)?;
    let _reserved = cur.read_u16()?;
    let size_unicode = cur.read_u32()? as usize;
    cur.skip(size_unicode)?;
    Some(())
}

fn parse_reference_control(cur: &mut DirCursor<'_>) -> Option<()> {
    let id = cur.read_u16()?;
    if id != 0x002F {
        return None;
    }
    let size_twiddled = cur.read_u32()? as usize;
    cur.skip(size_twiddled)?;

    // Optional NameRecordExtended (REFERENCENAME).
    if cur.peek_u16()? == 0x0016 {
        parse_referencename_record(cur)?;
    }

    let reserved3 = cur.read_u16()?;
    if reserved3 != 0x0030 {
        return None;
    }
    let size_extended = cur.read_u32()? as usize;
    cur.skip(size_extended)?;
    Some(())
}

fn parse_reference_control_data_for_hash(data: &[u8]) -> Option<(Vec<u8>, u32, u16)> {
    // `REFERENCECONTROL` record data used by MS-OVBA normalization pseudocode:
    // - u32 len + bytes (LibidTwiddled)
    // - Reserved1 (u32)
    // - Reserved2 (u16)
    //
    // Some producers may omit the u32 length prefix; be tolerant by falling back to treating the
    // final 6 bytes as the reserved fields.
    if data.len() < 6 {
        return None;
    }

    if data.len() >= 4 {
        let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
        let needed = 4usize
            .checked_add(n)?
            .checked_add(6)?;
        if needed <= data.len() {
            let start = 4usize;
            let end = start.checked_add(n)?;
            let reserved1_offset = end;
            let reserved2_offset = reserved1_offset.checked_add(4)?;
            let reserved1_end = reserved1_offset.checked_add(4)?;
            let reserved2_end = reserved2_offset.checked_add(2)?;
            let reserved1_bytes = data.get(reserved1_offset..reserved1_end)?;
            let reserved2_bytes = data.get(reserved2_offset..reserved2_end)?;
            let reserved1 = u32::from_le_bytes([
                reserved1_bytes[0],
                reserved1_bytes[1],
                reserved1_bytes[2],
                reserved1_bytes[3],
            ]);
            let reserved2 = u16::from_le_bytes([reserved2_bytes[0], reserved2_bytes[1]]);
            return Some((data[start..end].to_vec(), reserved1, reserved2));
        }
    }

    // Fallback: treat the final 6 bytes as Reserved1+Reserved2, with the leading bytes as Libid.
    let reserved1_offset = data.len() - 6;
    let reserved2_offset = reserved1_offset.checked_add(4)?;
    let reserved1_end = reserved1_offset.checked_add(4)?;
    let reserved2_end = reserved2_offset.checked_add(2)?;
    let reserved1_bytes = data.get(reserved1_offset..reserved1_end)?;
    let reserved2_bytes = data.get(reserved2_offset..reserved2_end)?;
    let reserved1 = u32::from_le_bytes([
        reserved1_bytes[0],
        reserved1_bytes[1],
        reserved1_bytes[2],
        reserved1_bytes[3],
    ]);
    let reserved2 = u16::from_le_bytes([reserved2_bytes[0], reserved2_bytes[1]]);
    Some((data[..reserved1_offset].to_vec(), reserved1, reserved2))
}

fn parse_reference_control_for_hash(cur: &mut DirCursor<'_>) -> Option<ReferenceForHash> {
    let id = cur.read_u16()?;
    if id != 0x002F {
        return None;
    }

    // The `REFERENCECONTROL` record's u32 field is the record data length; the fields we need for
    // hashing live inside that data (u32-len-prefixed libid + reserved ints).
    let size_control = cur.read_u32()? as usize;
    let control_data = cur.take(size_control)?.to_vec();
    let (libid_twiddled, reserved1, reserved2) = parse_reference_control_data_for_hash(&control_data)?;

    // Optional NameRecordExtended (REFERENCENAME).
    if cur.peek_u16()? == 0x0016 {
        parse_referencename_record(cur)?;
    }

    // ReferenceRecordExtended (0x0030) immediately follows the control record (as part of the
    // `ReferenceControl` structure in MS-OVBA §2.3.4.2.2).
    let reserved3 = cur.read_u16()?;
    if reserved3 != 0x0030 {
        return None;
    }
    let size_extended = cur.read_u32()? as usize;
    let extended = cur.take(size_extended)?.to_vec();

    Some(ReferenceForHash::Control {
        libid_twiddled,
        reserved1,
        reserved2,
        extended,
    })
}

fn parse_reference_for_hash(cur: &mut DirCursor<'_>) -> Option<Option<ReferenceForHash>> {
    // Optional NameRecord.
    if cur.peek_u16()? == 0x0016 {
        parse_referencename_record(cur)?;
    }

    let id = cur.peek_u16()?;
    match id {
        0x000D => {
            let _id = cur.read_u16()?;
            let size = cur.read_u32()? as usize;
            let libid = cur.take(size)?.to_vec();
            Some(Some(ReferenceForHash::Registered { libid }))
        }
        0x000E => {
            let _id = cur.read_u16()?;
            let size_total = cur.read_u32()? as usize;
            let start = cur.offset;

            let size_abs = cur.read_u32()?;
            let libid_abs = cur.take(size_abs as usize)?.to_vec();
            let size_rel = cur.read_u32()?;
            let libid_rel = cur.take(size_rel as usize)?.to_vec();
            let major = cur.read_u32()?;
            let minor = cur.read_u16()?;

            // Ensure we consumed exactly the expected number of bytes.
            let consumed = cur.offset.checked_sub(start)?;
            if consumed > size_total {
                return None;
            }
            if consumed < size_total {
                cur.skip(size_total - consumed)?;
            }

            Some(Some(ReferenceForHash::Project {
                libid_absolute: libid_abs,
                libid_relative: libid_rel,
                major_version: major,
                minor_version: minor,
            }))
        }
        0x002F => {
            let r = parse_reference_control_for_hash(cur)?;
            Some(Some(r))
        }
        0x0033 => {
            let _id = cur.read_u16()?;
            let size_libid = cur.read_u32()? as usize;
            let libid_raw = cur.take(size_libid)?.to_vec();
            let libid_original = if libid_raw.len() >= 4 {
                let n =
                    u32::from_le_bytes([libid_raw[0], libid_raw[1], libid_raw[2], libid_raw[3]])
                        as usize;
                if let Some(end) = 4usize.checked_add(n) {
                    if end <= libid_raw.len() {
                        libid_raw[4..end].to_vec()
                    } else {
                        libid_raw
                    }
                } else {
                    debug_assert!(false, "libid_original length overflow (n={n})");
                    libid_raw
                }
            } else {
                libid_raw
            };
            // Nested REFERENCECONTROL.
            parse_reference_control(cur)?;
            Some(Some(ReferenceForHash::Original { libid_original }))
        }
        _ => None,
    }
}

fn parse_modulename_record(cur: &mut DirCursor<'_>, expected_id: u16) -> Option<Vec<u8>> {
    let id = cur.read_u16()?;
    if id != expected_id {
        return None;
    }
    let size = cur.read_u32()? as usize;
    cur.take(size).map(|b| b.to_vec())
}

fn parse_module_stream_name(
    cur: &mut DirCursor<'_>,
    encoding: &'static Encoding,
) -> Option<String> {
    let id = cur.read_u16()?;
    if id != 0x001A {
        return None;
    }
    let size_name = cur.read_u32()? as usize;
    let raw_name = cur.take(size_name)?.to_vec();

    // Spec-compliant MODULESTREAMNAME can include a Reserved marker + Unicode stream name, but many
    // fixtures omit those fields. Only parse them when the marker is present.
    //
    // Marker values seen in the wild:
    // - 0x0032 (common)
    // - 0x0048 (alternate; some producers use the Unicode record id)
    let unicode_name = if matches!(cur.peek_u16(), Some(0x0032) | Some(0x0048)) {
        cur.read_u16()?;
        let size_unicode = cur.read_u32()? as usize;
        Some(cur.take(size_unicode)?.to_vec())
    } else {
        None
    };

    if let Some(unicode) = unicode_name {
        return Some(decode_dir_unicode_string(&unicode));
    }

    Some(decode_dir_string(trim_reserved_u16(&raw_name), encoding))
}

fn parse_moduledocstring_record(cur: &mut DirCursor<'_>) -> Option<()> {
    let id = cur.read_u16()?;
    // MS-OVBA record IDs vary across producers; accept both 0x001B and 0x001C for MODULEDOCSTRING.
    if id != 0x001B && id != 0x001C {
        return None;
    }
    let size = cur.read_u32()? as usize;
    cur.skip(size)?;
    // Optional Unicode docstring form:
    //   Reserved marker (commonly 0x0048) + SizeOfDocStringUnicode + DocStringUnicode
    if let Some(reserved) = cur.peek_u16() {
        if matches!(reserved, 0x0048 | 0x0049) {
            cur.read_u16()?;
            let size_unicode = cur.read_u32()? as usize;
            cur.skip(size_unicode)?;
        }
    }
    Some(())
}

fn parse_moduleoffset_record(cur: &mut DirCursor<'_>) -> Option<u32> {
    let id = cur.read_u16()?;
    if id != 0x0031 {
        return None;
    }
    let size = cur.read_u32()?;
    if size != 0x0000_0004 {
        return None;
    }
    cur.read_u32()
}

fn parse_dir_for_hash_strict(
    dir_decompressed: &[u8],
    encoding: &'static Encoding,
) -> Option<DirForHash> {
    let mut cur = DirCursor::new(dir_decompressed);

    // PROJECTINFORMATION record (MS-OVBA §2.3.4.2.1)
    parse_fixed_record(&mut cur, 0x0001, 0x0000_0004)?; // PROJECTSYSKIND
    if cur.peek_u16()? == 0x004A {
        parse_fixed_record(&mut cur, 0x004A, 0x0000_0004)?; // PROJECTCOMPATVERSION (optional)
    }
    parse_fixed_record(&mut cur, 0x0002, 0x0000_0004)?; // PROJECTLCID
    parse_fixed_record(&mut cur, 0x0014, 0x0000_0004)?; // PROJECTLCIDINVOKE
    parse_record_u16_u16(&mut cur, 0x0003, 0x0000_0002)?; // PROJECTCODEPAGE
    let project_name = parse_projectname_record(&mut cur)?;
    // Optional PROJECTNAMEUNICODE (record id 0x0040 in many layouts).
    //
    // This record is not incorporated in the v1/v2 ContentNormalizedData transcript (we use the
    // ANSI PROJECTNAME bytes), but we must skip it so parsing remains aligned for spec-compliant
    // `VBA/dir` streams that include it.
    if cur.peek_u16()? == 0x0040 {
        parse_modulename_record(&mut cur, 0x0040)?;
    }
    parse_projectdocstring_record(&mut cur)?;
    parse_projecthelpfilepath_record(&mut cur)?;
    parse_fixed_record(&mut cur, 0x0007, 0x0000_0004)?; // PROJECTHELPCONTEXT
    parse_fixed_record(&mut cur, 0x0008, 0x0000_0004)?; // PROJECTLIBFLAGS
    parse_projectversion_record(&mut cur)?;
    let project_constants = if cur.peek_u16()? == 0x000C {
        parse_projectconstants_record(&mut cur)?
    } else {
        Vec::new()
    };

    // PROJECTREFERENCES record (MS-OVBA §2.3.4.2.2): variable array terminated by PROJECTMODULES (0x000F).
    let mut references = Vec::new();
    loop {
        let next = cur.peek_u16()?;
        if next == 0x000F {
            break;
        }
        if let Some(r) = parse_reference_for_hash(&mut cur)? {
            references.push(r);
        }
    }

    // PROJECTMODULES record (MS-OVBA §2.3.4.2.3)
    let modules_id = cur.read_u16()?;
    if modules_id != 0x000F {
        return None;
    }
    let modules_size = cur.read_u32()?;
    if modules_size != 0x0000_0002 {
        return None;
    }
    let module_count = cur.read_u16()? as usize;

    // PROJECTCOOKIE record (MS-OVBA §2.3.4.2.3.1)
    let cookie_id = cur.read_u16()?;
    if cookie_id != 0x0013 {
        return None;
    }
    let cookie_size = cur.read_u32()?;
    if cookie_size != 0x0000_0002 {
        return None;
    }
    cur.read_u16()?; // Cookie (ignored)

    let mut modules = Vec::new();
    for _ in 0..module_count {
        // MODULENAME
        parse_modulename_record(&mut cur, 0x0019)?;
        // Optional MODULENAMEUNICODE (0x0047)
        if cur.peek_u16()? == 0x0047 {
            parse_modulename_record(&mut cur, 0x0047)?;
        }
        // MODULESTREAMNAME
        let stream_name = parse_module_stream_name(&mut cur, encoding)?;
        // MODULEDOCSTRING
        parse_moduledocstring_record(&mut cur)?;
        // MODULEOFFSET
        let text_offset = parse_moduleoffset_record(&mut cur)?;
        // MODULEHELPCONTEXT
        parse_fixed_record(&mut cur, 0x001E, 0x0000_0004)?;
        // MODULECOOKIE
        parse_fixed_record(&mut cur, 0x002C, 0x0000_0002)?;
        // MODULETYPE (id 0x0021 or 0x0022, followed by reserved u32)
        let module_type_id = cur.read_u16()?;
        if module_type_id != 0x0021 && module_type_id != 0x0022 {
            return None;
        }
        cur.read_u32()?; // Reserved
                         // Optional MODULEREADONLY (0x0025) and MODULEPRIVATE (0x0028)
        if cur.peek_u16()? == 0x0025 {
            cur.read_u16()?;
            cur.read_u32()?;
        }
        if cur.peek_u16()? == 0x0028 {
            cur.read_u16()?;
            cur.read_u32()?;
        }
        // Terminator + reserved
        let term = cur.read_u16()?;
        if term != 0x002B {
            return None;
        }
        let reserved = cur.read_u32()?;
        if reserved != 0 {
            return None;
        }

        modules.push(ModuleForHash {
            stream_name,
            text_offset,
        });
    }

    // Dir stream terminator + reserved (MS-OVBA §2.3.4.2)
    let end = cur.read_u16()?;
    if end != 0x0010 {
        return None;
    }
    let reserved = cur.read_u32()?;
    if reserved != 0 {
        return None;
    }
    if cur.offset != dir_decompressed.len() {
        return None;
    }

    Some(DirForHash {
        project_name,
        project_constants,
        references,
        modules,
    })
}

fn normalize_module_source_strict(bytes: &[u8]) -> Vec<u8> {
    // MS-OVBA §2.4.2.1: strip Attribute lines and normalize line endings to CRLF.
    //
    // Keep strict and fallback behaviors aligned: the strict dir-stream parser should not change
    // the ContentNormalizedData transcript semantics.
    normalize_module_source(bytes)
}

/// Compute the MS-OVBA §2.4.2.3 **Content Hash** (v1) for a VBA project.
///
/// Per MS-OSHARED §4.3, for legacy VBA signature streams, the digest bytes used for signature
/// binding are **MD5 (16 bytes)** even when the PKCS#7/CMS signature uses SHA-1/SHA-256 and even
/// when the Authenticode `DigestInfo` algorithm OID indicates SHA-256.
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

#[derive(Debug, Clone, Default)]
struct ModuleInfoV3 {
    stream_name: String,
    text_offset: Option<usize>,
    name_bytes: Vec<u8>,
    name_unicode_bytes: Option<Vec<u8>>,
    type_record_reserved: Option<[u8; 2]>,
    // Tracks whether we've seen any non-name module records (e.g. stream name / text offset /
    // module type). Used to disambiguate MODULENAMEUNICODE when the ANSI MODULENAME record is
    // absent and modules are described only by Unicode records.
    seen_non_name_record: bool,
    read_only_record_reserved: Option<[u8; 4]>,
    private_record_reserved: Option<[u8; 4]>,
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
        .and_then(crate::detect_project_codepage)
        .or_else(|| {
            DirStream::detect_codepage(&dir_decompressed)
                .map(|cp| crate::encoding_for_codepage(cp as u32))
        })
        .unwrap_or(WINDOWS_1252);

    let mut out = Vec::new();
    let mut current_module: Option<ModuleInfoV3> = None;
    // MS-OVBA `REFERENCENAME` (0x0016) can optionally be followed by a record id 0x003E containing
    // the UTF-16LE "NameUnicode" bytes. The v3 transcript includes these bytes, so keep a small
    // amount of state to associate the 0x003E record with the preceding 0x0016.
    let mut expect_reference_name_unicode = false;
    // Some `VBA/dir` layouts provide a Unicode module stream-name record immediately after
    // MODULESTREAMNAME (0x001A). The canonical ID is 0x0032, but some producers reuse 0x0048.
    // Track this expectation so we don't misinterpret MODULEDOCSTRINGUNICODE as a stream name.
    let mut expect_module_stream_name_unicode = false;

    let mut offset = 0usize;
    while offset < dir_decompressed.len() {
        let Some(id_end) = offset.checked_add(2) else {
            return Err(DirParseError::Truncated.into());
        };
        let Some(id_bytes) = dir_decompressed.get(offset..id_end) else {
            return Err(DirParseError::Truncated.into());
        };

        let id = u16::from_le_bytes([id_bytes[0], id_bytes[1]]);

        if expect_module_stream_name_unicode && !matches!(id, 0x0032 | 0x0048) {
            expect_module_stream_name_unicode = false;
        }

        // Most `VBA/dir` structures begin with an `Id` (u16) followed by a `Size` (u32) and
        // `Size` bytes of payload. However, some fixed-size records (notably `PROJECTVERSION`)
        // do not include a `Size` field in the on-disk representation (MS-OVBA §2.3.4.2.1.11).
        //
        // V3ContentNormalizedData (MS-OVBA §2.4.2.5) needs to be able to scan through these
        // project-information records to reach references/modules, so we special-case them here.
        if id == 0x0009 {
            // PROJECTVERSION (0x0009) record layout is fixed-length in MS-OVBA, but some producers
            // (and synthetic fixtures) encode it as a normal TLV record (`Id || Size || Data`).
            // Disambiguate the two encodings by checking which yields a plausible next record
            // boundary.
            let Some(hdr_end) = offset.checked_add(6) else {
                return Err(DirParseError::Truncated.into());
            };
            let Some(hdr) = dir_decompressed.get(offset..hdr_end) else {
                return Err(DirParseError::Truncated.into());
            };
            // For the fixed-length layout, the u32 after `Id` is `Reserved`.
            // For the TLV layout, the u32 after `Id` is `Size` (and must be excluded from the
            // transcript).
            let size_or_reserved = u32::from_le_bytes([hdr[2], hdr[3], hdr[4], hdr[5]]) as usize;
            let tlv_end = offset
                .checked_add(6)
                .and_then(|v| v.checked_add(size_or_reserved));
            let fixed_end = offset
                .checked_add(12)
                .ok_or_else(|| DirParseError::Truncated)?;

            let tlv_next_ok = tlv_end.is_some_and(|end| {
                looks_like_projectversion_following_record(&dir_decompressed, end)
            });
            let fixed_next_ok =
                looks_like_projectversion_following_record(&dir_decompressed, fixed_end);

            // Prefer the fixed-length interpretation when the TLV interpretation would leave us at
            // an implausible record boundary, or when the u32 field is too small to be a valid TLV
            // payload size (the TLV payload must contain 10 bytes: Reserved||Major||Minor).
            if fixed_end <= dir_decompressed.len()
                && fixed_next_ok
                && (!tlv_next_ok || size_or_reserved < 10)
            {
                let bytes = dir_decompressed
                    .get(offset..fixed_end)
                    .ok_or_else(|| DirParseError::Truncated)?;
                out.extend_from_slice(bytes);
                offset = fixed_end;
                expect_reference_name_unicode = false;
                continue;
            }

            // Fall back to the TLV form: emit the normalized fixed-length bytes (exclude the Size
            // field but include the 10-byte payload prefix).
            let data_start = offset
                .checked_add(6)
                .ok_or_else(|| DirParseError::Truncated)?;
            let data_end = data_start.checked_add(size_or_reserved).ok_or_else(|| {
                DirParseError::BadRecordLength {
                    id,
                    len: size_or_reserved,
                }
            })?;
            if data_end > dir_decompressed.len() {
                return Err(DirParseError::BadRecordLength {
                    id,
                    len: size_or_reserved,
                }
                .into());
            }
            if size_or_reserved < 10 {
                return Err(DirParseError::Truncated.into());
            }

            out.extend_from_slice(&0x0009u16.to_le_bytes());
            let Some(prefix_end) = data_start.checked_add(10) else {
                return Err(DirParseError::Truncated.into());
            };
            let Some(prefix) = dir_decompressed.get(data_start..prefix_end) else {
                return Err(DirParseError::Truncated.into());
            };
            out.extend_from_slice(prefix);
            offset = data_end;
            expect_reference_name_unicode = false;
            continue;
        }

        // MODULESTREAMNAME (0x001A) record layout is also special: the u32 after Id is
        // `SizeOfStreamName`, and a spec-compliant record may include a Reserved=0x0032 marker plus
        // a UTF-16LE `StreamNameUnicode` field *after* the MBCS stream name bytes.
        //
        // If we treat the u32 as the total record length, we'll mis-align parsing when the Unicode
        // stream name is present. Parse the full record and advance `offset` correctly.
        if id == 0x001A {
            let Some(hdr_end) = offset.checked_add(6) else {
                return Err(DirParseError::Truncated.into());
            };
            let Some(hdr) = dir_decompressed.get(offset..hdr_end) else {
                return Err(DirParseError::Truncated.into());
            };
            // For disambiguation, compute `SizeOfStreamName` from the raw bytes. In spec-compliant
            // records this is the MBCS stream-name length; in TLV-ish fixtures it is typically the
            // full payload length.
            let size_name = u32::from_le_bytes([hdr[2], hdr[3], hdr[4], hdr[5]]) as usize;
            let mut cur = DirCursor::new(&dir_decompressed[offset..]);
            let stream_name =
                parse_module_stream_name(&mut cur, encoding).ok_or(DirParseError::Truncated)?;
            offset = offset
                .checked_add(cur.offset)
                .ok_or_else(|| DirParseError::Truncated)?;
            expect_reference_name_unicode = false;

            if let Some(m) = current_module.as_mut() {
                m.stream_name = stream_name;
                m.seen_non_name_record = true;
            }
            // Only expect a separate Unicode stream-name record when the MODULESTREAMNAME record did
            // not include an in-record Unicode tail (Reserved=0x0032 + StreamNameUnicode bytes).
            let expected = 6usize.checked_add(size_name);
            expect_module_stream_name_unicode =
                expected.is_some_and(|v| v == cur.offset) && current_module.is_some();
            continue;
        }

        // Some `REFERENCE*` records include redundant size fields that MS-OVBA specifies MUST be
        // ignored. Do not trust these fields for record framing; instead, derive boundaries from
        // the subsequent size-of-libid fields.
        //
        // This also keeps us robust to malformed size fields (common in hand-crafted fixtures).
        match id {
            0x000D => {
                expect_reference_name_unicode = false;
                append_v3_reference_registered(&mut out, &dir_decompressed, &mut offset)?;
                continue;
            }
            0x000E => {
                expect_reference_name_unicode = false;
                append_v3_reference_project(&mut out, &dir_decompressed, &mut offset)?;
                continue;
            }
            0x002F => {
                expect_reference_name_unicode = false;
                append_v3_reference_control(&mut out, &dir_decompressed, &mut offset)?;
                continue;
            }
            0x0030 => {
                expect_reference_name_unicode = false;
                append_v3_reference_extended(&mut out, &dir_decompressed, &mut offset)?;
                continue;
            }
            _ => {}
        }

        let Some(hdr_end) = offset.checked_add(6) else {
            return Err(DirParseError::Truncated.into());
        };
        let Some(hdr) = dir_decompressed.get(offset..hdr_end) else {
            return Err(DirParseError::Truncated.into());
        };

        let len = u32::from_le_bytes([hdr[2], hdr[3], hdr[4], hdr[5]]) as usize;

        let record_start = offset;
        let header_end = offset.checked_add(6).ok_or_else(|| DirParseError::Truncated)?;
        offset = header_end;
        let record_end = offset
            .checked_add(len)
            .ok_or_else(|| DirParseError::BadRecordLength { id, len })?;
        if record_end > dir_decompressed.len() {
            return Err(DirParseError::BadRecordLength { id, len }.into());
        }
        let data = dir_decompressed
            .get(offset..record_end)
            .ok_or_else(|| DirParseError::Truncated)?;
        offset = record_end;

        if id != 0x003E {
            expect_reference_name_unicode = false;
        }

        match id {
            // ---- Project information (MS-OVBA §2.4.2.5) ----
            //
            // NOTE: The v3 pseudocode appends only a subset of fields for some records.
            // In particular:
            // - `PROJECTSYSKIND` includes Id+Size but not SysKind.
            // - `PROJECTCODEPAGE` includes Id+Size but not CodePage.
            // - `PROJECTDOCSTRING` includes Id+Size and the unicode sub-record header (0x0040 + size),
            //   but not the DocString bytes.
            // - `PROJECTHELPFILEPATH` includes Id+Size and the second path sub-record header (0x003D + size),
            //   but not the HelpFile bytes.
            // - `PROJECTHELPCONTEXT` includes Id+Size but not HelpContext.
            //
            // These rules are easy to regress by accidentally appending entire records.

            // PROJECTSYSKIND (0x0001): include header only.
            0x0001 => out.extend_from_slice(&dir_decompressed[record_start..header_end]),

            // PROJECTLCID (0x0002): include full record.
            0x0002 => out.extend_from_slice(&dir_decompressed[record_start..record_end]),

            // PROJECTLCIDINVOKE (0x0014): include full record.
            0x0014 => out.extend_from_slice(&dir_decompressed[record_start..record_end]),

            // PROJECTCODEPAGE (0x0003): include header only.
            0x0003 => out.extend_from_slice(&dir_decompressed[record_start..header_end]),

            // PROJECTNAME (0x0004): include full record.
            0x0004 => out.extend_from_slice(&dir_decompressed[record_start..record_end]),

            // PROJECTDOCSTRING (0x0005): include header only.
            0x0005 => out.extend_from_slice(&dir_decompressed[record_start..header_end]),
            // PROJECTDOCSTRING unicode sub-record: include header only.
            // Common IDs seen:
            // - 0x0040 (reserved marker)
            // - 0x0041 (Unicode record id in some producers)
            0x0040 | 0x0041 => out.extend_from_slice(&dir_decompressed[record_start..header_end]),

            // PROJECTHELPFILEPATH (0x0006): include header only.
            0x0006 => out.extend_from_slice(&dir_decompressed[record_start..header_end]),
            // PROJECTHELPFILEPATH "HelpFile2" / Unicode sub-record: include header only.
            // Common IDs seen:
            // - 0x003D (reserved marker)
            // - 0x0042 (Unicode record id in some producers)
            0x003D | 0x0042 => out.extend_from_slice(&dir_decompressed[record_start..header_end]),

            // PROJECTHELPCONTEXT (0x0007): include header only.
            0x0007 => out.extend_from_slice(&dir_decompressed[record_start..header_end]),

            // PROJECTLIBFLAGS (0x0008): include full record.
            0x0008 => out.extend_from_slice(&dir_decompressed[record_start..record_end]),

            // PROJECTCONSTANTS (0x000C): include full record.
            0x000C => out.extend_from_slice(&dir_decompressed[record_start..record_end]),
            // PROJECTCONSTANTS unicode sub-record: include full record.
            // Common IDs seen:
            // - 0x003C (reserved marker)
            // - 0x0043 (Unicode record id in some producers)
            0x003C | 0x0043 => out.extend_from_slice(&dir_decompressed[record_start..record_end]),

            // ---- References ----
            //
            // See MS-OVBA §2.4.2 for which reference record types are incorporated in each content
            // hash version.

            // REFERENCENAME
            0x0016 => {
                out.extend_from_slice(&id.to_le_bytes());
                out.extend_from_slice(&(len as u32).to_le_bytes());
                out.extend_from_slice(data);
                expect_reference_name_unicode = true;
            }

            // REFERENCENAME "NameUnicode" marker/record id. This follows a 0x0016 record when the
            // reference includes a UTF-16LE name.
            0x003E if expect_reference_name_unicode => {
                out.extend_from_slice(&id.to_le_bytes());
                out.extend_from_slice(&(len as u32).to_le_bytes());
                out.extend_from_slice(data);
                expect_reference_name_unicode = false;
            }

            // REFERENCEORIGINAL
            0x0033 => {
                out.extend_from_slice(&id.to_le_bytes());
                // The `Size` field for this record is `SizeOfLibidOriginal` and is incorporated by
                // the MS-OVBA v3 pseudocode.
                out.extend_from_slice(&(len as u32).to_le_bytes());
                out.extend_from_slice(data);
                // `REFERENCEORIGINAL` embeds an immediate `REFERENCECONTROL` record; MS-OVBA v3
                // does not incorporate that embedded control in the transcript.
                skip_referenceoriginal_embedded_control(&dir_decompressed, &mut offset)?;
            }

            // ---- ProjectModules / ProjectCookie / dir Terminator ----
            //
            // MS-OVBA §2.4.2.5 V3ContentNormalizedData explicitly appends:
            // - PROJECTMODULES.Id and PROJECTMODULES.Size (but not Count),
            // - PROJECTCOOKIE.Id and PROJECTCOOKIE.Size (but not Cookie),
            // - the dir stream Terminator (0x0010) and Reserved (u32 0) trailer.
            //
            // In the decompressed `VBA/dir` stream we parse records in a generic `Id(u16) || Size(u32)
            // || Data(Size)` form, so including "Id+Size" in the transcript corresponds to emitting
            // the record header bytes and intentionally excluding the record payload.
            0x000F | 0x0013 => {
                out.extend_from_slice(&id.to_le_bytes());
                out.extend_from_slice(&(len as u32).to_le_bytes());
            }
            0x0010 => {
                // The `dir` stream terminator appears after the final module. Flush the pending
                // module so the Terminator/Reserved bytes are the *trailing* bytes of the v3
                // transcript (matching MS-OVBA §2.4.2.5).
                if let Some(m) = current_module.take() {
                    append_v3_module(&mut out, &mut ole, &m)?;
                }

                out.extend_from_slice(&id.to_le_bytes());
                out.extend_from_slice(&(len as u32).to_le_bytes());
            }

            // ---- Modules ----

            // MODULENAME: start a new module record group.
            0x0019 => {
                if let Some(m) = current_module.take() {
                    append_v3_module(&mut out, &mut ole, &m)?;
                }
                current_module = Some(ModuleInfoV3 {
                    stream_name: decode_dir_string(data, encoding),
                    text_offset: None,
                    name_bytes: data.to_vec(),
                    name_unicode_bytes: None,
                    type_record_reserved: None,
                    seen_non_name_record: false,
                    read_only_record_reserved: None,
                    private_record_reserved: None,
                });
            }

            // MODULENAMEUNICODE (UTF-16LE).
            //
            // Some producers emit a Unicode module name record without an ANSI MODULENAME record.
            // Treat it as a module-group start when we're not currently in a module group or when
            // we've already consumed non-name module records for the current module.
            0x0047 => {
                let start_new = match current_module.as_ref() {
                    None => true,
                    Some(m) => m.seen_non_name_record,
                };

                if start_new {
                    if let Some(m) = current_module.take() {
                        append_v3_module(&mut out, &mut ole, &m)?;
                    }
                    current_module = Some(ModuleInfoV3 {
                        stream_name: decode_dir_unicode_string(data),
                        text_offset: None,
                        name_bytes: Vec::new(),
                        // Some producers embed an internal u32 length prefix before the UTF-16LE
                        // bytes. Strip it when it is consistent with the remaining payload.
                        name_unicode_bytes: Some(trim_u32_len_prefix_unicode_string(data).to_vec()),
                        type_record_reserved: None,
                        seen_non_name_record: false,
                        read_only_record_reserved: None,
                        private_record_reserved: None,
                    });
                } else if let Some(m) = current_module.as_mut() {
                    m.name_unicode_bytes =
                        Some(trim_u32_len_prefix_unicode_string(data).to_vec());
                    // If we haven't yet seen an explicit stream name, update the default stream
                    // name derived from MODULENAME.
                    if !m.seen_non_name_record {
                        m.stream_name = decode_dir_unicode_string(data);
                    }
                }
            }

            // MODULESTREAMNAMEUNICODE (UTF-16LE).
            //
            // MS-OVBA `VBA/dir` streams can provide a Unicode variant immediately following
            // MODULESTREAMNAME. Prefer it when present since the stream name is used for OLE lookup.
            0x0032 => {
                if let Some(m) = current_module.as_mut() {
                    m.stream_name = decode_dir_unicode_string(data);
                    m.seen_non_name_record = true;
                }
                expect_module_stream_name_unicode = false;
            }
            // Some producers use 0x0048 as the module stream name Unicode record id (even though
            // 0x0048 is canonically MODULEDOCSTRINGUNICODE). Accept it only when it immediately
            // follows MODULESTREAMNAME.
            0x0048 if expect_module_stream_name_unicode => {
                if let Some(m) = current_module.as_mut() {
                    m.stream_name = decode_dir_unicode_string(data);
                    m.seen_non_name_record = true;
                }
                expect_module_stream_name_unicode = false;
            }

            // MODULETYPE
            //
            // MS-OVBA defines MODULETYPE records with an `Id` of either:
            // - 0x0021 (procedural module)
            // - 0x0022 (non-procedural module)
            //
            // For the v3 transcript (MS-OVBA §2.4.2.5), we only append the module's TypeRecord bytes
            // when `TypeRecord.Id == 0x0021`, and the bytes to append are:
            // `TypeRecord.Id (u16 LE) || TypeRecord.Reserved (u16 LE)`.
            0x0021 => {
                if let Some(m) = current_module.as_mut() {
                    let reserved = if data.len() >= 2 {
                        [data[0], data[1]]
                    } else {
                        [0u8, 0u8]
                    };
                    m.type_record_reserved = Some(reserved);
                    m.seen_non_name_record = true;
                }
            }
            0x0022 => {
                // Explicitly ignored: non-procedural module type records do not contribute to the
                // v3 transcript per MS-OVBA §2.4.2.5 pseudocode.
                if let Some(m) = current_module.as_mut() {
                    m.seen_non_name_record = true;
                }
            }

            // MODULEREADONLY
            //
            // MS-OVBA §2.4.2.5 includes the bytes `ReadOnlyRecord.Id || ReadOnlyRecord.Reserved`
            // when the record is present in the module record.
            0x0025 => {
                if let Some(m) = current_module.as_mut() {
                    // In the decompressed `VBA/dir` stream we parse using a generic `Id(u16) ||
                    // Size(u32) || Data(Size)` form. For fixed-size records like MODULEREADONLY,
                    // the `Reserved` field is stored in the u32 slot and MUST be 0x00000000.
                    let reserved = hdr
                        .get(2..6)
                        .and_then(|bytes| bytes.try_into().ok())
                        .unwrap_or([0u8; 4]);
                    m.read_only_record_reserved = Some(reserved);
                    m.seen_non_name_record = true;
                }
            }

            // MODULEPRIVATE
            //
            // MS-OVBA §2.4.2.5 includes the bytes `PrivateRecord.Id || PrivateRecord.Reserved`
            // when the record is present in the module record.
            0x0028 => {
                if let Some(m) = current_module.as_mut() {
                    let reserved = hdr
                        .get(2..6)
                        .and_then(|bytes| bytes.try_into().ok())
                        .unwrap_or([0u8; 4]);
                    m.private_record_reserved = Some(reserved);
                    m.seen_non_name_record = true;
                }
            }

            // MODULETEXTOFFSET (u32 LE).
            0x0031 => {
                if let Some(m) = current_module.as_mut() {
                    if data.len() >= 4 {
                        let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
                        m.text_offset = Some(n);
                        m.seen_non_name_record = true;
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

fn skip_referenceoriginal_embedded_control(
    dir_decompressed: &[u8],
    offset: &mut usize,
) -> Result<(), ParseError> {
    // Embedded `REFERENCECONTROL` immediately follows the `REFERENCEORIGINAL` libid bytes.
    // If the next record isn't 0x002F, this isn't a spec-compliant embedded control (or we're at
    // end-of-buffer).
    if peek_next_record_id(dir_decompressed, *offset) != Some(0x002F) {
        return Ok(());
    }

    // Skip the embedded REFERENCECONTROL (0x002F) twiddled part.
    skip_v3_reference_control(dir_decompressed, offset)?;

    // Optional NameRecordExtended (0x0016) + NameUnicode (0x003E).
    if peek_next_record_id(dir_decompressed, *offset) == Some(0x0016) {
        skip_dir_record(dir_decompressed, offset)?;
        if peek_next_record_id(dir_decompressed, *offset) == Some(0x003E) {
            skip_dir_record(dir_decompressed, offset)?;
        }
    }

    // Extended control tail (Reserved3=0x0030 marker + size + payload).
    if peek_next_record_id(dir_decompressed, *offset) == Some(0x0030) {
        skip_v3_reference_extended(dir_decompressed, offset)?;
    }

    Ok(())
}

fn append_v3_reference_registered(
    out: &mut Vec<u8>,
    dir_decompressed: &[u8],
    offset: &mut usize,
) -> Result<(), ParseError> {
    let mut cur = DirCursor {
        bytes: dir_decompressed,
        offset: *offset,
    };
    let id = cur.read_u16().ok_or(DirParseError::Truncated)?;
    if id != 0x000D {
        return Err(DirParseError::UnexpectedRecordId { expected: 0x000D, found: id }.into());
    }
    // Record-size field (u32). MS-OVBA v3 transcript does not include this value, but we may use it
    // as an upper bound when advancing the cursor (to safely skip any trailing bytes).
    let size_total = cur.read_u32().ok_or(DirParseError::Truncated)? as usize;
    let payload_start = cur.offset;
    let size_of_libid = cur.read_u32().ok_or(DirParseError::Truncated)? as usize;
    cur.skip(size_of_libid).ok_or(DirParseError::Truncated)?;
    cur.read_u32().ok_or(DirParseError::Truncated)?; // Reserved1
    cur.read_u16().ok_or(DirParseError::Truncated)?; // Reserved2
    let payload_end = cur.offset;

    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&dir_decompressed[payload_start..payload_end]);
    let minimal_len = payload_end.saturating_sub(payload_start);
    let record_end_by_size = payload_start.saturating_add(size_total);
    *offset = if size_total > minimal_len
        && record_end_by_size <= dir_decompressed.len()
        && looks_like_dir_record_header(dir_decompressed, record_end_by_size)
    {
        record_end_by_size
    } else {
        payload_end
    };
    Ok(())
}

fn append_v3_reference_project(
    out: &mut Vec<u8>,
    dir_decompressed: &[u8],
    offset: &mut usize,
) -> Result<(), ParseError> {
    let mut cur = DirCursor {
        bytes: dir_decompressed,
        offset: *offset,
    };
    let id = cur.read_u16().ok_or(DirParseError::Truncated)?;
    if id != 0x000E {
        return Err(DirParseError::UnexpectedRecordId { expected: 0x000E, found: id }.into());
    }
    // Record-size field (u32). MS-OVBA v3 transcript does not include this value, but we use it to
    // safely advance past any trailing bytes after the version fields.
    let size_total = cur.read_u32().ok_or(DirParseError::Truncated)? as usize;
    let payload_start = cur.offset;

    let size_abs = cur.read_u32().ok_or(DirParseError::Truncated)? as usize;
    cur.skip(size_abs).ok_or(DirParseError::Truncated)?;
    let size_rel = cur.read_u32().ok_or(DirParseError::Truncated)? as usize;
    cur.skip(size_rel).ok_or(DirParseError::Truncated)?;
    cur.read_u32().ok_or(DirParseError::Truncated)?; // MajorVersion
    cur.read_u16().ok_or(DirParseError::Truncated)?; // MinorVersion
    let payload_end = cur.offset;

    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&dir_decompressed[payload_start..payload_end]);
    let minimal_len = payload_end.saturating_sub(payload_start);
    let record_end_by_size = payload_start.saturating_add(size_total);
    *offset = if size_total > minimal_len
        && record_end_by_size <= dir_decompressed.len()
        && looks_like_dir_record_header(dir_decompressed, record_end_by_size)
    {
        record_end_by_size
    } else {
        payload_end
    };
    Ok(())
}

fn append_v3_reference_control(
    out: &mut Vec<u8>,
    dir_decompressed: &[u8],
    offset: &mut usize,
) -> Result<(), ParseError> {
    let mut cur = DirCursor {
        bytes: dir_decompressed,
        offset: *offset,
    };
    let id = cur.read_u16().ok_or(DirParseError::Truncated)?;
    if id != 0x002F {
        return Err(DirParseError::UnexpectedRecordId { expected: 0x002F, found: id }.into());
    }
    // Record-size field (u32). MS-OVBA v3 transcript does not include this value, but we use it as
    // an upper bound when advancing the cursor.
    let size_total = cur.read_u32().ok_or(DirParseError::Truncated)? as usize;
    let payload_start = cur.offset;

    let size_of_libid_twiddled = cur.read_u32().ok_or(DirParseError::Truncated)? as usize;
    cur.skip(size_of_libid_twiddled)
        .ok_or(DirParseError::Truncated)?;
    cur.read_u32().ok_or(DirParseError::Truncated)?; // Reserved1
    cur.read_u16().ok_or(DirParseError::Truncated)?; // Reserved2
    let payload_end = cur.offset;

    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&dir_decompressed[payload_start..payload_end]);
    let minimal_len = payload_end.saturating_sub(payload_start);
    let record_end_by_size = payload_start.saturating_add(size_total);
    *offset = if size_total > minimal_len
        && record_end_by_size <= dir_decompressed.len()
        && looks_like_dir_record_header(dir_decompressed, record_end_by_size)
    {
        record_end_by_size
    } else {
        payload_end
    };
    Ok(())
}

fn append_v3_reference_extended(
    out: &mut Vec<u8>,
    dir_decompressed: &[u8],
    offset: &mut usize,
) -> Result<(), ParseError> {
    let mut cur = DirCursor {
        bytes: dir_decompressed,
        offset: *offset,
    };
    let id = cur.read_u16().ok_or(DirParseError::Truncated)?;
    if id != 0x0030 {
        return Err(DirParseError::UnexpectedRecordId { expected: 0x0030, found: id }.into());
    }
    // Record-size field (u32). MS-OVBA v3 transcript does not include this value, but we use it as
    // an upper bound when advancing the cursor.
    let size_total = cur.read_u32().ok_or(DirParseError::Truncated)? as usize;
    let payload_start = cur.offset;

    let size_of_libid_extended = cur.read_u32().ok_or(DirParseError::Truncated)? as usize;
    cur.skip(size_of_libid_extended)
        .ok_or(DirParseError::Truncated)?;
    cur.read_u32().ok_or(DirParseError::Truncated)?; // Reserved4
    cur.read_u16().ok_or(DirParseError::Truncated)?; // Reserved5
    cur.skip(16).ok_or(DirParseError::Truncated)?; // OriginalTypeLib GUID
    cur.read_u32().ok_or(DirParseError::Truncated)?; // Cookie
    let payload_end = cur.offset;

    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&dir_decompressed[payload_start..payload_end]);
    let minimal_len = payload_end.saturating_sub(payload_start);
    let record_end_by_size = payload_start.saturating_add(size_total);
    *offset = if size_total > minimal_len
        && record_end_by_size <= dir_decompressed.len()
        && looks_like_dir_record_header(dir_decompressed, record_end_by_size)
    {
        record_end_by_size
    } else {
        payload_end
    };
    Ok(())
}

fn skip_v3_reference_control(dir_decompressed: &[u8], offset: &mut usize) -> Result<(), ParseError> {
    // Reuse the same parsing logic as the append path, but discard output.
    let mut cur = DirCursor {
        bytes: dir_decompressed,
        offset: *offset,
    };
    let id = cur.read_u16().ok_or(DirParseError::Truncated)?;
    if id != 0x002F {
        return Err(DirParseError::UnexpectedRecordId { expected: 0x002F, found: id }.into());
    }
    let size_total = cur.read_u32().ok_or(DirParseError::Truncated)? as usize;
    let payload_start = cur.offset;
    let size_of_libid_twiddled = cur.read_u32().ok_or(DirParseError::Truncated)? as usize;
    cur.skip(size_of_libid_twiddled)
        .ok_or(DirParseError::Truncated)?;
    cur.read_u32().ok_or(DirParseError::Truncated)?; // Reserved1
    cur.read_u16().ok_or(DirParseError::Truncated)?; // Reserved2
    let payload_end = cur.offset;
    let minimal_len = payload_end.saturating_sub(payload_start);
    let record_end_by_size = payload_start.saturating_add(size_total);
    *offset = if size_total > minimal_len
        && record_end_by_size <= dir_decompressed.len()
        && looks_like_dir_record_header(dir_decompressed, record_end_by_size)
    {
        record_end_by_size
    } else {
        payload_end
    };
    Ok(())
}

fn skip_v3_reference_extended(dir_decompressed: &[u8], offset: &mut usize) -> Result<(), ParseError> {
    let mut cur = DirCursor {
        bytes: dir_decompressed,
        offset: *offset,
    };
    let id = cur.read_u16().ok_or(DirParseError::Truncated)?;
    if id != 0x0030 {
        return Err(DirParseError::UnexpectedRecordId { expected: 0x0030, found: id }.into());
    }
    let size_total = cur.read_u32().ok_or(DirParseError::Truncated)? as usize;
    let payload_start = cur.offset;
    let size_of_libid_extended = cur.read_u32().ok_or(DirParseError::Truncated)? as usize;
    cur.skip(size_of_libid_extended)
        .ok_or(DirParseError::Truncated)?;
    cur.read_u32().ok_or(DirParseError::Truncated)?; // Reserved4
    cur.read_u16().ok_or(DirParseError::Truncated)?; // Reserved5
    cur.skip(16).ok_or(DirParseError::Truncated)?; // OriginalTypeLib GUID
    cur.read_u32().ok_or(DirParseError::Truncated)?; // Cookie
    let payload_end = cur.offset;
    let minimal_len = payload_end.saturating_sub(payload_start);
    let record_end_by_size = payload_start.saturating_add(size_total);
    *offset = if size_total > minimal_len
        && record_end_by_size <= dir_decompressed.len()
        && looks_like_dir_record_header(dir_decompressed, record_end_by_size)
    {
        record_end_by_size
    } else {
        payload_end
    };
    Ok(())
}

fn skip_dir_record(dir_decompressed: &[u8], offset: &mut usize) -> Result<(), ParseError> {
    let Some(hdr_end) = (*offset).checked_add(6) else {
        return Err(DirParseError::Truncated.into());
    };
    let Some(hdr) = dir_decompressed.get(*offset..hdr_end) else {
        return Err(DirParseError::Truncated.into());
    };

    let id = u16::from_le_bytes([hdr[0], hdr[1]]);
    let len = u32::from_le_bytes([hdr[2], hdr[3], hdr[4], hdr[5]]) as usize;
    *offset += 6;
    let Some(end) = (*offset).checked_add(len) else {
        return Err(DirParseError::BadRecordLength { id, len }.into());
    };
    if end > dir_decompressed.len() {
        return Err(DirParseError::BadRecordLength { id, len }.into());
    }
    *offset = end;
    Ok(())
}

fn append_v3_module(
    out: &mut Vec<u8>,
    ole: &mut OleFile,
    module: &ModuleInfoV3,
) -> Result<(), ParseError> {
    // MS-OVBA §2.4.2.5: Include the module's TypeRecord only when its record id is 0x0021.
    if let Some(reserved) = module.type_record_reserved {
        out.extend_from_slice(&0x0021u16.to_le_bytes());
        out.extend_from_slice(&reserved);
    }

    if let Some(reserved) = module.read_only_record_reserved {
        out.extend_from_slice(&0x0025u16.to_le_bytes());
        out.extend_from_slice(&reserved);
    }

    if let Some(reserved) = module.private_record_reserved {
        out.extend_from_slice(&0x0028u16.to_le_bytes());
        out.extend_from_slice(&reserved);
    }

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

    let module_name = module
        .name_unicode_bytes
        .as_deref()
        .unwrap_or(&module.name_bytes);
    append_v3_normalized_module_source(&source, module_name, out);
    Ok(())
}

/// Build `formula-vba`'s current v3 signature-binding transcript for a `vbaProject.bin`.
///
/// Important: despite the historical naming, this is **not** MS-OVBA §2.4.2.6 `ProjectNormalizedData`
/// (`NormalizeProjectStream`), and the concatenation order does not match MS-OVBA §2.4.2.7
/// (`ContentBuffer = V3ContentNormalizedData || ProjectNormalizedData`).
///
/// Current `formula-vba` transcript (best-effort):
/// `(filtered PROJECT stream lines) || V3ContentNormalizedData || FormsNormalizedData`.
///
/// This is hashed as a 32-byte SHA-256 by [`contents_hash_v3`] (common `DigitalSignatureExt`
/// behavior).
///
/// Spec reference: MS-OVBA §2.4.2 ("Contents Hash" version 3).
pub fn project_normalized_data_v3_transcript(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    // Normalize/filter the textual `PROJECT` stream for v3 binding. MS-OVBA v3 binds the signature
    // not just to module content, but also to a filtered subset of the `PROJECT` stream properties
    // (see §2.4.2.6 "ProjectNormalizedData").
    //
    // Some `PROJECT` properties MUST be excluded because they are either security-sensitive or can
    // change without affecting macro semantics (e.g. CMG/DPB/GC protection fields).
    //
    // Additionally, the optional `[Workspace]` section is machine/user-local and MUST NOT influence
    // signature binding, so we stop processing when a non-`[Host Extender Info]` section header is
    // encountered.
    let mut ole = OleFile::open(vba_project_bin)?;
    let project_stream_bytes = ole
        .read_stream_opt("PROJECT")?
        .ok_or(ParseError::MissingStream("PROJECT"))?;

    let mut out = normalize_project_stream_properties_v3(&project_stream_bytes);

    out.extend_from_slice(&v3_content_normalized_data(vba_project_bin)?);
    let forms = forms_normalized_data(vba_project_bin)?;
    out.extend_from_slice(&forms);
    Ok(out)
}

fn normalize_project_stream_properties_v3(project_stream_bytes: &[u8]) -> Vec<u8> {
    // The `PROJECT` stream is line-oriented ASCII/MBCS text. We normalize by:
    // - splitting on NWLN (CRLF or LFCR; with tolerance for lone CR/LF),
    // - filtering out excluded property keys (case-insensitive),
    // - ignoring the optional `[Workspace]` (and any later) section, and
    // - emitting each included raw `key=value` line terminated with CRLF.
    //
    // Note: we intentionally operate on bytes rather than decoding to UTF-8 so that the
    // transcript preserves MBCS bytes verbatim.

fn is_excluded_key(key: &[u8]) -> bool {
        // Exclusions used by the v3 `PROJECT` stream transcript (case-insensitive).
        key.eq_ignore_ascii_case(b"ID")
            || key.eq_ignore_ascii_case(b"Document")
            || key.eq_ignore_ascii_case(b"DocModule")
            || key.eq_ignore_ascii_case(b"CMG")
            || key.eq_ignore_ascii_case(b"DPB")
            || key.eq_ignore_ascii_case(b"GC")
            || key.eq_ignore_ascii_case(b"ProtectionState")
            || key.eq_ignore_ascii_case(b"Password")
            || key.eq_ignore_ascii_case(b"VisibilityState")
    }

    let mut out = Vec::new();
    let mut saw_host_extender_info = false;

    for raw_line in split_nwln_lines(project_stream_bytes) {
        let mut line = trim_ascii_whitespace(raw_line);
        if line.is_empty() {
            continue;
        }

        // Some writers may include a UTF-8 BOM at the start of the stream. Strip it for key
        // matching and output stability.
        if line.starts_with(&[0xEF, 0xBB, 0xBF]) {
            line = trim_ascii_whitespace(&line[3..]);
            if line.is_empty() {
                continue;
            }
        }

        // Section headers are bracketed, e.g. `[Host Extender Info]` or `[Workspace]`.
        if line.starts_with(b"[") && line.ends_with(b"]") {
            if line.eq_ignore_ascii_case(b"[Host Extender Info]") {
                saw_host_extender_info = true;
                continue;
            }
            // Any other section (including `[Workspace]`) is ignored for signature binding.
            break;
        }

        let Some(eq) = line.iter().position(|&b| b == b'=') else {
            continue;
        };

        let key = trim_ascii_whitespace(&line[..eq]);
        if key.is_empty() {
            continue;
        }
        if is_excluded_key(key) {
            continue;
        }

        // For `[Host Extender Info]` we currently preserve the raw `key=value` bytes as written (the
        // same policy as the main ProjectProperties block). This is intentionally conservative: it
        // matches the behavior of the rest of the v3 project transcript, which aims to avoid
        // lossy decoding and avoids attempting to interpret unknown host-extender fields.
        //
        // If we haven't seen the Host Extender section, `saw_host_extender_info` is false, but the
        // behavior is identical; the flag is kept for future-proofing and readability.
        let _ = saw_host_extender_info;

        out.extend_from_slice(line);
        out.extend_from_slice(b"\r\n");
    }

    out
}

// NOTE: `normalize_project_stream_properties_v3` intentionally does not share the
// `project_properties_normalized_bytes` / `host_extender_info_normalized_bytes` logic above. Those
// helpers implement a different (more structured) `PROJECT`-stream normalization, whereas the
// current v3 binding transcript preserves filtered raw `key=value` lines with CRLF-normalized
// endings.
//
/// Compute a SHA-256 digest over [`project_normalized_data_v3_transcript`]'s transcript.
///
/// This is a convenience/helper API: real-world `\x05DigitalSignatureExt` signatures most commonly
/// use a 32-byte SHA-256 binding digest, but MS-OVBA v3 is defined in terms of hashing a
/// `ContentBuffer` (`V3ContentNormalizedData || ProjectNormalizedData`) and producers may vary.
///
/// For other algorithms (debugging/out-of-spec), callers can use
/// [`crate::compute_vba_project_digest_v3`].
pub fn contents_hash_v3(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    let normalized = project_normalized_data_v3_transcript(vba_project_bin)?;
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
    //
    // Note: in spec-compliant `VBA/dir` streams, `SizeOfLibidOriginal` is the record's `len` field
    // (i.e., it is not repeated inside the record payload). Many synthetic fixtures in our test
    // suite use a simplified TLV-like encoding where the libid is itself u32-len-prefixed inside
    // the record data. Support both encodings for robustness.
    if data.len() >= 4 {
        let len = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
        if let Some(end) = 4usize.checked_add(len) {
            if let Some(payload) = data.get(4..end) {
                return Ok(copy_until_nul(payload));
            }
        }
    }

    Ok(copy_until_nul(data))
}

fn read_u32_len_prefixed_bytes(cur: &mut &[u8]) -> Result<Vec<u8>, ParseError> {
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

fn append_v3_normalized_module_source(source: &[u8], module_name: &[u8], out: &mut Vec<u8>) {
    // MS-OVBA §2.4.2.5 `HashModuleNameFlag` is set when at least one line contributes.
    let mut hash_module_name_flag = false;

    let mut text_buffer: Vec<u8> = Vec::new();
    let mut previous_char: u8 = 0;

    for &ch in source {
        if ch == b'\r' || (ch == b'\n' && previous_char != b'\r') {
            hash_module_name_flag |= append_v3_line(&text_buffer, out);
            text_buffer.clear();
            previous_char = ch;
            continue;
        }

        if ch == b'\n' && previous_char == b'\r' {
            // Ignore the LF of CRLF; keep PreviousChar as 0x0D (matches MS-OVBA pseudocode).
            continue;
        }

        text_buffer.push(ch);
        previous_char = ch;
    }

    // Process the trailing (possibly empty) line.
    hash_module_name_flag |= append_v3_line(&text_buffer, out);

    if hash_module_name_flag {
        out.extend_from_slice(module_name);
        out.push(b'\n');
    }
}

fn append_v3_line(line: &[u8], out: &mut Vec<u8>) -> bool {
    let is_attribute = starts_with_ignore_ascii_case(line, b"attribute");
    if !is_attribute {
        out.extend_from_slice(line);
        out.push(b'\n');
        return true;
    }

    // Skip `Attribute VB_Name = ...` lines.
    if starts_with_ignore_ascii_case(line, V3_VB_NAME_PREFIX) {
        return false;
    }

    // Skip DefaultAttributes lines by exact byte equality (case-sensitive).
    if V3_DEFAULT_ATTRIBUTES.contains(&line) {
        return false;
    }

    out.extend_from_slice(line);
    out.push(b'\n');
    true
}

fn normalize_module_source(bytes: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    let _ = out.try_reserve_exact(bytes.len());

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

fn starts_with_ignore_ascii_case(haystack: &[u8], needle: &[u8]) -> bool {
    if haystack.len() < needle.len() {
        return false;
    }
    haystack
        .iter()
        .take(needle.len())
        .zip(needle.iter())
        .all(|(a, b)| a.eq_ignore_ascii_case(b))
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

fn decode_dir_unicode_string(bytes: &[u8]) -> String {
    let bytes = trim_u32_len_prefix_unicode_string(bytes);
    let (cow, _) = UTF_16LE.decode_without_bom_handling(bytes);
    let mut s = cow.into_owned();
    // Defensively strip embedded NULs; stream/module names should not contain them.
    s.retain(|c| c != '\u{0000}');
    s
}

fn trim_u32_len_prefix_unicode_string(bytes: &[u8]) -> &[u8] {
    // Some MS-OVBA `*_UNICODE` record payloads are specified/observed as an optional u32 length
    // prefix followed by UTF-16LE bytes.
    //
    // Heuristics:
    // - if the first u32 equals the remaining byte count, treat it as a byte-length prefix.
    // - if the first u32 equals the remaining UTF-16 code unit count, treat it as a char-length prefix.
    fn trim_trailing_utf16_nul(bytes: &[u8]) -> &[u8] {
        if bytes.len() >= 2 && bytes.len().is_multiple_of(2) && bytes.ends_with(&[0x00, 0x00]) {
            &bytes[..bytes.len() - 2]
        } else {
            bytes
        }
    }

    if bytes.len() < 4 {
        return trim_trailing_utf16_nul(bytes);
    }

    let n = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]) as usize;
    let rest = &bytes[4..];

    // Treat the leading u32 as a length prefix only when it is consistent with the remaining bytes.
    let mut out = if n == rest.len() || (rest.len().is_multiple_of(2) && n.saturating_mul(2) == rest.len())
    {
        rest
    } else if rest.len() >= 2
        && rest.ends_with(&[0x00, 0x00])
        && (n.saturating_add(2) == rest.len()
            || (rest.len().is_multiple_of(2) && n.saturating_mul(2).saturating_add(2) == rest.len()))
    {
        // Some producers include a trailing UTF-16 NUL terminator but do not count it in the
        // internal length prefix. Accept that form too, but only when an actual terminator is
        // present (to avoid misclassifying random leading u32 values).
        &rest[..rest.len() - 2]
    } else {
        bytes
    };

    // Some producers include a trailing UTF-16 NUL terminator regardless of whether it is counted
    // by the internal length prefix. Since stream/module names should not contain embedded NULs,
    // strip a single trailing terminator to keep hashing/transcripts stable across producers.
    out = trim_trailing_utf16_nul(out);
    out
}

fn looks_like_utf16le(bytes: &[u8]) -> bool {
    if bytes.len() < 2 || !bytes.len().is_multiple_of(2) {
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
