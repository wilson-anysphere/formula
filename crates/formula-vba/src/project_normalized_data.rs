use crate::{decompress_container, DirParseError, OleFile, ParseError};

#[derive(Debug, Default, Clone, Copy)]
struct ProjectUnicodePresence {
    project_name: bool,
    project_docstring: bool,
    project_help_filepath: bool,
    project_constants: bool,
}

#[derive(Debug, Default, Clone, Copy)]
struct ModuleUnicodePresence {
    module_name: bool,
    module_stream_name: bool,
    module_docstring: bool,
    module_help_filepath: bool,
}

/// Build a dir-record-only v3 project/module metadata transcript derived from selected `VBA/dir`
/// records (MS-OVBA §2.4.2.6).
///
/// This helper is useful for debugging/spec work, but it does **not** include:
/// - v3 module source normalization (`V3ContentNormalizedData`), or
/// - designer storage bytes (`FormsNormalizedData`).
///
/// For the transcript used by `ContentsHashV3` / `\x05DigitalSignatureExt` binding, use
/// [`crate::project_normalized_data_v3`].
///
/// Notes / invariants enforced for hashing:
/// - Records are processed in the **stored order** from the decompressed `VBA/dir` stream.
/// - The record header (`id` + `len`) is **not** included in the output; only normalized record
///   *data* is concatenated.
/// - For string-like fields that have both ANSI/MBCS and Unicode record variants, the Unicode
///   record is preferred when present (and the ANSI record is omitted).
pub fn project_normalized_data_v3_dir_records(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
    let mut ole = OleFile::open(vba_project_bin)?;

    let dir_bytes = ole
        .read_stream_opt("VBA/dir")?
        .ok_or(ParseError::MissingStream("VBA/dir"))?;
    let dir_decompressed = decompress_container(&dir_bytes)?;

    // First pass: identify which Unicode record variants exist so we can omit the ANSI record
    // variants (MS-OVBA V3 behavior) without having to buffer/rewind output.
    let (project_unicode, modules_unicode) = scan_unicode_presence(&dir_decompressed)?;

    // Second pass: concatenate normalized record data in stored order.
    let mut out = Vec::new();

    let mut offset = 0usize;
    let mut module_idx = 0usize;
    let mut current_module_unicode = ModuleUnicodePresence::default();
    while offset < dir_decompressed.len() {
        let (id, data, next_offset) = read_dir_record(&dir_decompressed, offset)?;
        offset = next_offset;

        match id {
            // ---------------------------------------------------------------------
            // Project-level records (metadata)
            // ---------------------------------------------------------------------
            // PROJECTSYSKIND, PROJECTLCID, PROJECTCODEPAGE, PROJECTHELPCONTEXT,
            // PROJECTLIBFLAGS, PROJECTVERSION.
            0x0001 | 0x0002 | 0x0003 | 0x0007 | 0x0008 | 0x0009 => {
                out.extend_from_slice(data);
            }

            // PROJECTNAME (ANSI) vs PROJECTNAMEUNICODE (Unicode)
            0x0004 if !project_unicode.project_name => {
                out.extend_from_slice(data);
            }
            // PROJECTDOCSTRING (ANSI) vs PROJECTDOCSTRINGUNICODE (Unicode)
            0x0005 if !project_unicode.project_docstring => {
                out.extend_from_slice(data);
            }
            // PROJECTHELPFILEPATH (ANSI) vs PROJECTHELPFILEPATHUNICODE (Unicode)
            0x0006 if !project_unicode.project_help_filepath => {
                out.extend_from_slice(data);
            }
            // PROJECTCONSTANTS (ANSI) vs PROJECTCONSTANTSUNICODE (Unicode)
            0x000C if !project_unicode.project_constants => {
                out.extend_from_slice(data);
            }

            // Unicode project string record variants: normalize by dropping the internal length
            // prefix (u32) and appending only the UTF-16LE payload bytes.
            0x0040 | 0x0041 | 0x0042 | 0x0043 => {
                out.extend_from_slice(unicode_record_payload(data)?);
            }

            // ---------------------------------------------------------------------
            // Module-level records (metadata)
            // ---------------------------------------------------------------------
            // MODULENAME: also establishes the current module record group.
            0x0019 => {
                current_module_unicode = modules_unicode
                    .get(module_idx)
                    .copied()
                    .unwrap_or_default();
                module_idx = module_idx.saturating_add(1);

                if !current_module_unicode.module_name {
                    out.extend_from_slice(data);
                }
            }

            // MODULESTREAMNAME (ANSI) vs MODULESTREAMNAMEUNICODE (Unicode).
            //
            // For ANSI, strip a trailing reserved u16 when present (commonly `0x0000`).
            0x001A if !current_module_unicode.module_stream_name => {
                out.extend_from_slice(trim_reserved_u16(data));
            }

            // MODULEDOCSTRING (ANSI) vs MODULEDOCSTRINGUNICODE (Unicode)
            0x001B if !current_module_unicode.module_docstring => {
                out.extend_from_slice(data);
            }

            // MODULEHELPFILEPATH (ANSI) vs MODULEHELPFILEPATHUNICODE (Unicode)
            0x001D if !current_module_unicode.module_help_filepath => {
                out.extend_from_slice(data);
            }

            // Non-string module metadata records included in V3.
            0x001E | 0x0021 | 0x0025 | 0x0028 => {
                out.extend_from_slice(data);
            }

            // Unicode module string record variants.
            0x0047 | 0x0048 | 0x0049 => {
                out.extend_from_slice(unicode_record_payload(data)?);
            }

            // 0x004A is an unfortunate ID collision between:
            // - PROJECTCOMPATVERSION (project-level, optional), and
            // - MODULEHELPFILEPATHUNICODE (module-level Unicode string variant).
            //
            // For our TLV-style dir-record fixtures, PROJECTCOMPATVERSION appears in the project
            // information section before any module record groups; it MUST be ignored for hashing.
            //
            // If we've already encountered a module group, treat it as the module Unicode variant.
            0x004A => {
                if module_idx != 0 {
                    out.extend_from_slice(unicode_record_payload(data)?);
                }
            }

            // All other records (references, module offsets, etc.) are excluded from V3
            // ProjectNormalizedData.
            _ => {}
        }
    }

    Ok(out)
}

fn scan_unicode_presence(
    dir_decompressed: &[u8],
) -> Result<(ProjectUnicodePresence, Vec<ModuleUnicodePresence>), ParseError> {
    let mut project_unicode = ProjectUnicodePresence::default();
    let mut modules_unicode: Vec<ModuleUnicodePresence> = Vec::new();
    let mut current_module: Option<usize> = None;

    let mut offset = 0usize;
    while offset < dir_decompressed.len() {
        let (id, _data, next_offset) = read_dir_record(dir_decompressed, offset)?;
        offset = next_offset;

        match id {
            // New module record group starts at MODULENAME.
            0x0019 => {
                current_module = Some(modules_unicode.len());
                modules_unicode.push(ModuleUnicodePresence::default());
            }

            // Project-level Unicode string variants.
            0x0040 => project_unicode.project_name = true,
            0x0041 => project_unicode.project_docstring = true,
            0x0042 => project_unicode.project_help_filepath = true,
            0x0043 => project_unicode.project_constants = true,

            // Module-level Unicode string variants.
            0x0047 => {
                if let Some(idx) = current_module {
                    modules_unicode[idx].module_name = true;
                }
            }
            0x0048 => {
                if let Some(idx) = current_module {
                    modules_unicode[idx].module_stream_name = true;
                }
            }
            0x0049 => {
                if let Some(idx) = current_module {
                    modules_unicode[idx].module_docstring = true;
                }
            }
            0x004A => {
                if let Some(idx) = current_module {
                    modules_unicode[idx].module_help_filepath = true;
                }
            }

            _ => {}
        }
    }

    Ok((project_unicode, modules_unicode))
}

fn read_dir_record<'a>(
    buf: &'a [u8],
    offset: usize,
) -> Result<(u16, &'a [u8], usize), DirParseError> {
    if offset + 6 > buf.len() {
        return Err(DirParseError::Truncated);
    }
    let id = u16::from_le_bytes([buf[offset], buf[offset + 1]]);
    let len = u32::from_le_bytes([
        buf[offset + 2],
        buf[offset + 3],
        buf[offset + 4],
        buf[offset + 5],
    ]) as usize;
    let data_start = offset + 6;
    let data_end = data_start + len;
    if data_end > buf.len() {
        return Err(DirParseError::BadRecordLength { id, len });
    }
    Ok((id, &buf[data_start..data_end], data_end))
}

fn trim_reserved_u16(bytes: &[u8]) -> &[u8] {
    if bytes.len() >= 2 && bytes[bytes.len() - 2..] == [0x00, 0x00] {
        &bytes[..bytes.len() - 2]
    } else {
        bytes
    }
}

fn unicode_record_payload(data: &[u8]) -> Result<&[u8], DirParseError> {
    // MS-OVBA Unicode string record payloads contain a u32 length prefix followed by UTF-16LE data.
    // The length is commonly a UTF-16 code unit count, but some producers treat it as a byte count;
    // accept both as long as bounds are valid (deterministic choice: prefer the interpretation that
    // exactly consumes the record, otherwise prefer the spec's UTF-16 code unit count).
    if data.len() < 4 {
        return Err(DirParseError::Truncated);
    }

    let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
    let remaining = data.len() - 4;

    let bytes_by_units = n.checked_mul(2);

    // Prefer interpretations that exactly match the record length.
    if let Some(bytes) = bytes_by_units {
        if bytes == remaining {
            return Ok(&data[4..4 + bytes]);
        }
    }
    if n == remaining {
        return Ok(&data[4..4 + n]);
    }

    // Otherwise, prefer UTF-16 code unit count when it fits; fall back to byte count.
    if let Some(bytes) = bytes_by_units {
        if bytes <= remaining {
            return Ok(&data[4..4 + bytes]);
        }
    }
    if n <= remaining {
        return Ok(&data[4..4 + n]);
    }

    Err(DirParseError::Truncated)
}
