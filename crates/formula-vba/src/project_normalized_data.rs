use crate::{decompress_container, DirParseError, OleFile, ParseError};

#[derive(Debug, Default, Clone, Copy)]
struct ProjectUnicodePresence {
    project_docstring: bool,
    project_help_filepath: bool,
    project_constants: bool,
}

#[derive(Debug, Default, Clone, Copy)]
struct ModuleUnicodePresence {
    module_name: bool,
    module_stream_name: bool,
    module_docstring: bool,
}

/// Build a dir-record-only v3 project/module metadata transcript derived from selected `VBA/dir`
/// records (MS-OVBA §2.4.2.6).
///
/// This helper is useful for debugging/spec work, but it does **not** include:
/// - v3 module source normalization (`V3ContentNormalizedData`), or
/// - designer storage bytes (`FormsNormalizedData`).
///
/// For the v3 signature binding transcript used by `ContentsHashV3` / `\x05DigitalSignatureExt`,
/// use [`crate::project_normalized_data_v3_transcript`].
///
/// For legacy/test helpers and building blocks, see:
/// - [`crate::project_normalized_data`] (MS-OVBA `ProjectNormalizedData`), and
/// - [`crate::v3_content_normalized_data`] (MS-OVBA `V3ContentNormalizedData`).
///
/// Notes / invariants enforced for hashing:
/// - Records are processed in the **stored order** from the decompressed `VBA/dir` stream.
/// - The record header (`id` + `len`) is **not** included in the output; only normalized record
///   *data* is concatenated.
/// - For string-like fields that have both ANSI/MBCS and Unicode record variants, the Unicode
///   record is preferred when present (and the ANSI record is omitted).
pub fn project_normalized_data_v3(vba_project_bin: &[u8]) -> Result<Vec<u8>, ParseError> {
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
    // Tracks whether we've consumed any non-name records for the current module record group.
    // Used to disambiguate MODULENAMEUNICODE when MODULENAME is absent (non-spec, but observed in
    // some TLV-ish dir streams).
    let mut in_module = false;
    let mut module_seen_non_name_record = false;
    while offset < dir_decompressed.len() {
        let (id, data, next_offset) = read_dir_record(&dir_decompressed, offset)?;
        offset = next_offset;

        match id {
            // ---------------------------------------------------------------------
            // Project-level records (metadata)
            // ---------------------------------------------------------------------
            // PROJECTSYSKIND, PROJECTLCID, PROJECTCODEPAGE, PROJECTHELPCONTEXT,
            // PROJECTLIBFLAGS, PROJECTVERSION.
            0x0001 | 0x0002 | 0x0003 | 0x0007 | 0x0008 | 0x0009 | 0x0014 => {
                out.extend_from_slice(data);
            }

            // PROJECTNAME (ANSI)
            //
            // Note: this repo's hashing implementation does not currently special-case a Unicode
            // variant for PROJECTNAME.
            0x0004 => out.extend_from_slice(data),
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

            // Unicode/alternate project string record variants: normalize by dropping an internal
            // u32 length prefix when present, and appending only the UTF-16LE payload bytes.
            //
            // Observed record IDs:
            // - 0x0040: PROJECTDOCSTRING (Unicode)
            // - 0x003D: PROJECTHELPFILEPATH (second form / commonly Unicode)
            // - 0x003C: PROJECTCONSTANTS (Unicode)
            0x0040 | 0x003D | 0x003C => {
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
                in_module = true;
                module_seen_non_name_record = false;

                if !current_module_unicode.module_name {
                    out.extend_from_slice(data);
                }
            }

            // MODULENAMEUNICODE (Unicode).
            //
            // In spec-compliant streams this usually appears immediately after MODULENAME and is an
            // alternate representation of the current module's name. Some producers omit MODULENAME
            // entirely and emit only MODULENAMEUNICODE; when that happens, treat it as the start of
            // a module record group so that Unicode-vs-ANSI selection for subsequent records works.
            0x0047 => {
                let start_new = !in_module || module_seen_non_name_record;
                if start_new {
                    current_module_unicode = modules_unicode
                        .get(module_idx)
                        .copied()
                        .unwrap_or_default();
                    module_idx = module_idx.saturating_add(1);
                    in_module = true;
                    module_seen_non_name_record = false;
                }
                out.extend_from_slice(unicode_record_payload(data)?);
            }

            // MODULESTREAMNAME (ANSI) vs MODULESTREAMNAMEUNICODE (Unicode).
            //
            // For ANSI, strip a trailing reserved u16 when present (commonly `0x0000`).
            0x001A => {
                if !current_module_unicode.module_stream_name {
                    out.extend_from_slice(trim_reserved_u16(data));
                }
                if in_module {
                    module_seen_non_name_record = true;
                }
            }

            // MODULEDOCSTRING (ANSI) vs MODULEDOCSTRINGUNICODE (Unicode)
            0x001C => {
                if !current_module_unicode.module_docstring {
                    out.extend_from_slice(data);
                }
                if in_module {
                    module_seen_non_name_record = true;
                }
            }

            // Non-string module metadata records included in V3.
            0x001E | 0x0021 | 0x0025 | 0x0028 => {
                out.extend_from_slice(data);
                if in_module {
                    module_seen_non_name_record = true;
                }
            }

            // Unicode module string record variants.
            //
            // Observed record IDs:
            // - 0x0032: MODULESTREAMNAME (Unicode)
            // - 0x0048: MODULEDOCSTRING (Unicode)
            0x0032 | 0x0048 => {
                out.extend_from_slice(unicode_record_payload(data)?);
                if in_module {
                    module_seen_non_name_record = true;
                }
            }

            // All other records (references, module offsets, etc.) are excluded from V3
            // ProjectNormalizedData.
            _ => {}
        }
    }

    Ok(out)
}

/// Backwards-compatible alias for [`project_normalized_data_v3`].
pub fn project_normalized_data_v3_dir_records(
    vba_project_bin: &[u8],
) -> Result<Vec<u8>, ParseError> {
    project_normalized_data_v3(vba_project_bin)
}

fn scan_unicode_presence(
    dir_decompressed: &[u8],
) -> Result<(ProjectUnicodePresence, Vec<ModuleUnicodePresence>), ParseError> {
    let mut project_unicode = ProjectUnicodePresence::default();
    let mut modules_unicode: Vec<ModuleUnicodePresence> = Vec::new();
    let mut current_module: Option<usize> = None;
    let mut current_module_seen_non_name_record = false;

    let mut offset = 0usize;
    while offset < dir_decompressed.len() {
        let (id, _data, next_offset) = read_dir_record(dir_decompressed, offset)?;
        offset = next_offset;

        match id {
            // New module record group starts at MODULENAME.
            0x0019 => {
                current_module = Some(modules_unicode.len());
                modules_unicode.push(ModuleUnicodePresence::default());
                current_module_seen_non_name_record = false;
            }

            // Project-level Unicode/alternate string variants.
            0x0040 => project_unicode.project_docstring = true,
            0x003D => project_unicode.project_help_filepath = true,
            0x003C => project_unicode.project_constants = true,

            // Module-level Unicode string variants.
            0x0047 => {
                let start_new = current_module.is_none() || current_module_seen_non_name_record;
                if start_new {
                    current_module = Some(modules_unicode.len());
                    let mut presence = ModuleUnicodePresence::default();
                    presence.module_name = true;
                    modules_unicode.push(presence);
                    current_module_seen_non_name_record = false;
                } else if let Some(idx) = current_module {
                    modules_unicode[idx].module_name = true;
                }
            }
            0x0032 => {
                if let Some(idx) = current_module {
                    modules_unicode[idx].module_stream_name = true;
                    current_module_seen_non_name_record = true;
                }
            }
            0x0048 => {
                if let Some(idx) = current_module {
                    modules_unicode[idx].module_docstring = true;
                    current_module_seen_non_name_record = true;
                }
            }
            // Any non-name module record helps disambiguate MODULENAMEUNICODE as either an alternate
            // representation (immediately after MODULENAME) or the start of a new module group.
            0x001A | 0x001C | 0x001E | 0x0021 | 0x0025 | 0x0028 | 0x002B | 0x002C | 0x0031 => {
                if current_module.is_some() {
                    current_module_seen_non_name_record = true;
                }
            }

            _ => {}
        }
    }

    Ok((project_unicode, modules_unicode))
}

fn read_dir_record(buf: &[u8], offset: usize) -> Result<(u16, &[u8], usize), DirParseError> {
    if offset + 6 > buf.len() {
        return Err(DirParseError::Truncated);
    }
    let id = u16::from_le_bytes([buf[offset], buf[offset + 1]]);

    // MS-OVBA `VBA/dir` streams are usually encoded as `Id(u16) || Size(u32) || Data(Size)`.
    // However, some fixed-length records (notably PROJECTVERSION) are stored without an explicit
    // `Size` field in real-world projects.
    //
    // PROJECTVERSION (0x0009) record layout:
    //   Id(u16) || Reserved(u32) || VersionMajor(u32) || VersionMinor(u16)
    // (12 bytes total, with 10 bytes of "data" after the Id).
    //
    // For v3 transcripts, we want to be able to scan past this record even when it uses the
    // spec-compliant fixed-length form.
    if id == 0x0009 {
        let reserved_or_size = u32::from_le_bytes([
            buf[offset + 2],
            buf[offset + 3],
            buf[offset + 4],
            buf[offset + 5],
        ]) as usize;
        if reserved_or_size == 0 {
            let end = offset + 12;
            if end > buf.len() {
                return Err(DirParseError::Truncated);
            }
            // Return the record "data" bytes following the Id.
            return Ok((id, &buf[offset + 2..end], end));
        }
    }

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
    // MS-OVBA "Unicode" record payloads are commonly UTF-16LE bytes.
    //
    // Some producers also embed an *internal* u32 length prefix before the UTF-16LE bytes
    // (length can be either code units or byte count). Accept both:
    // - raw UTF-16LE bytes (no internal prefix), and
    // - `u32 length || utf16le_bytes`.
    if data.len() < 4 {
        return Ok(data);
    }

    let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
    let remaining = data.len() - 4;

    // Treat the leading u32 as a length prefix only when it is consistent with the remaining bytes.
    if n == remaining || n.saturating_mul(2) == remaining {
        return Ok(&data[4..]);
    }

    // Otherwise assume the record is raw UTF-16LE bytes.
    Ok(data)
}
