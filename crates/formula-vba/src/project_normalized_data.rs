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

/// Build a dir-record allowlist transcript of project/module metadata.
///
/// This helper concatenates selected `VBA/dir` record payload bytes in stored order and applies the
/// ANSI-vs-Unicode selection rules needed for stable hashing.
///
/// Important: despite the historical naming, this is **not** MS-OVBA §2.4.2.6
/// `ProjectNormalizedData` (`NormalizeProjectStream`), which is derived from the textual `PROJECT`
/// stream.
///
/// For `formula-vba`'s current v3 signature-binding transcript, use
/// [`crate::project_normalized_data_v3_transcript`].
///
/// For spec-driven v3 building blocks, see [`crate::v3_content_normalized_data`].
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
    // Some producers appear to use 0x0049 as the Unicode marker/record for MODULEDOCSTRING rather
    // than the canonical 0x0048. Track whether we've just seen MODULEDOCSTRING so we can
    // distinguish that from MODULEHELPFILEPATHUNICODE (0x0049), which this helper intentionally
    // excludes.
    let mut expect_module_docstring_unicode = false;
    while offset < dir_decompressed.len() {
        let (id, data, next_offset) = read_dir_record(&dir_decompressed, offset)?;
        offset = next_offset;

        if expect_module_docstring_unicode && !matches!(id, 0x0048 | 0x0049) {
            expect_module_docstring_unicode = false;
        }

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
            // - 0x0040: PROJECTDOCSTRING (Unicode; canonical)
            // - 0x0041: PROJECTDOCSTRING (Unicode; non-canonical variant)
            // - 0x003D: PROJECTHELPFILEPATH (second form / commonly Unicode; canonical)
            // - 0x0042: PROJECTHELPFILEPATH (non-canonical variant)
            // - 0x003C: PROJECTCONSTANTS (Unicode; canonical)
            // - 0x0043: PROJECTCONSTANTS (Unicode; non-canonical variant)
            0x0040 | 0x0041 | 0x003D | 0x0042 | 0x003C | 0x0043 => {
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
            //
            // Some producers use record id 0x001B instead of 0x001C for the ANSI form; accept both.
            0x001B | 0x001C => {
                if !current_module_unicode.module_docstring {
                    out.extend_from_slice(data);
                }
                if in_module {
                    module_seen_non_name_record = true;
                }
                expect_module_docstring_unicode = true;
            }

            // Non-string module metadata records included in V3.
            0x001E | 0x0021 | 0x0022 | 0x0025 | 0x0028 => {
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
                expect_module_docstring_unicode = false;
            }
            // Some producers use 0x0049 as the MODULEDOCSTRING Unicode marker/record. Only treat it
            // as a docstring variant when it follows MODULEDOCSTRING; otherwise it is more likely
            // MODULEHELPFILEPATHUNICODE (which this helper intentionally excludes).
            0x0049 if expect_module_docstring_unicode => {
                out.extend_from_slice(unicode_record_payload(data)?);
                if in_module {
                    module_seen_non_name_record = true;
                }
                expect_module_docstring_unicode = false;
            }

            // All other records (references, module offsets, etc.) are excluded from this helper's
            // output.
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
    // Some TLV-ish `VBA/dir` layouts store a Unicode module stream name as a separate record
    // immediately following MODULESTREAMNAME (0x001A). In most files this record id is 0x0032, but
    // some real-world projects reuse 0x0048. Track this expectation so we can disambiguate 0x0048
    // as either MODULESTREAMNAMEUNICODE or MODULEDOCSTRINGUNICODE depending on context.
    let mut expect_module_stream_name_unicode = false;
    // Some producers appear to use 0x0049 as the Unicode marker/record for MODULEDOCSTRING rather
    // than the canonical 0x0048. Track whether we've just seen MODULEDOCSTRING so we can
    // distinguish that from MODULEHELPFILEPATHUNICODE (0x0049).
    let mut expect_module_docstring_unicode = false;

    let mut offset = 0usize;
    while offset < dir_decompressed.len() {
        let (id, _data, next_offset) = read_dir_record(dir_decompressed, offset)?;
        offset = next_offset;

        if expect_module_stream_name_unicode && !matches!(id, 0x0032 | 0x0048) {
            expect_module_stream_name_unicode = false;
        }
        if expect_module_docstring_unicode && !matches!(id, 0x0048 | 0x0049) {
            expect_module_docstring_unicode = false;
        }

        match id {
            // New module record group starts at MODULENAME.
            0x0019 => {
                current_module = Some(modules_unicode.len());
                modules_unicode.push(ModuleUnicodePresence::default());
                current_module_seen_non_name_record = false;
                expect_module_stream_name_unicode = false;
                expect_module_docstring_unicode = false;
            }

            // Project-level Unicode/alternate string variants.
            0x0040 | 0x0041 => project_unicode.project_docstring = true,
            0x003D | 0x0042 => project_unicode.project_help_filepath = true,
            0x003C | 0x0043 => project_unicode.project_constants = true,

            // Module-level Unicode string variants.
            0x0047 => {
                let start_new = current_module.is_none() || current_module_seen_non_name_record;
                if start_new {
                    current_module = Some(modules_unicode.len());
                    modules_unicode.push(ModuleUnicodePresence {
                        module_name: true,
                        ..Default::default()
                    });
                    current_module_seen_non_name_record = false;
                    expect_module_stream_name_unicode = false;
                    expect_module_docstring_unicode = false;
                } else if let Some(idx) = current_module {
                    modules_unicode[idx].module_name = true;
                }
            }
            0x001A => {
                if current_module.is_some() {
                    current_module_seen_non_name_record = true;
                    // Expect a Unicode stream name record to follow (0x0032 in many encodings, but
                    // some files reuse 0x0048).
                    expect_module_stream_name_unicode = true;
                } else {
                    expect_module_stream_name_unicode = false;
                }
                expect_module_docstring_unicode = false;
            }
            0x001B | 0x001C => {
                if current_module.is_some() {
                    current_module_seen_non_name_record = true;
                }
                expect_module_docstring_unicode = true;
            }
            0x0032 => {
                if let Some(idx) = current_module {
                    modules_unicode[idx].module_stream_name = true;
                    current_module_seen_non_name_record = true;
                }
                expect_module_stream_name_unicode = false;
                expect_module_docstring_unicode = false;
            }
            0x0048 => {
                if let Some(idx) = current_module {
                    current_module_seen_non_name_record = true;
                    if expect_module_stream_name_unicode {
                        // Some producers use 0x0048 as the module stream-name Unicode record id.
                        modules_unicode[idx].module_stream_name = true;
                    } else {
                        modules_unicode[idx].module_docstring = true;
                    }
                }
                expect_module_stream_name_unicode = false;
                expect_module_docstring_unicode = false;
            }
            0x0049 => {
                if let Some(idx) = current_module {
                    current_module_seen_non_name_record = true;
                    if expect_module_docstring_unicode {
                        modules_unicode[idx].module_docstring = true;
                    }
                }
                expect_module_stream_name_unicode = false;
                expect_module_docstring_unicode = false;
            }
            // Any non-name module record helps disambiguate MODULENAMEUNICODE as either an alternate
            // representation (immediately after MODULENAME) or the start of a new module group.
            0x001E | 0x0021 | 0x0022 | 0x0025 | 0x0028 | 0x002B | 0x002C | 0x0031 => {
                if current_module.is_some() {
                    current_module_seen_non_name_record = true;
                }
                expect_module_stream_name_unicode = false;
                expect_module_docstring_unicode = false;
            }

            _ => {}
        }
    }

    Ok((project_unicode, modules_unicode))
}

fn read_dir_record(buf: &[u8], offset: usize) -> Result<(u16, &[u8], usize), DirParseError> {
    let Some(hdr_end) = offset.checked_add(6) else {
        return Err(DirParseError::Truncated);
    };
    let Some(hdr) = buf.get(offset..hdr_end) else {
        return Err(DirParseError::Truncated);
    };
    let id = u16::from_le_bytes([hdr[0], hdr[1]]);

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
        // PROJECTVERSION (0x0009) is fixed-length in MS-OVBA, but many fixtures (and some
        // producers) encode it as a normal TLV record (`Id || Size || Data`).
        //
        // Disambiguate by checking which interpretation yields a plausible next record header.
        let size_or_reserved = u32::from_le_bytes([hdr[2], hdr[3], hdr[4], hdr[5]]) as usize;

        let tlv_end = offset.saturating_add(6).saturating_add(size_or_reserved);
        let fixed_end = offset.saturating_add(12);

        let tlv_next_ok = looks_like_projectversion_following_record(buf, tlv_end);
        let fixed_next_ok = looks_like_projectversion_following_record(buf, fixed_end);

        // Prefer the fixed-length interpretation when the TLV interpretation would leave us at an
        // implausible record boundary, or when the `Size` field is `0` (which is a common reserved
        // value in fixed-length PROJECTVERSION records).
        if fixed_end <= buf.len() && fixed_next_ok && (!tlv_next_ok || size_or_reserved == 0) {
            let start = offset.checked_add(2).ok_or(DirParseError::Truncated)?;
            let data = buf.get(start..fixed_end).ok_or(DirParseError::Truncated)?;
            return Ok((id, data, fixed_end));
        }

        // Fall back to treating it as a TLV record (use the declared payload length).
        let data_start = offset.checked_add(6).ok_or(DirParseError::Truncated)?;
        let data_end = data_start
            .checked_add(size_or_reserved)
            .ok_or(DirParseError::BadRecordLength {
                id,
                len: size_or_reserved,
            })?;
        if data_end > buf.len() {
            return Err(DirParseError::BadRecordLength {
                id,
                len: size_or_reserved,
            });
        }
        let data = buf.get(data_start..data_end).ok_or(DirParseError::Truncated)?;
        return Ok((id, data, data_end));
    }

    let len = u32::from_le_bytes([hdr[2], hdr[3], hdr[4], hdr[5]]) as usize;
    let data_start = offset.checked_add(6).ok_or(DirParseError::Truncated)?;
    let data_end = data_start
        .checked_add(len)
        .ok_or(DirParseError::BadRecordLength { id, len })?;
    if data_end > buf.len() {
        return Err(DirParseError::BadRecordLength { id, len });
    }
    let data = buf.get(data_start..data_end).ok_or(DirParseError::Truncated)?;
    Ok((id, data, data_end))
}

fn looks_like_projectversion_following_record(bytes: &[u8], offset: usize) -> bool {
    if offset == bytes.len() {
        return true;
    }
    let Some(hdr_end) = offset.checked_add(6) else {
        return false;
    };
    let Some(hdr) = bytes.get(offset..hdr_end) else {
        return false;
    };
    let id = u16::from_le_bytes([hdr[0], hdr[1]]);
    // After PROJECTVERSION, we expect either PROJECTCONSTANTS (0x000C), a reference record, the
    // ProjectModules header, or (in some real-world streams) PROJECTCOMPATVERSION.
    if !matches!(
        id,
        0x000C | 0x003C | 0x004A | 0x000D | 0x000E | 0x0016 | 0x002F | 0x0030 | 0x0033 | 0x000F
            | 0x0013 | 0x0010 | 0x0019 | 0x0047 | 0x001A | 0x0032
    ) {
        return false;
    }
    let len = u32::from_le_bytes([hdr[2], hdr[3], hdr[4], hdr[5]]) as usize;
    offset
        .checked_add(6)
        .and_then(|v| v.checked_add(len))
        .is_some_and(|end| end <= bytes.len())
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
    fn trim_trailing_utf16_nul(bytes: &[u8]) -> &[u8] {
        if bytes.len() >= 2 && bytes.len().is_multiple_of(2) && bytes.ends_with(&[0x00, 0x00]) {
            &bytes[..bytes.len() - 2]
        } else {
            bytes
        }
    }

    if data.len() < 4 {
        return Ok(trim_trailing_utf16_nul(data));
    }

    let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
    let rest = &data[4..];

    // Treat the leading u32 as a length prefix only when it is consistent with the remaining bytes.
    let out = if n == rest.len()
        || (rest.len().is_multiple_of(2) && n.saturating_mul(2) == rest.len())
    {
        rest
    } else if rest.len() >= 2
        && rest.ends_with(&[0x00, 0x00])
        && (n.saturating_add(2) == rest.len()
            || (rest.len().is_multiple_of(2) && n.saturating_mul(2).saturating_add(2) == rest.len()))
    {
        // Some producers include a trailing UTF-16 NUL terminator but do not count it in the
        // internal length prefix.
        &rest[..rest.len() - 2]
    } else {
        // Otherwise assume the record is raw UTF-16LE bytes.
        data
    };

    // Some producers include a trailing UTF-16 NUL terminator regardless of whether it is counted
    // by the internal length prefix. Since these records represent string-like fields, strip a
    // single trailing terminator for stable hashing.
    Ok(trim_trailing_utf16_nul(out))
}
