use encoding_rs::{Encoding, UTF_16LE, WINDOWS_1252};
use thiserror::Error;

#[derive(Debug, Copy, Clone, PartialEq, Eq)]
pub enum ModuleType {
    Standard,
    Class,
    Document,
    UserForm,
    Unknown(u16),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ModuleRecord {
    pub name: String,
    pub stream_name: String,
    pub module_type: ModuleType,
    pub text_offset: Option<usize>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DirStream {
    pub project_name: Option<String>,
    pub constants: Option<String>,
    pub modules: Vec<ModuleRecord>,
}

#[derive(Debug, Error)]
pub enum DirParseError {
    #[error("dir stream is truncated")]
    Truncated,
    #[error("dir record claims a length beyond the remaining bytes (id={id:#06x}, len={len})")]
    BadRecordLength { id: u16, len: usize },
    #[error("unexpected dir record id (expected {expected:#06x}, found {found:#06x})")]
    UnexpectedRecordId { expected: u16, found: u16 },
    #[error("invalid dir stream: {0}")]
    Invalid(&'static str),
}

impl DirStream {
    /// Parse a decompressed `VBA/dir` stream.
    ///
    /// We only interpret a small subset of records required to recover:
    /// - project name
    /// - constants
    /// - module list with stream names + text offsets
    pub fn parse_with_encoding(
        decompressed: &[u8],
        encoding: &'static Encoding,
    ) -> Result<Self, DirParseError> {
        let mut offset = 0usize;
        let mut project_name = None;
        let mut constants = None;
        let mut modules: Vec<ModuleRecord> = Vec::new();
        let mut current_module: Option<ModuleRecord> = None;
        // Some `VBA/dir` layouts store the Unicode module stream name as a separate record
        // immediately following MODULESTREAMNAME (0x001A). Track that expectation so we can support
        // alternate record IDs without misinterpreting unrelated records.
        let mut expect_stream_name_unicode = false;

        while offset < decompressed.len() {
            let Some(id_end) = offset.checked_add(2) else {
                return Err(DirParseError::Truncated);
            };
            let Some(id_bytes) = decompressed.get(offset..id_end) else {
                return Err(DirParseError::Truncated);
            };
            let id = u16::from_le_bytes([id_bytes[0], id_bytes[1]]);

            // Some spec-compliant `VBA/dir` record layouts include fixed-length records that do not
            // follow the common `Id(u16) || Size(u32) || Data(Size)` pattern. The one we most care
            // about for module enumeration is PROJECTVERSION (0x0009), which is:
            //   Id(u16) || Reserved(u32) || VersionMajor(u32) || VersionMinor(u16)
            // (12 bytes total).
            //
            // If we treat the u32 as a generic `Size` field, we can mis-align parsing and fail to
            // reach the module record groups.
            if id == 0x0009 {
                expect_stream_name_unicode = false;

                // Attempt to disambiguate between:
                // - a TLV-ish variant: Id(u16) || Size(u32) || Data(Size)
                // - the MS-OVBA fixed-length layout described above.
                //
                // We pick the layout whose end offset is followed by a plausible record header.
                let record_start = offset;
                let record_end = record_start
                    .checked_add(6)
                    .ok_or(DirParseError::Truncated)?;
                let Some(header) = decompressed.get(record_start..record_end) else {
                    return Err(DirParseError::Truncated);
                };
                let size_field =
                    u32::from_le_bytes([header[2], header[3], header[4], header[5]]) as usize;

                let tlv_end = record_end.checked_add(size_field).unwrap_or(usize::MAX);
                let fixed_end = record_start.checked_add(12).unwrap_or(usize::MAX);

                let tlv_next_ok = looks_like_projectversion_following_record(decompressed, tlv_end);
                let fixed_next_ok =
                    looks_like_projectversion_following_record(decompressed, fixed_end);

                if fixed_end <= decompressed.len() && fixed_next_ok && !tlv_next_ok {
                    offset = fixed_end;
                    continue;
                }
                if fixed_end <= decompressed.len() && fixed_next_ok && size_field == 0 {
                    // Size=0 is a strong signal for the fixed-length layout: a PROJECTVERSION record
                    // with no payload bytes is unlikely.
                    offset = fixed_end;
                    continue;
                }

                // Fall back to treating it as a TLV record (skip the declared payload length).
                if tlv_end > decompressed.len() {
                    return Err(DirParseError::BadRecordLength {
                        id,
                        len: size_field,
                    });
                }
                offset = tlv_end;
                continue;
            }

            // MODULESTREAMNAME (0x001A) can also deviate from the common TLV-ish pattern.
            //
            // Per MS-OVBA §2.3.4.2.3.2.2, MODULESTREAMNAME is:
            //   Id(u16)
            //   SizeOfStreamName(u32)
            //   StreamName(MBCS bytes)
            //   Reserved(u16)=0x0032
            //   SizeOfStreamNameUnicode(u32)
            //   StreamNameUnicode(UTF-16LE bytes)
            //
            // i.e. the u32 after the Id is **SizeOfStreamName**, not the total record length.
            // If we treat it as a generic `Size` field, parsing becomes misaligned when the
            // Reserved=0x0032 + Unicode tail is present.
            if id == 0x001A {
                let record_start = offset;
                let header_end = record_start
                    .checked_add(6)
                    .ok_or(DirParseError::Truncated)?;
                let Some(header) = decompressed.get(record_start..header_end) else {
                    return Err(DirParseError::Truncated);
                };
                let size_name =
                    u32::from_le_bytes([header[2], header[3], header[4], header[5]]) as usize;

                let name_start = record_start
                    .checked_add(6)
                    .ok_or(DirParseError::Truncated)?;
                let name_end = name_start
                    .checked_add(size_name)
                    .ok_or(DirParseError::BadRecordLength { id, len: size_name })?;
                if name_end > decompressed.len() {
                    return Err(DirParseError::BadRecordLength { id, len: size_name });
                }

                // Parse the Unicode stream name when the reserved marker is present immediately
                // after the MBCS bytes.
                if name_end.checked_add(6).is_some_and(|end| end <= decompressed.len()) {
                    let reserved_end = name_end
                        .checked_add(2)
                        .ok_or(DirParseError::Truncated)?;
                    let Some(reserved_bytes) = decompressed.get(name_end..reserved_end) else {
                        return Err(DirParseError::Truncated);
                    };
                    let reserved = u16::from_le_bytes([reserved_bytes[0], reserved_bytes[1]]);
                    if reserved == 0x0032 {
                        let len_start = reserved_end;
                        let len_end = len_start
                            .checked_add(4)
                            .ok_or(DirParseError::Truncated)?;
                        let Some(unicode_len_bytes) = decompressed.get(len_start..len_end) else {
                            return Err(DirParseError::Truncated);
                        };
                        let size_unicode = u32::from_le_bytes([
                            unicode_len_bytes[0],
                            unicode_len_bytes[1],
                            unicode_len_bytes[2],
                            unicode_len_bytes[3],
                        ]) as usize;
                        let unicode_start = name_end
                            .checked_add(6)
                            .ok_or(DirParseError::Truncated)?;
                        let unicode_end = unicode_start
                            .checked_add(size_unicode)
                            .ok_or(DirParseError::BadRecordLength {
                                id,
                                len: size_unicode,
                            })?;
                        if unicode_end > decompressed.len() {
                            return Err(DirParseError::BadRecordLength {
                                id,
                                len: size_unicode,
                            });
                        }

                        if let Some(m) = current_module.as_mut() {
                            m.stream_name =
                                decode_unicode_bytes(&decompressed[unicode_start..unicode_end]);
                        }
                        expect_stream_name_unicode = false;
                        offset = unicode_end;
                        continue;
                    }
                }

                // Otherwise, treat it as a normal `Id || Size || Data` record and decode the MBCS
                // bytes (trimming a common trailing reserved u16=0x0000).
                let data = &decompressed[name_start..name_end];
                if let Some(m) = current_module.as_mut() {
                    m.stream_name = decode_bytes(trim_reserved_u16(data), encoding);
                    // Some `VBA/dir` layouts provide a Unicode stream-name record immediately after
                    // MODULESTREAMNAME (either MODULESTREAMNAMEUNICODE (0x0032) or, in some files,
                    // record id 0x0048). Preserve the pre-existing behavior by tracking that
                    // expectation here.
                    expect_stream_name_unicode = true;
                } else {
                    expect_stream_name_unicode = false;
                }
                offset = name_end;
                continue;
            }

            let Some(header_end) = offset.checked_add(6) else {
                return Err(DirParseError::Truncated);
            };
            let Some(header) = decompressed.get(offset..header_end) else {
                return Err(DirParseError::Truncated);
            };
            let len = u32::from_le_bytes([header[2], header[3], header[4], header[5]]) as usize;
            offset += 6;
            let Some(end) = offset.checked_add(len) else {
                return Err(DirParseError::BadRecordLength { id, len });
            };
            if end > decompressed.len() {
                return Err(DirParseError::BadRecordLength { id, len });
            }
            let data = decompressed.get(offset..end).ok_or(DirParseError::Truncated)?;
            offset = end;

            if expect_stream_name_unicode && !matches!(id, 0x0032 | 0x0048) {
                expect_stream_name_unicode = false;
            }

            match id {
                0x0003 => {
                    // PROJECTCODEPAGE (u16 LE). We currently use the `PROJECT` stream as the
                    // primary source of the codepage, but exposing this record allows callers
                    // to fall back when `CodePage=` is missing.
                }
                0x0004 => project_name = Some(decode_bytes(data, encoding)),
                0x000C => constants = Some(decode_bytes(data, encoding)),

                // Module records.
                0x0019 => {
                    // MODULENAME: start a new module
                    if let Some(m) = current_module.take() {
                        modules.push(m);
                    }
                    expect_stream_name_unicode = false;
                    current_module = Some(ModuleRecord {
                        name: decode_bytes(data, encoding),
                        stream_name: String::new(),
                        module_type: ModuleType::Unknown(0),
                        text_offset: None,
                    });
                }
                0x0047 => {
                    // MODULENAMEUNICODE.
                    //
                    // When both ANSI and Unicode variants are present, this record usually appears
                    // immediately after MODULENAME and should be treated as an alternate
                    // representation for the current module's name. When the ANSI MODULENAME record
                    // is absent, treat this as the start of a new module.
                    let start_new = match current_module.as_ref() {
                        None => true,
                        Some(m) => {
                            !m.stream_name.is_empty()
                                || m.text_offset.is_some()
                                || !matches!(m.module_type, ModuleType::Unknown(0))
                        }
                    };

                    if start_new {
                        if let Some(m) = current_module.take() {
                            modules.push(m);
                        }
                        current_module = Some(ModuleRecord {
                            name: decode_unicode_bytes(data),
                            stream_name: String::new(),
                            module_type: ModuleType::Unknown(0),
                            text_offset: None,
                        });
                    } else if let Some(m) = current_module.as_mut() {
                        m.name = decode_unicode_bytes(data);
                    }
                }
                0x001A => {
                    // MODULESTREAMNAME. Some files include a reserved u16 at the end.
                    if let Some(m) = current_module.as_mut() {
                        m.stream_name = decode_bytes(trim_reserved_u16(data), encoding);
                        expect_stream_name_unicode = true;
                    }
                }
                0x0032 => {
                    // MODULESTREAMNAMEUNICODE.
                    //
                    // Some real-world `VBA/dir` streams provide a UTF-16LE module stream name in a
                    // separate record (ID 0x0032). Prefer it for OLE stream lookup.
                    if let Some(m) = current_module.as_mut() {
                        m.stream_name = decode_unicode_bytes(data);
                    }
                    expect_stream_name_unicode = false;
                }
                0x0048 if expect_stream_name_unicode => {
                    // Some producers use 0x0048 as the module stream name Unicode record id.
                    if let Some(m) = current_module.as_mut() {
                        m.stream_name = decode_unicode_bytes(data);
                    }
                    expect_stream_name_unicode = false;
                }
                0x0031 => {
                    // MODULETEXTOFFSET (u32 LE)
                    if let Some(m) = current_module.as_mut() {
                        if data.len() >= 4 {
                            let n =
                                u32::from_le_bytes([data[0], data[1], data[2], data[3]]) as usize;
                            m.text_offset = Some(n);
                        }
                    }
                }
                0x0021 => {
                    // MODULETYPE (u16)
                    if let Some(m) = current_module.as_mut() {
                        if data.len() >= 2 {
                            let t = u16::from_le_bytes([data[0], data[1]]);
                            m.module_type = match t {
                                0x0000 => ModuleType::Standard,
                                0x0001 => ModuleType::Class,
                                0x0002 => ModuleType::Document,
                                0x0003 => ModuleType::UserForm,
                                other => ModuleType::Unknown(other),
                            };
                        }
                    }
                }
                _ => {}
            }
        }

        if let Some(m) = current_module.take() {
            modules.push(m);
        }

        // If stream_name wasn't recorded, default to module name (common case).
        for m in &mut modules {
            if m.stream_name.is_empty() {
                m.stream_name = m.name.clone();
            }
        }

        Ok(Self {
            project_name,
            constants,
            modules,
        })
    }

    /// Parse a decompressed `VBA/dir` stream assuming the default VBA codepage (Windows-1252).
    pub fn parse(decompressed: &[u8]) -> Result<Self, DirParseError> {
        Self::parse_with_encoding(decompressed, WINDOWS_1252)
    }

    /// Best-effort extraction of `PROJECTCODEPAGE` from a decompressed `VBA/dir` stream.
    pub fn detect_codepage(decompressed: &[u8]) -> Option<u16> {
        let mut offset = 0usize;
        while let Some(hdr_end) = offset.checked_add(6) {
            let Some(hdr) = decompressed.get(offset..hdr_end) else {
                break;
            };
            let id = u16::from_le_bytes([hdr[0], hdr[1]]);
            let len = u32::from_le_bytes([hdr[2], hdr[3], hdr[4], hdr[5]]) as usize;
            offset = hdr_end;

            let Some(end) = offset.checked_add(len) else {
                break;
            };
            let Some(data) = decompressed.get(offset..end) else {
                break;
            };
            offset = end;
            if id == 0x0003 && data.len() >= 2 {
                return Some(u16::from_le_bytes([data[0], data[1]]));
            }
        }
        None
    }
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
        0x000C
            | 0x003C
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

fn decode_bytes(bytes: &[u8], encoding: &'static Encoding) -> String {
    // MS-OVBA strings are generally stored using the project codepage, but some
    // records may appear in UTF-16LE form. We do a lightweight heuristic to
    // decode UTF-16LE when it looks plausible (common for simple ASCII names).
    if looks_like_utf16le(bytes) {
        let (cow, _) = UTF_16LE.decode_without_bom_handling(bytes);
        return cow.into_owned();
    }

    let (cow, _, _) = encoding.decode(bytes);
    cow.into_owned()
}

fn decode_unicode_bytes(bytes: &[u8]) -> String {
    let mut utf16_bytes = bytes;

    // Some producers include an internal u32 length prefix. Accept both:
    // - length == remaining bytes (byte count), or
    // - length*2 == remaining bytes (UTF-16 code units).
    if bytes.len() >= 4 {
        let n = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]) as usize;
        let rest = &bytes[4..];
        if n == rest.len() || n.saturating_mul(2) == rest.len() {
            utf16_bytes = rest;
        } else if rest.len() >= 2 && rest.ends_with(&[0x00, 0x00]) {
            // Some producers include a trailing UTF-16 NUL terminator but do not count it in the
            // internal length prefix.
            if n.saturating_add(2) == rest.len()
                || n.saturating_mul(2).saturating_add(2) == rest.len()
            {
                utf16_bytes = rest;
            }
        }
    }

    let (cow, _) = UTF_16LE.decode_without_bom_handling(utf16_bytes);
    let mut s = cow.into_owned();
    // Stream/module names should not contain NULs; strip defensively.
    s.retain(|c| c != '\u{0000}');
    s
}

fn looks_like_utf16le(bytes: &[u8]) -> bool {
    if bytes.len() < 2 || !bytes.len().is_multiple_of(2) {
        return false;
    }
    // If a substantial portion of the high bytes are NUL, it's probably
    // UTF-16LE for ASCII-range characters.
    let high_bytes = bytes.iter().skip(1).step_by(2);
    let total = bytes.len() / 2;
    let nul_count = high_bytes.filter(|&&b| b == 0).count();
    // Use a ceiling half threshold. For very short inputs (e.g. 2 bytes), `total / 2` is 0 and
    // would incorrectly classify any 2-byte MBCS string as UTF-16LE.
    nul_count * 2 >= total
}
