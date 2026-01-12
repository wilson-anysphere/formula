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

        while offset < decompressed.len() {
            if offset + 6 > decompressed.len() {
                return Err(DirParseError::Truncated);
            }
            let id = u16::from_le_bytes([decompressed[offset], decompressed[offset + 1]]);
            let len = u32::from_le_bytes([
                decompressed[offset + 2],
                decompressed[offset + 3],
                decompressed[offset + 4],
                decompressed[offset + 5],
            ]) as usize;
            offset += 6;
            if offset + len > decompressed.len() {
                return Err(DirParseError::BadRecordLength { id, len });
            }
            let data = &decompressed[offset..offset + len];
            offset += len;

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
                    current_module = Some(ModuleRecord {
                        name: decode_bytes(data, encoding),
                        stream_name: String::new(),
                        module_type: ModuleType::Unknown(0),
                        text_offset: None,
                    });
                }
                0x001A => {
                    // MODULESTREAMNAME. Some files include a reserved u16 at the end.
                    if let Some(m) = current_module.as_mut() {
                        m.stream_name = decode_bytes(trim_reserved_u16(data), encoding);
                    }
                }
                0x0031 => {
                    // MODULETEXTOFFSET (u32 LE)
                    if let Some(m) = current_module.as_mut() {
                        if data.len() >= 4 {
                            let n = u32::from_le_bytes([data[0], data[1], data[2], data[3]])
                                as usize;
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
        while offset + 6 <= decompressed.len() {
            let id = u16::from_le_bytes([decompressed[offset], decompressed[offset + 1]]);
            let len = u32::from_le_bytes([
                decompressed[offset + 2],
                decompressed[offset + 3],
                decompressed[offset + 4],
                decompressed[offset + 5],
            ]) as usize;
            offset += 6;
            if offset + len > decompressed.len() {
                break;
            }
            let data = &decompressed[offset..offset + len];
            offset += len;
            if id == 0x0003 && data.len() >= 2 {
                return Some(u16::from_le_bytes([data[0], data[1]]));
            }
        }
        None
    }
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

fn looks_like_utf16le(bytes: &[u8]) -> bool {
    if bytes.len() < 2 || bytes.len() % 2 != 0 {
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
