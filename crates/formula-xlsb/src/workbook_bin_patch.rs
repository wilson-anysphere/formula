use std::io::{self, Cursor};

use crate::biff12_varint;
use crate::parser::Error;

/// BIFF record ids we need to recognize/emit for NameX add-in function interning.
///
/// These values mirror the heuristics in `parser.rs` so the patcher stays aligned
/// with the reader.
mod biff {
    // BIFF8 `SupBook` and `ExternName` record ids (stored as BIFF12 varint records in XLSB).
    pub const SUPBOOK: u32 = 0x00AE;
    pub const EXTERN_NAME: u32 = 0x0023;
    pub const END_SUPBOOK: u32 = 0x00AF;

    // BIFF12 workbook start/end markers (MS-XLSB `BrtBeginBook` / `BrtEndBook`).
    pub const END_BOOK: u32 = 0x0084;
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct InsertedNamexFunction {
    pub name: String,
    pub supbook_index: u16,
    pub name_index: u16,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct NamexFunctionPatch {
    pub workbook_bin: Vec<u8>,
    pub inserted: Vec<InsertedNamexFunction>,
    pub created_supbook: bool,
}

/// Patch `xl/workbook.bin` to ensure the workbook contains `ExternName` records for the provided
/// add-in / forward-compatible functions.
///
/// BIFF12 (XLSB) encodes "future" functions (often `_xlfn.*`) as user-defined calls:
/// - `PtgNameX` referencing an `ExternName` entry flagged as a function, followed by
/// - `PtgFuncVar(..., iftab=0x00FF)` (the UDF sentinel).
///
/// This patcher inserts missing `ExternName` records (layout A) into an AddIn `SupBook` so the
/// formula encoder can reference them:
///
/// `ExternName` layout A (minimal, parseable by `parser::parse_extern_name`):
/// - `[flags: u16]` (bit `0x0002` => `is_function`)
/// - `[scope: u16]` (`0xFFFF` => workbook scope)
/// - `[name: xlWideString]` (`u32 cch + UTF-16LE code units`)
///
/// Preservation rules:
/// - Existing workbook records are copied byte-for-byte.
/// - New `ExternName` records are inserted deterministically at the end of the chosen AddIn
///   SupBook's extern-name list (before `EndSupBook` when present, otherwise before the first
///   non-ExternName record following that SupBook).
pub(crate) fn patch_workbook_bin_intern_namex_functions(
    workbook_bin: &[u8],
    function_names: &[String],
) -> Result<Option<NamexFunctionPatch>, Error> {
    if function_names.is_empty() {
        return Ok(None);
    }

    #[derive(Debug, Clone, Copy)]
    struct SupBookState {
        index: u16,
        is_addin: bool,
        extern_name_count: u16,
    }

    let mut out = Vec::with_capacity(workbook_bin.len() + function_names.len() * 64);
    let mut inserted: Vec<InsertedNamexFunction> = Vec::new();

    let mut offset = 0usize;
    let mut supbook_count: u16 = 0;
    let mut current_supbook: Option<SupBookState> = None;
    let mut target_addin_supbook: Option<u16> = None;
    let mut inserted_into_existing = false;

    // Helper to insert new ExternName records into the current supbook.
    let mut insert_extern_names = |out: &mut Vec<u8>,
                                   supbook_index: u16,
                                   start_index: u16|
     -> Result<(), Error> {
        for (delta, name) in function_names.iter().enumerate() {
            let name_index = start_index
                .checked_add(delta as u16)
                .ok_or(Error::UnexpectedEof)?;
            let payload = build_extern_name_function_payload(name);
            write_record(out, biff::EXTERN_NAME, &payload)?;
            inserted.push(InsertedNamexFunction {
                name: name.clone(),
                supbook_index,
                name_index,
            });
        }
        Ok(())
    };

    while offset < workbook_bin.len() {
        let record_start = offset;
        let id = read_record_id(workbook_bin, &mut offset)?;
        let len = read_record_len(workbook_bin, &mut offset)? as usize;
        let payload_start = offset;
        let payload_end = payload_start
            .checked_add(len)
            .filter(|&end| end <= workbook_bin.len())
            .ok_or(Error::UnexpectedEof)?;
        let payload = &workbook_bin[payload_start..payload_end];
        offset = payload_end;

        let record_end = payload_end;
        let record_bytes = &workbook_bin[record_start..record_end];

        // Track SupBook state so we can find/patch the AddIn supbook.
        if is_supbook_record(id) {
            // If we were in a supbook without an `EndSupBook`, treat this as an implicit boundary.
            if let Some(sb) = current_supbook {
                if Some(sb.index) == target_addin_supbook && !inserted_into_existing {
                    insert_extern_names(
                        &mut out,
                        sb.index,
                        sb.extern_name_count.saturating_add(1),
                    )?;
                    inserted_into_existing = true;
                }
            }

            let raw_name = parse_supbook_raw_name(payload);
            let is_addin = raw_name.as_deref() == Some("\u{0001}");
            let index = supbook_count;
            supbook_count = supbook_count.saturating_add(1);
            current_supbook = Some(SupBookState {
                index,
                is_addin,
                extern_name_count: 0,
            });
            if is_addin && target_addin_supbook.is_none() {
                target_addin_supbook = Some(index);
            }

            out.extend_from_slice(record_bytes);
            continue;
        }

        // If we hit EndBook and never found an AddIn supbook, synthesize one just before EndBook.
        if id == biff::END_BOOK && target_addin_supbook.is_none() && !inserted_into_existing {
            let addin_index = supbook_count;
            insert_addin_supbook_section(&mut out, addin_index, function_names, &mut inserted)?;
            // Copy the EndBook record after our inserted section.
            out.extend_from_slice(record_bytes);
            return Ok(Some(NamexFunctionPatch {
                workbook_bin: out,
                inserted,
                created_supbook: true,
            }));
        }

        // When patching an existing AddIn supbook, insert before EndSupBook (or before the first
        // non-ExternName record following the SupBook when `EndSupBook` is missing).
        if let Some(mut sb) = current_supbook {
            if Some(sb.index) == target_addin_supbook && sb.is_addin && !inserted_into_existing {
                if is_end_supbook_record(id) {
                    insert_extern_names(
                        &mut out,
                        sb.index,
                        sb.extern_name_count.saturating_add(1),
                    )?;
                    inserted_into_existing = true;
                } else if !is_extern_name_record(id) {
                    insert_extern_names(
                        &mut out,
                        sb.index,
                        sb.extern_name_count.saturating_add(1),
                    )?;
                    inserted_into_existing = true;
                }
            }

            if is_extern_name_record(id) {
                sb.extern_name_count = sb.extern_name_count.saturating_add(1);
            }

            if is_end_supbook_record(id) {
                current_supbook = None;
            } else {
                current_supbook = Some(sb);
            }
        }

        out.extend_from_slice(record_bytes);
    }

    // EOF fallback: if we never found an AddIn supbook, append a new one at the end.
    if target_addin_supbook.is_none() && !inserted_into_existing {
        let addin_index = supbook_count;
        insert_addin_supbook_section(&mut out, addin_index, function_names, &mut inserted)?;
        return Ok(Some(NamexFunctionPatch {
            workbook_bin: out,
            inserted,
            created_supbook: true,
        }));
    }

    // EOF fallback: if we were inside the target AddIn supbook and never saw a boundary, append
    // the new extern names at the end.
    if !inserted_into_existing {
        if let Some(sb) = current_supbook {
            if Some(sb.index) == target_addin_supbook {
                insert_extern_names(&mut out, sb.index, sb.extern_name_count.saturating_add(1))?;
            }
        }
    }

    if inserted.is_empty() {
        return Ok(None);
    }

    Ok(Some(NamexFunctionPatch {
        workbook_bin: out,
        inserted,
        created_supbook: false,
    }))
}

fn is_supbook_record(id: u32) -> bool {
    matches!(id, 0x00AE | 0x0162 | 0x0161)
}

fn is_end_supbook_record(id: u32) -> bool {
    matches!(id, 0x0163 | 0x00AF)
}

fn is_extern_name_record(id: u32) -> bool {
    matches!(id, 0x0023 | 0x0168)
}

fn insert_addin_supbook_section(
    out: &mut Vec<u8>,
    supbook_index: u16,
    function_names: &[String],
    inserted: &mut Vec<InsertedNamexFunction>,
) -> Result<(), Error> {
    let _ = supbook_index;
    let supbook_payload = build_addin_supbook_payload();
    write_record(out, biff::SUPBOOK, &supbook_payload)?;

    for (idx, name) in function_names.iter().enumerate() {
        let name_index = u16::try_from(idx + 1).map_err(|_| Error::UnexpectedEof)?;
        let payload = build_extern_name_function_payload(name);
        write_record(out, biff::EXTERN_NAME, &payload)?;
        inserted.push(InsertedNamexFunction {
            name: name.clone(),
            supbook_index,
            name_index,
        });
    }

    // `EndSupBook` is optional in some real-world files, but including it keeps the stream
    // well-structured and prevents the workbook parser from misclassifying trailing records.
    write_record(out, biff::END_SUPBOOK, &[])?;
    Ok(())
}

fn build_addin_supbook_payload() -> Vec<u8> {
    // Minimal AddIn `SupBook` payload:
    //   [ctab:u16=0][raw_name: xlWideString("\u{0001}")]
    let mut out = Vec::new();
    out.extend_from_slice(&0u16.to_le_bytes());
    write_xl_wide_string(&mut out, "\u{0001}");
    out
}

fn build_extern_name_function_payload(name: &str) -> Vec<u8> {
    // Minimal `ExternName` layout A:
    //   [flags:u16][scope:u16=0xFFFF][name: xlWideString]
    //
    // `flags` bit 0x0002 indicates "is function".
    let mut out = Vec::new();
    out.extend_from_slice(&0x0002u16.to_le_bytes());
    out.extend_from_slice(&0xFFFFu16.to_le_bytes());
    write_xl_wide_string(&mut out, name);
    out
}

fn parse_supbook_raw_name(payload: &[u8]) -> Option<String> {
    // Try BIFF8-like layout first: [ctab:u16][raw_name: xlWideString]
    if payload.len() >= 2 + 4 {
        let mut off = 2usize;
        if let Some(name) = read_xl_wide_string(payload, &mut off) {
            return Some(name);
        }
    }

    // BIFF12-like layout: [ctab:u32][raw_name: xlWideString]
    if payload.len() >= 4 + 4 {
        let mut off = 4usize;
        if let Some(name) = read_xl_wide_string(payload, &mut off) {
            return Some(name);
        }
    }

    None
}

fn read_xl_wide_string(data: &[u8], offset: &mut usize) -> Option<String> {
    if *offset + 4 > data.len() {
        return None;
    }
    let cch = u32::from_le_bytes(data[*offset..*offset + 4].try_into().ok()?) as usize;
    *offset += 4;
    let byte_len = cch.checked_mul(2)?;
    if *offset + byte_len > data.len() {
        return None;
    }
    let bytes = &data[*offset..*offset + byte_len];
    *offset += byte_len;

    // Strict decode: return `None` on invalid surrogate sequences.
    let mut out = String::with_capacity(bytes.len());
    let iter = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]));
    for decoded in std::char::decode_utf16(iter) {
        out.push(decoded.ok()?);
    }
    Some(out)
}

fn write_xl_wide_string(out: &mut Vec<u8>, s: &str) {
    let units: Vec<u16> = s.encode_utf16().collect();
    out.extend_from_slice(&(units.len() as u32).to_le_bytes());
    for u in units {
        out.extend_from_slice(&u.to_le_bytes());
    }
}

fn write_record(out: &mut Vec<u8>, id: u32, payload: &[u8]) -> Result<(), Error> {
    biff12_varint::write_record_id(out, id).map_err(map_io_error)?;
    biff12_varint::write_record_len(out, payload.len() as u32).map_err(map_io_error)?;
    out.extend_from_slice(payload);
    Ok(())
}

fn read_record_id(data: &[u8], offset: &mut usize) -> Result<u32, Error> {
    let mut cursor = Cursor::new(data.get(*offset..).ok_or(Error::UnexpectedEof)?);
    let id = biff12_varint::read_record_id(&mut cursor).map_err(map_io_error)?;
    let Some(id) = id else {
        return Err(Error::UnexpectedEof);
    };
    *offset = offset
        .checked_add(cursor.position() as usize)
        .ok_or(Error::UnexpectedEof)?;
    Ok(id)
}

fn read_record_len(data: &[u8], offset: &mut usize) -> Result<u32, Error> {
    let mut cursor = Cursor::new(data.get(*offset..).ok_or(Error::UnexpectedEof)?);
    let len = biff12_varint::read_record_len(&mut cursor).map_err(map_io_error)?;
    let Some(len) = len else {
        return Err(Error::UnexpectedEof);
    };
    *offset = offset
        .checked_add(cursor.position() as usize)
        .ok_or(Error::UnexpectedEof)?;
    Ok(len)
}

fn map_io_error(err: io::Error) -> Error {
    if err.kind() == io::ErrorKind::UnexpectedEof {
        Error::UnexpectedEof
    } else {
        Error::Io(err)
    }
}
