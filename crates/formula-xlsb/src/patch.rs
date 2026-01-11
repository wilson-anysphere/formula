use std::collections::HashMap;
use std::io::{self, Cursor};

use crate::biff12_varint;
use crate::parser::{biff12, CellValue, Error};
use crate::writer::Biff12Writer;

/// A single cell update to apply while patch-writing a worksheet `.bin` part.
///
/// Row/col are zero-based, matching the XLSB internal representation used by the parser.
#[derive(Debug, Clone)]
pub struct CellEdit {
    pub row: u32,
    pub col: u32,
    pub new_value: CellValue,
    /// If set, replaces the raw formula token stream (`rgce`) for formula cells.
    pub new_formula: Option<Vec<u8>>,
}

/// Patch a worksheet stream (`xl/worksheets/sheetN.bin`) by rewriting only the targeted
/// cell records inside `BrtSheetData`, while copying every other record byte-for-byte.
///
/// This is a minimal bridge between the current read-only XLSB implementation and a
/// full writer:
/// - no row/column insertion
/// - updates only existing cells that appear in the stream
/// - rewrites only selected supported cell record types
pub fn patch_sheet_bin(sheet_bin: &[u8], edits: &[CellEdit]) -> Result<Vec<u8>, Error> {
    if edits.is_empty() {
        return Ok(sheet_bin.to_vec());
    }

    let mut edits_by_coord: HashMap<(u32, u32), usize> = HashMap::with_capacity(edits.len());
    for (idx, edit) in edits.iter().enumerate() {
        if edits_by_coord.insert((edit.row, edit.col), idx).is_some() {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("duplicate cell edit for ({}, {})", edit.row, edit.col),
            )));
        }
    }
    let mut applied = vec![false; edits.len()];

    let mut out = Vec::with_capacity(sheet_bin.len());
    let mut writer = Biff12Writer::new(&mut out);

    let mut offset = 0usize;
    let mut in_sheet_data = false;
    let mut current_row: Option<u32> = None;

    while offset < sheet_bin.len() {
        let record_start = offset;
        let id = read_record_id(sheet_bin, &mut offset)?;
        let len = read_record_len(sheet_bin, &mut offset)? as usize;
        let payload_start = offset;
        let payload_end = payload_start.checked_add(len).ok_or(Error::UnexpectedEof)?;
        let payload = sheet_bin
            .get(payload_start..payload_end)
            .ok_or(Error::UnexpectedEof)?;
        offset = payload_end;

        let record_end = payload_end;

        match id {
            biff12::SHEETDATA => {
                in_sheet_data = true;
                current_row = None;
                writer.write_raw(&sheet_bin[record_start..record_end])?;
            }
            biff12::SHEETDATA_END => {
                in_sheet_data = false;
                current_row = None;
                writer.write_raw(&sheet_bin[record_start..record_end])?;
            }
            biff12::ROW if in_sheet_data => {
                current_row = Some(read_u32(payload, 0)?);
                writer.write_raw(&sheet_bin[record_start..record_end])?;
            }
            biff12::FLOAT | biff12::STRING | biff12::CELL_ST | biff12::FORMULA_FLOAT
                if in_sheet_data =>
            {
                let row = current_row.unwrap_or(0);
                let col = read_u32(payload, 0)?;
                let style = read_u32(payload, 4)?;

                let Some(&edit_idx) = edits_by_coord.get(&(row, col)) else {
                    writer.write_raw(&sheet_bin[record_start..record_end])?;
                    continue;
                };

                applied[edit_idx] = true;
                let edit = &edits[edit_idx];

                match id {
                    biff12::FORMULA_FLOAT => {
                        patch_fmla_num(&mut writer, payload, col, style, edit)?;
                    }
                    _ => {
                        if edit.new_formula.is_some() {
                            return Err(Error::Io(io::Error::new(
                                io::ErrorKind::InvalidInput,
                                format!(
                                    "attempted to set formula for non-formula cell at ({row}, {col})"
                                ),
                            )));
                        }
                        patch_value_cell(&mut writer, col, style, edit)?;
                    }
                }
            }
            _ => {
                writer.write_raw(&sheet_bin[record_start..record_end])?;
            }
        }
    }

    if applied.iter().any(|&ok| !ok) {
        let mut missing = Vec::new();
        for (&(row, col), &idx) in edits_by_coord.iter() {
            if !applied[idx] {
                missing.push(format!("({row}, {col})"));
            }
        }
        missing.sort();
        return Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "cell edits not applied (cells not found): {}",
                missing.join(", ")
            ),
        )));
    }

    Ok(out)
}

fn patch_value_cell<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    col: u32,
    style: u32,
    edit: &CellEdit,
) -> Result<(), Error> {
    match &edit.new_value {
        CellValue::Number(v) => {
            let mut payload = [0u8; 16];
            payload[0..4].copy_from_slice(&col.to_le_bytes());
            payload[4..8].copy_from_slice(&style.to_le_bytes());
            payload[8..16].copy_from_slice(&v.to_le_bytes());
            writer.write_record(biff12::FLOAT, &payload)?;
        }
        CellValue::Text(s) => {
            let char_len = s.encode_utf16().count();
            let payload_len = 4u32
                .checked_add(4)
                .and_then(|v| v.checked_add(4))
                .and_then(|v| v.checked_add((char_len * 2) as u32))
                .ok_or(Error::UnexpectedEof)?;

            writer.write_record_header(biff12::CELL_ST, payload_len)?;
            writer.write_u32(col)?;
            writer.write_u32(style)?;
            writer.write_utf16_string(s)?;
        }
        other => {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!(
                    "unsupported value {:?} for cell edit at ({}, {})",
                    other, edit.row, edit.col
                ),
            )));
        }
    }
    Ok(())
}

fn patch_fmla_num<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    payload: &[u8],
    col: u32,
    style: u32,
    edit: &CellEdit,
) -> Result<(), Error> {
    // BrtFmlaNum: [col: u32][style: u32][value: f64][flags: u16][cce: u32][rgce bytes...]
    let flags = read_u16(payload, 16)?;
    let cce = read_u32(payload, 18)? as usize;
    let rgce_offset = 22usize;
    let rgce_end = rgce_offset.checked_add(cce).ok_or(Error::UnexpectedEof)?;
    let rgce = payload
        .get(rgce_offset..rgce_end)
        .ok_or(Error::UnexpectedEof)?;

    let new_rgce: &[u8] = edit.new_formula.as_deref().unwrap_or(rgce);
    let cached = match &edit.new_value {
        CellValue::Number(v) => *v,
        _ => {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!(
                    "BrtFmlaNum edit requires numeric cached value at ({}, {})",
                    edit.row, edit.col
                ),
            )));
        }
    };

    let payload_len = 22u32
        .checked_add(new_rgce.len() as u32)
        .ok_or(Error::UnexpectedEof)?;

    writer.write_record_header(biff12::FORMULA_FLOAT, payload_len)?;
    writer.write_u32(col)?;
    writer.write_u32(style)?;
    writer.write_f64(cached)?;
    writer.write_u16(flags)?;
    writer.write_u32(new_rgce.len() as u32)?;
    writer.write_raw(new_rgce)?;
    Ok(())
}

fn read_u16(data: &[u8], offset: usize) -> Result<u16, Error> {
    let bytes: [u8; 2] = data
        .get(offset..offset + 2)
        .ok_or(Error::UnexpectedEof)?
        .try_into()
        .unwrap();
    Ok(u16::from_le_bytes(bytes))
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32, Error> {
    let bytes: [u8; 4] = data
        .get(offset..offset + 4)
        .ok_or(Error::UnexpectedEof)?
        .try_into()
        .unwrap();
    Ok(u32::from_le_bytes(bytes))
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
