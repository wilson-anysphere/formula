use std::collections::HashMap;
use std::io::{self, Read, Write};

use crate::parser::{biff12, Error};

use super::{Biff12Writer, CellEdit};

#[derive(Debug)]
struct RawRecordHeader {
    id: u32,
    id_raw: Vec<u8>,
    len: u32,
    len_raw: Vec<u8>,
}

/// Patch a worksheet record stream (`xl/worksheets/sheetN.bin`) without materializing the whole
/// part in memory.
///
/// Returns `Ok(true)` when at least one record was rewritten, and `Ok(false)` when the output is
/// byte-identical to the input stream.
pub fn patch_sheet_bin_streaming<R: Read, W: Write>(
    mut input: R,
    output: W,
    edits: &[CellEdit],
) -> Result<bool, Error> {
    let mut writer = Biff12Writer::new(output);

    if edits.is_empty() {
        copy_remaining(&mut input, &mut writer)?;
        return Ok(false);
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

    let mut in_sheet_data = false;
    let mut current_row: Option<u32> = None;
    let mut changed = false;

    while let Some(header) = read_record_header(&mut input)? {
        let id = header.id;
        let len = header.len as usize;

        match id {
            biff12::SHEETDATA => {
                in_sheet_data = true;
                current_row = None;
                write_raw_header(&mut writer, &header)?;
                copy_exact(&mut input, &mut writer, len)?;
            }
            biff12::SHEETDATA_END => {
                in_sheet_data = false;
                current_row = None;
                write_raw_header(&mut writer, &header)?;
                copy_exact(&mut input, &mut writer, len)?;
            }
            biff12::ROW if in_sheet_data => {
                let payload = read_payload(&mut input, len)?;
                current_row = Some(super::read_u32(&payload, 0)?);
                write_raw_header(&mut writer, &header)?;
                writer.write_raw(&payload)?;
            }
            biff12::BLANK
            | biff12::BOOLERR
            | biff12::BOOL
            | biff12::NUM
            | biff12::FLOAT
            | biff12::STRING
            | biff12::CELL_ST
            | biff12::FORMULA_FLOAT
            | biff12::FORMULA_STRING
            | biff12::FORMULA_BOOL
            | biff12::FORMULA_BOOLERR
                if in_sheet_data =>
            {
                let payload = read_payload(&mut input, len)?;
                let row = current_row.unwrap_or(0);
                let col = super::read_u32(&payload, 0)?;
                let style = super::read_u32(&payload, 4)?;

                let Some(&edit_idx) = edits_by_coord.get(&(row, col)) else {
                    write_raw_header(&mut writer, &header)?;
                    writer.write_raw(&payload)?;
                    continue;
                };

                applied[edit_idx] = true;
                let edit = &edits[edit_idx];

                match id {
                    biff12::FORMULA_FLOAT => {
                        if super::formula_edit_is_noop(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            changed = true;
                            super::patch_fmla_num(&mut writer, &payload, col, style, edit)?;
                        }
                    }
                    biff12::FORMULA_STRING => {
                        if super::formula_string_edit_is_noop(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            changed = true;
                            super::patch_fmla_string(&mut writer, &payload, col, style, edit)?;
                        }
                    }
                    biff12::FORMULA_BOOL => {
                        if super::formula_bool_edit_is_noop(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            changed = true;
                            super::patch_fmla_bool(&mut writer, &payload, col, style, edit)?;
                        }
                    }
                    biff12::FORMULA_BOOLERR => {
                        if super::formula_error_edit_is_noop(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            changed = true;
                            super::patch_fmla_error(&mut writer, &payload, col, style, edit)?;
                        }
                    }
                    biff12::FLOAT => {
                        if super::value_edit_is_noop_float(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            if edit.new_formula.is_some() {
                                return Err(Error::Io(io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    format!(
                                        "attempted to set formula for non-formula cell at ({row}, {col})"
                                    ),
                                )));
                            }
                            changed = true;
                            super::patch_value_cell(&mut writer, col, style, edit)?;
                        }
                    }
                    biff12::NUM => {
                        if super::value_edit_is_noop_rk(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            if edit.new_formula.is_some() {
                                return Err(Error::Io(io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    format!(
                                        "attempted to set formula for non-formula cell at ({row}, {col})"
                                    ),
                                )));
                            }
                            changed = true;
                            super::patch_rk_cell(&mut writer, col, style, &payload, edit)?;
                        }
                    }
                    biff12::CELL_ST => {
                        if super::value_edit_is_noop_inline_string(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            if edit.new_formula.is_some() {
                                return Err(Error::Io(io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    format!(
                                        "attempted to set formula for non-formula cell at ({row}, {col})"
                                    ),
                                )));
                            }
                            changed = true;
                            super::patch_value_cell(&mut writer, col, style, edit)?;
                        }
                    }
                    biff12::BOOL => {
                        if super::value_edit_is_noop_bool(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            if edit.new_formula.is_some() {
                                return Err(Error::Io(io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    format!(
                                        "attempted to set formula for non-formula cell at ({row}, {col})"
                                    ),
                                )));
                            }
                            changed = true;
                            super::patch_value_cell(&mut writer, col, style, edit)?;
                        }
                    }
                    biff12::BOOLERR => {
                        if super::value_edit_is_noop_error(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            if edit.new_formula.is_some() {
                                return Err(Error::Io(io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    format!(
                                        "attempted to set formula for non-formula cell at ({row}, {col})"
                                    ),
                                )));
                            }
                            changed = true;
                            super::patch_value_cell(&mut writer, col, style, edit)?;
                        }
                    }
                    biff12::BLANK => {
                        if super::value_edit_is_noop_blank(edit) {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            if edit.new_formula.is_some() {
                                return Err(Error::Io(io::Error::new(
                                    io::ErrorKind::InvalidInput,
                                    format!(
                                        "attempted to set formula for non-formula cell at ({row}, {col})"
                                    ),
                                )));
                            }
                            changed = true;
                            super::patch_value_cell(&mut writer, col, style, edit)?;
                        }
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
                        changed = true;
                        super::patch_value_cell(&mut writer, col, style, edit)?;
                    }
                }
            }
            _ => {
                write_raw_header(&mut writer, &header)?;
                copy_exact(&mut input, &mut writer, len)?;
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

    Ok(changed)
}

fn read_record_header<R: Read>(r: &mut R) -> Result<Option<RawRecordHeader>, Error> {
    let Some((id, id_raw)) = read_record_id_raw(r)? else {
        return Ok(None);
    };
    let (len, len_raw) = read_record_len_raw(r)?;
    Ok(Some(RawRecordHeader {
        id,
        id_raw,
        len,
        len_raw,
    }))
}

fn read_record_id_raw<R: Read>(r: &mut R) -> Result<Option<(u32, Vec<u8>)>, Error> {
    let mut v: u32 = 0;
    let mut raw = Vec::with_capacity(4);

    for i in 0..4 {
        let mut buf = [0u8; 1];
        let n = r.read(&mut buf).map_err(super::map_io_error)?;
        if n == 0 {
            if i == 0 {
                return Ok(None);
            }
            return Err(Error::UnexpectedEof);
        }

        let byte = buf[0];
        raw.push(byte);
        v |= (byte as u32) << (8 * i);
        if byte & 0x80 == 0 {
            return Ok(Some((v, raw)));
        }
    }

    Err(Error::Io(io::Error::new(
        io::ErrorKind::InvalidData,
        "invalid BIFF12 record id (more than 4 bytes)",
    )))
}

fn read_record_len_raw<R: Read>(r: &mut R) -> Result<(u32, Vec<u8>), Error> {
    let mut v: u32 = 0;
    let mut raw = Vec::with_capacity(4);

    for i in 0..4 {
        let mut buf = [0u8; 1];
        let n = r.read(&mut buf).map_err(super::map_io_error)?;
        if n == 0 {
            return Err(Error::UnexpectedEof);
        }

        let byte = buf[0];
        raw.push(byte);
        v |= ((byte & 0x7F) as u32) << (7 * i);
        if byte & 0x80 == 0 {
            return Ok((v, raw));
        }
    }

    Err(Error::Io(io::Error::new(
        io::ErrorKind::InvalidData,
        "invalid BIFF12 record length (more than 4 bytes)",
    )))
}

fn write_raw_header<W: Write>(writer: &mut Biff12Writer<W>, header: &RawRecordHeader) -> io::Result<()> {
    writer.write_raw(&header.id_raw)?;
    writer.write_raw(&header.len_raw)?;
    Ok(())
}

fn read_payload<R: Read>(r: &mut R, len: usize) -> Result<Vec<u8>, Error> {
    let mut payload = vec![0u8; len];
    r.read_exact(&mut payload).map_err(super::map_io_error)?;
    Ok(payload)
}

fn copy_exact<R: Read, W: Write>(
    input: &mut R,
    writer: &mut Biff12Writer<W>,
    mut len: usize,
) -> Result<(), Error> {
    let mut buf = [0u8; 16 * 1024];
    while len > 0 {
        let chunk_len = buf.len().min(len);
        input
            .read_exact(&mut buf[..chunk_len])
            .map_err(super::map_io_error)?;
        writer.write_raw(&buf[..chunk_len])?;
        len = len.saturating_sub(chunk_len);
    }
    Ok(())
}

fn copy_remaining<R: Read, W: Write>(input: &mut R, writer: &mut Biff12Writer<W>) -> Result<(), Error> {
    let mut buf = [0u8; 16 * 1024];
    loop {
        let n = input.read(&mut buf).map_err(super::map_io_error)?;
        if n == 0 {
            break;
        }
        writer.write_raw(&buf[..n])?;
    }
    Ok(())
}
