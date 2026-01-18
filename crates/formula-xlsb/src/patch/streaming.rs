use std::collections::HashMap;
use std::io::{self, Read, Write};

use crate::parser::{biff12, CellValue, Error};
use crate::opc::max_xlsb_zip_part_bytes;

use super::{Biff12Writer, CellEdit};

const READ_PAYLOAD_CHUNK_BYTES: usize = 64 * 1024; // 64 KiB

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
/// This streaming patcher supports the same insertion semantics as [`super::patch_sheet_bin`]:
/// - missing rows/cells inside `BrtSheetData` are inserted in row-major order
/// - edits that clear a missing cell to blank (`new_value = Blank` + no formula payload) are
///   treated as no-ops and do not materialize new records
/// - the worksheet `BrtWsDim` / `DIMENSION` record is expanded (never shrunk) to include edits
///   that materialize cell records, including formatting-only updates to blank cells
///
/// Returns `Ok(true)` when the output differs from the input stream, and `Ok(false)` when the
/// output is byte-identical.
pub fn patch_sheet_bin_streaming<R: Read, W: Write>(
    mut input: R,
    output: W,
    edits: &[CellEdit],
) -> Result<bool, Error> {
    let max_sheet_bytes = max_xlsb_zip_part_bytes();

    if edits.is_empty() {
        let mut writer = Biff12Writer::new(output);
        copy_remaining(&mut input, &mut writer)?;
        return Ok(false);
    }

    super::validate_cell_edits(edits)?;

    let mut edits_by_coord: HashMap<(u32, u32), usize> = HashMap::new();
    let _ = edits_by_coord.try_reserve(edits.len());
    for (idx, edit) in edits.iter().enumerate() {
        if edit.clear_formula
            && (edit.new_formula.is_some()
                || edit.new_rgcb.is_some()
                || edit.new_formula_flags.is_some())
        {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!(
                    "cell edit for ({}, {}) cannot set new_formula/new_rgcb/new_formula_flags when clear_formula=true",
                    edit.row, edit.col
                ),
            )));
        }
        if edits_by_coord.insert((edit.row, edit.col), idx).is_some() {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("duplicate cell edit for ({}, {})", edit.row, edit.col),
            )));
        }
    }
    let mut ordered_edits: Vec<usize> = (0..edits.len()).collect();
    ordered_edits.sort_by_key(|&idx| (edits[idx].row, edits[idx].col));
    let mut applied = vec![false; edits.len()];
    let mut insert_cursor = 0usize;

    // For streaming we cannot "backpatch" the output stream, so precompute the bounding box of
    // all edits that can materialize new cell records (including formatted blanks). This is a
    // conservative expansion that keeps the used range consistent when we insert new rows/cells.
    let mut requested_bounds: Option<super::Bounds> = None;
    for edit in edits {
        if super::insertion_is_noop(edit) {
            continue;
        }
        super::bounds_include(&mut requested_bounds, edit.row, edit.col);
    }

    // When the worksheet stream lacks BrtWsDim (DIMENSION), we need to synthesize it so used-range
    // metadata stays consistent after inserting non-blank cells. The streaming patcher cannot
    // backpatch already-written bytes, so when we *might* need to insert DIMENSION we first scan
    // the record prefix:
    // - If we encounter DIMENSION before SHEETDATA, we can proceed streaming normally.
    // - If we hit SHEETDATA first, DIMENSION may appear later (non-standard) or be missing
    //   entirely. In that case we fall back to the in-memory patcher so we can deterministically
    //   insert DIMENSION before SHEETDATA when needed.
    let mut prefix_bytes: Vec<u8> = Vec::new();
    let mut changed = false;
    if requested_bounds.is_some() {
        let mut saw_dimension = false;
        while let Some(header) = read_record_header(&mut input)? {
            let id = header.id;
            let len = header.len as usize;

            match id {
                biff12::DIMENSION => {
                    let mut payload = read_payload(&mut input, len)?;
                    if len < 16 {
                        return Err(Error::UnexpectedEof);
                    }

                    let r1 = super::read_u32(&payload, 0)?;
                    let r2 = super::read_u32(&payload, 4)?;
                    let c1 = super::read_u32(&payload, 8)?;
                    let c2 = super::read_u32(&payload, 12)?;

                    let mut new_r1 = r1;
                    let mut new_r2 = r2;
                    let mut new_c1 = c1;
                    let mut new_c2 = c2;

                    if let Some(bounds) = requested_bounds {
                        new_r1 = new_r1.min(bounds.min_row);
                        new_r2 = new_r2.max(bounds.max_row);
                        new_c1 = new_c1.min(bounds.min_col);
                        new_c2 = new_c2.max(bounds.max_col);
                    }

                    if (new_r1, new_r2, new_c1, new_c2) != (r1, r2, c1, c2) {
                        payload[0..4].copy_from_slice(&new_r1.to_le_bytes());
                        payload[4..8].copy_from_slice(&new_r2.to_le_bytes());
                        payload[8..12].copy_from_slice(&new_c1.to_le_bytes());
                        payload[12..16].copy_from_slice(&new_c2.to_le_bytes());
                        changed = true;
                    }

                    prefix_bytes.extend_from_slice(&header.id_raw);
                    prefix_bytes.extend_from_slice(&header.len_raw);
                    prefix_bytes.extend_from_slice(&payload);
                    saw_dimension = true;
                    break;
                }
                biff12::SHEETDATA => {
                    // No DIMENSION record appeared before SHEETDATA. Buffer the entire stream and
                    // delegate to the in-memory patcher so we can insert DIMENSION before SHEETDATA
                    // iff it is missing entirely.
                    prefix_bytes.extend_from_slice(&header.id_raw);
                    prefix_bytes.extend_from_slice(&header.len_raw);
                    let payload = read_payload(&mut input, len)?;
                    prefix_bytes.extend_from_slice(&payload);

                    let prefix_len = prefix_bytes.len() as u64;
                    if prefix_len > max_sheet_bytes {
                        return Err(Error::PartTooLarge {
                            part: "worksheet stream".to_string(),
                            size: prefix_len,
                            max: max_sheet_bytes,
                        });
                    }
                    let remaining_limit =
                        max_sheet_bytes.saturating_sub(prefix_len).saturating_add(1);
                    let mut sheet_bin = prefix_bytes;
                    (&mut input)
                        .take(remaining_limit)
                        .read_to_end(&mut sheet_bin)
                        .map_err(super::map_io_error)?;
                    if sheet_bin.len() as u64 > max_sheet_bytes {
                        return Err(Error::PartTooLarge {
                            part: "worksheet stream".to_string(),
                            size: sheet_bin.len() as u64,
                            max: max_sheet_bytes,
                        });
                    }
                    let patched = super::patch_sheet_bin(&sheet_bin, edits)?;

                    let mut writer = Biff12Writer::new(output);
                    writer.write_raw(&patched)?;
                    return Ok(patched != sheet_bin);
                }
                _ => {
                    prefix_bytes.extend_from_slice(&header.id_raw);
                    prefix_bytes.extend_from_slice(&header.len_raw);
                    let payload = read_payload(&mut input, len)?;
                    prefix_bytes.extend_from_slice(&payload);
                }
            }
        }

        if !saw_dimension {
            // End-of-stream (or malformed worksheet) without any DIMENSION record.
            let prefix_len = prefix_bytes.len() as u64;
            if prefix_len > max_sheet_bytes {
                return Err(Error::PartTooLarge {
                    part: "worksheet stream".to_string(),
                    size: prefix_len,
                    max: max_sheet_bytes,
                });
            }
            let remaining_limit = max_sheet_bytes.saturating_sub(prefix_len).saturating_add(1);
            let mut sheet_bin = prefix_bytes;
            (&mut input)
                .take(remaining_limit)
                .read_to_end(&mut sheet_bin)
                .map_err(super::map_io_error)?;
            if sheet_bin.len() as u64 > max_sheet_bytes {
                return Err(Error::PartTooLarge {
                    part: "worksheet stream".to_string(),
                    size: sheet_bin.len() as u64,
                    max: max_sheet_bytes,
                });
            }
            let patched = super::patch_sheet_bin(&sheet_bin, edits)?;

            let mut writer = Biff12Writer::new(output);
            writer.write_raw(&patched)?;
            return Ok(patched != sheet_bin);
        }
    }

    let mut writer = Biff12Writer::new(output);
    if !prefix_bytes.is_empty() {
        writer.write_raw(&prefix_bytes)?;
    }

    let mut in_sheet_data = false;
    let mut current_row: Option<u32> = None;
    let mut row_template: Option<Vec<u8>> = None;
    let mut dim_additions: Option<super::Bounds> = None;

    while let Some(header) = read_record_header(&mut input)? {
        let id = header.id;
        let len = header.len as usize;

        match id {
            biff12::DIMENSION => {
                let mut payload = read_payload(&mut input, len)?;
                if len < 16 {
                    return Err(Error::UnexpectedEof);
                }

                let r1 = super::read_u32(&payload, 0)?;
                let r2 = super::read_u32(&payload, 4)?;
                let c1 = super::read_u32(&payload, 8)?;
                let c2 = super::read_u32(&payload, 12)?;

                let mut new_r1 = r1;
                let mut new_r2 = r2;
                let mut new_c1 = c1;
                let mut new_c2 = c2;

                if let Some(bounds) = requested_bounds {
                    new_r1 = new_r1.min(bounds.min_row);
                    new_r2 = new_r2.max(bounds.max_row);
                    new_c1 = new_c1.min(bounds.min_col);
                    new_c2 = new_c2.max(bounds.max_col);
                }
                if let Some(bounds) = dim_additions {
                    new_r1 = new_r1.min(bounds.min_row);
                    new_r2 = new_r2.max(bounds.max_row);
                    new_c1 = new_c1.min(bounds.min_col);
                    new_c2 = new_c2.max(bounds.max_col);
                }

                if (new_r1, new_r2, new_c1, new_c2) != (r1, r2, c1, c2) {
                    payload[0..4].copy_from_slice(&new_r1.to_le_bytes());
                    payload[4..8].copy_from_slice(&new_r2.to_le_bytes());
                    payload[8..12].copy_from_slice(&new_c1.to_le_bytes());
                    payload[12..16].copy_from_slice(&new_c2.to_le_bytes());
                    changed = true;
                }

                write_raw_header(&mut writer, &header)?;
                writer.write_raw(&payload)?;
            }
            biff12::SHEETDATA => {
                in_sheet_data = true;
                current_row = None;
                write_raw_header(&mut writer, &header)?;
                copy_exact(&mut input, &mut writer, len)?;
            }
            biff12::SHEETDATA_END => {
                if in_sheet_data {
                    if let Some(row) = current_row {
                        if super::flush_remaining_cells_in_row(
                            &mut writer,
                            edits,
                            &mut applied,
                            &ordered_edits,
                            &mut insert_cursor,
                            row,
                            &mut dim_additions,
                        )? {
                            changed = true;
                        }
                    }
                    if super::flush_remaining_rows(
                        &mut writer,
                        row_template.as_deref(),
                        edits,
                        &mut applied,
                        &ordered_edits,
                        &mut insert_cursor,
                        &mut dim_additions,
                    )? {
                        changed = true;
                    }
                }
                in_sheet_data = false;
                current_row = None;
                write_raw_header(&mut writer, &header)?;
                copy_exact(&mut input, &mut writer, len)?;
            }
            biff12::ROW if in_sheet_data => {
                // Before advancing to a new row, insert any missing cells for the prior row.
                if let Some(row) = current_row {
                    if super::flush_remaining_cells_in_row(
                        &mut writer,
                        edits,
                        &mut applied,
                        &ordered_edits,
                        &mut insert_cursor,
                        row,
                        &mut dim_additions,
                    )? {
                        changed = true;
                    }
                }

                let payload = read_payload(&mut input, len)?;
                if row_template.is_none() && payload.len() >= 4 {
                    row_template = Some(payload.clone());
                }
                let row = super::read_u32(&payload, 0)?;

                // Insert any missing rows before this one (row-major order).
                if super::flush_missing_rows_before(
                    &mut writer,
                    row_template.as_deref(),
                    edits,
                    &mut applied,
                    &ordered_edits,
                    &mut insert_cursor,
                    row,
                    &mut dim_additions,
                )? {
                    changed = true;
                }

                current_row = Some(row);
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
                // Cell records always start with `[col: u32][style: u32]`. For large sheets, most
                // cells are untouched. Avoid allocating a fresh `Vec` for every cell by reading
                // just the fixed prefix, and only materializing the full payload when we need to
                // patch the record.
                if len < 8 {
                    return Err(Error::UnexpectedEof);
                }
                let mut prefix = [0u8; 8];
                input.read_exact(&mut prefix).map_err(super::map_io_error)?;
                let row = current_row.unwrap_or(0);
                let col = u32::from_le_bytes([prefix[0], prefix[1], prefix[2], prefix[3]]);
                let style = u32::from_le_bytes([prefix[4], prefix[5], prefix[6], prefix[7]]);

                if super::flush_missing_cells_before(
                    &mut writer,
                    edits,
                    &mut applied,
                    &ordered_edits,
                    &mut insert_cursor,
                    row,
                    col,
                    &mut dim_additions,
                )? {
                    changed = true;
                }

                let Some(&edit_idx) = edits_by_coord.get(&(row, col)) else {
                    write_raw_header(&mut writer, &header)?;
                    writer.write_raw(&prefix)?;
                    copy_exact(&mut input, &mut writer, len.saturating_sub(prefix.len()))?;
                    continue;
                };

                let mut payload = Vec::new();
                let _ = payload.try_reserve_exact(len.min(READ_PAYLOAD_CHUNK_BYTES));
                payload.extend_from_slice(&prefix);
                read_exact_into_vec(
                    &mut input,
                    &mut payload,
                    len.saturating_sub(prefix.len()),
                )?;

                applied[edit_idx] = true;
                let edit = &edits[edit_idx];
                let style_out = edit.new_style.unwrap_or(style);
                super::advance_insert_cursor(&ordered_edits, &applied, &mut insert_cursor);

                if id == biff12::BLANK
                    && !super::value_edit_is_noop_blank(style, edit)
                    && (!matches!(edit.new_value, CellValue::Blank) || style_out != 0)
                {
                    super::bounds_include(&mut dim_additions, row, col);
                }

                let existing_header = super::ExistingRecordHeader {
                    in_id: header.id,
                    in_len: header.len,
                    id_raw: &header.id_raw,
                    len_raw: &header.len_raw,
                };

                match id {
                    biff12::FORMULA_FLOAT => {
                        if super::formula_edit_is_noop(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            changed = true;
                            super::patch_fmla_num(
                                &mut writer,
                                &payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::FORMULA_STRING => {
                        if super::formula_string_edit_is_noop(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            changed = true;
                            super::patch_fmla_string(
                                &mut writer,
                                &payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::FORMULA_BOOL => {
                        if super::formula_bool_edit_is_noop(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            changed = true;
                            super::patch_fmla_bool(
                                &mut writer,
                                &payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::FORMULA_BOOLERR => {
                        if super::formula_error_edit_is_noop(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            changed = true;
                            super::patch_fmla_error(
                                &mut writer,
                                &payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::FLOAT => {
                        if super::value_edit_is_noop_float(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            super::convert_value_record_to_formula(
                                &mut writer,
                                id,
                                &payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            changed = true;
                            super::patch_fixed_value_cell_preserving_trailing_bytes(
                                &mut writer,
                                id,
                                &payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::NUM => {
                        if super::value_edit_is_noop_rk(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            super::convert_value_record_to_formula(
                                &mut writer,
                                id,
                                &payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            changed = true;
                            super::patch_rk_cell(
                                &mut writer,
                                col,
                                style_out,
                                &payload,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::STRING => {
                        if super::value_edit_is_noop_shared_string(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            super::convert_value_record_to_formula(
                                &mut writer,
                                id,
                                &payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else if let (CellValue::Text(_), Some(isst)) =
                            (&edit.new_value, edit.shared_string_index)
                        {
                            super::reject_formula_payload_edit(edit, row, col)?;

                            changed = true;
                            if payload.len() < 12 {
                                return Err(Error::UnexpectedEof);
                            }

                            // BrtCellIsst: [col: u32][style: u32][isst: u32]
                            //
                            // Preserve the original record header varint bytes (including
                            // non-canonical encodings) and any unknown trailing payload bytes.
                            // This keeps diffs minimal while still updating the referenced `isst`.
                            payload[4..8].copy_from_slice(&style_out.to_le_bytes());
                            payload[8..12].copy_from_slice(&isst.to_le_bytes());
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else {
                            // No shared-string index provided: fall back to the generic writer
                            // (FLOAT / inline string).
                            //
                            // NOTE: When writing a text value, this converts `BrtCellIsst` (shared
                            // string reference) cells into `BrtCellSt` (inline string) because the
                            // streaming patcher cannot update the workbook shared strings part.
                            // Use the shared-strings-aware workbook APIs
                            // (`XlsbWorkbook::save_with_cell_edits_shared_strings` or
                            // `XlsbWorkbook::save_with_cell_edits_streaming_shared_strings`) to
                            // keep shared-string semantics.
                            super::reject_formula_payload_edit(edit, row, col)?;
                            changed = true;
                            super::patch_value_cell(
                                &mut writer,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::CELL_ST => {
                        if super::value_edit_is_noop_inline_string(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            super::convert_value_record_to_formula(
                                &mut writer,
                                id,
                                &payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            super::reject_formula_payload_edit(edit, row, col)?;
                            changed = true;
                            super::patch_cell_st(
                                &mut writer,
                                &payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::BOOL => {
                        if super::value_edit_is_noop_bool(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            super::convert_value_record_to_formula(
                                &mut writer,
                                id,
                                &payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            changed = true;
                            super::patch_fixed_value_cell_preserving_trailing_bytes(
                                &mut writer,
                                id,
                                &payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::BOOLERR => {
                        if super::value_edit_is_noop_error(&payload, edit)? {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            super::convert_value_record_to_formula(
                                &mut writer,
                                id,
                                &payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            changed = true;
                            super::patch_fixed_value_cell_preserving_trailing_bytes(
                                &mut writer,
                                id,
                                &payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::BLANK => {
                        if super::value_edit_is_noop_blank(style, edit) {
                            write_raw_header(&mut writer, &header)?;
                            writer.write_raw(&payload)?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            super::convert_value_record_to_formula(
                                &mut writer,
                                id,
                                &payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            changed = true;
                            super::patch_fixed_value_cell_preserving_trailing_bytes(
                                &mut writer,
                                id,
                                &payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    _ => {
                        if edit.new_formula.is_some() {
                            changed = true;
                            super::convert_value_record_to_formula(
                                &mut writer,
                                id,
                                &payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            super::reject_formula_payload_edit(edit, row, col)?;
                            changed = true;
                            super::patch_value_cell(
                                &mut writer,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                }
            }
            _ => {
                write_raw_header(&mut writer, &header)?;
                copy_exact(&mut input, &mut writer, len)?;
            }
        }
    }

    if in_sheet_data {
        // Best-effort: if the stream ends without `BrtEndSheetData`, still insert any trailing
        // rows/cells and validate that all edits were applied.
        if let Some(row) = current_row {
            if super::flush_remaining_cells_in_row(
                &mut writer,
                edits,
                &mut applied,
                &ordered_edits,
                &mut insert_cursor,
                row,
                &mut dim_additions,
            )? {
                changed = true;
            }
        }
        if super::flush_remaining_rows(
            &mut writer,
            row_template.as_deref(),
            edits,
            &mut applied,
            &ordered_edits,
            &mut insert_cursor,
            &mut dim_additions,
        )? {
            changed = true;
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
            format!("cell edits not applied: {}", missing.join(", ")),
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
    let mut raw = Vec::new();
    let _ = raw.try_reserve_exact(4);

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
        v |= ((byte & 0x7F) as u32) << (7 * i);
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
    let mut raw = Vec::new();
    let _ = raw.try_reserve_exact(4);

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

fn write_raw_header<W: Write>(
    writer: &mut Biff12Writer<W>,
    header: &RawRecordHeader,
) -> io::Result<()> {
    writer.write_raw(&header.id_raw)?;
    writer.write_raw(&header.len_raw)?;
    Ok(())
}

fn read_payload<R: Read>(r: &mut R, len: usize) -> Result<Vec<u8>, Error> {
    // Record lengths are attacker-controlled. Avoid allocating the full payload up-front; instead
    // grow the buffer as bytes are successfully read (important when the underlying stream is
    // truncated or size-limited).
    let mut payload = Vec::new();
    let _ = payload.try_reserve_exact(len.min(READ_PAYLOAD_CHUNK_BYTES));
    read_exact_into_vec(r, &mut payload, len)?;
    Ok(payload)
}

fn read_exact_into_vec<R: Read>(r: &mut R, out: &mut Vec<u8>, mut len: usize) -> Result<(), Error> {
    while len > 0 {
        let chunk_len = READ_PAYLOAD_CHUNK_BYTES.min(len);
        let start = out.len();
        out.resize(start + chunk_len, 0);
        if let Err(err) = r.read_exact(&mut out[start..]) {
            out.truncate(start);
            return Err(super::map_io_error(err));
        }
        len = len.saturating_sub(chunk_len);
    }
    Ok(())
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

fn copy_remaining<R: Read, W: Write>(
    input: &mut R,
    writer: &mut Biff12Writer<W>,
) -> Result<(), Error> {
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
