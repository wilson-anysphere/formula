use std::collections::HashMap;
use std::io::{self, Cursor};

use crate::biff12_varint;
use crate::parser::{biff12, CellValue, Error};
use crate::strings::FlagsWidth;
use crate::writer::Biff12Writer;

mod streaming;

pub use streaming::patch_sheet_bin_streaming;

// BIFF12 "wide string" flags (used by BrtCellSt inline strings and BrtFmlaString cached values).
const FLAG_RICH: u16 = 0x0001;
const FLAG_PHONETIC: u16 = 0x0002;

// Size (in bytes) of a single rich text formatting run in BIFF12.
//
// `StrRun` entries are 8 bytes:
//   [ich: u32][ifnt: u16][reserved: u16]
const RICH_RUN_BYTE_LEN: usize = 8;

/// Raw BIFF12 record header bytes for an existing record in the input stream.
///
/// When patching an existing record, we sometimes want to preserve the original varint
/// encodings of the record id and length fields (including non-canonical encodings) as long as
/// the patched record keeps the same id and payload length.
#[derive(Clone, Copy)]
struct ExistingRecordHeader<'a> {
    in_id: u32,
    in_len: u32,
    id_raw: &'a [u8],
    len_raw: &'a [u8],
}

fn write_record_header_preserving_varints<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    out_id: u32,
    out_len: u32,
    existing: Option<ExistingRecordHeader<'_>>,
) -> io::Result<()> {
    if let Some(existing) = existing {
        if existing.in_id == out_id && existing.in_len == out_len {
            writer.write_raw(existing.id_raw)?;
            writer.write_raw(existing.len_raw)?;
            return Ok(());
        }
    }
    writer.write_record_header(out_id, out_len)
}

fn write_record_preserving_varints<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    out_id: u32,
    payload: &[u8],
    existing: Option<ExistingRecordHeader<'_>>,
) -> io::Result<()> {
    let out_len = u32::try_from(payload.len()).map_err(|_| {
        io::Error::new(
            io::ErrorKind::InvalidInput,
            "record payload length does not fit in u32",
        )
    })?;
    write_record_header_preserving_varints(writer, out_id, out_len, existing)?;
    writer.write_raw(payload)
}

// Excel worksheet grid limits (0-based, inclusive).
//
// These are fixed by the XLSX/XLSB spec: 1,048,576 rows and 16,384 columns (A..XFD).
const EXCEL_MAX_ROW: u32 = 1_048_575;
const EXCEL_MAX_COL: u32 = 16_383;

/// A single cell update to apply while patch-writing a worksheet `.bin` part.
///
/// Row/col are zero-based, matching the XLSB internal representation used by the parser.
#[derive(Debug, Clone)]
pub struct CellEdit {
    pub row: u32,
    pub col: u32,
    pub new_value: CellValue,
    /// Optional style (XF) index override.
    ///
    /// - When editing an existing cell record, `None` preserves the existing style index.
    /// - When inserting a new cell record, `None` uses style index `0` (the default / "Normal"
    ///   style), matching historical behavior.
    pub new_style: Option<u32>,
    /// If true, drop any existing formula payload and write this cell as a plain value cell.
    ///
    /// This is the "paste values" operation for formula cells: the output worksheet stream will
    /// contain `BrtCell*`/`BrtBlank` records instead of `BrtFmla*`, so re-reading the workbook will
    /// yield `cell.formula == None`.
    ///
    /// When false (the default), formula cells keep their formula unless [`Self::new_formula`] is
    /// set to replace it. (Historical behavior: setting `new_value=Blank` with no formula payload
    /// also clears the formula by rewriting the record to `BrtBlank`.)
    ///
    /// Non-formula cells ignore this flag.
    ///
    /// When `clear_formula` is true, [`Self::new_formula`], [`Self::new_rgcb`], and
    /// [`Self::new_formula_flags`] must all be `None`.
    pub clear_formula: bool,
    /// If set, replaces the raw formula token stream (`rgce`) for formula cells.
    pub new_formula: Option<Vec<u8>>,
    /// If set, replaces the trailing BIFF12 `rgcb` payload (a.k.a. `Formula.extra`) that appears
    /// after the `rgce` token stream for some formulas (e.g. array constants / `PtgArray`).
    ///
    /// If `None`, the existing bytes are preserved.
    pub new_rgcb: Option<Vec<u8>>,
    /// If set, replaces the raw BIFF12 formula flags (`BrtFmla*.flags`, `grbitFmla`) for formula
    /// cells.
    ///
    /// - When patching an existing formula cell and this is `None`, the existing flags are
    ///   preserved.
    /// - When inserting a new formula cell and this is `None`, flags default to `0`.
    pub new_formula_flags: Option<u16>,
    /// Optional shared string table index (`isst`) to use when writing `CellValue::Text`.
    ///
    /// XLSB can store text cells either as inline strings (`BrtCellSt`, record id `0x0006`)
    /// or as shared-string references (`BrtCellIsst`, record id `0x0007` / [`biff12::STRING`]).
    ///
    /// When patching an existing `BrtCellIsst` record, providing `shared_string_index` lets the
    /// patcher keep the cell as a shared-string reference. When this is `None`, the patcher
    /// falls back to writing an inline string because it has no access to (or ability to update)
    /// the workbook's shared strings part.
    pub shared_string_index: Option<u32>,
}

fn validate_cell_edits(edits: &[CellEdit]) -> Result<(), Error> {
    for edit in edits {
        if edit.row > EXCEL_MAX_ROW || edit.col > EXCEL_MAX_COL {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!(
                    "cell edit coordinate out of range: row={}, col={} (max row={}, max col={})",
                    edit.row, edit.col, EXCEL_MAX_ROW, EXCEL_MAX_COL
                ),
            )));
        }
    }
    Ok(())
}

impl CellEdit {
    /// Convenience helper for updating a formula cell from Excel formula text using workbook
    /// context.
    ///
    /// When the `write` feature is enabled, this prefers the AST-based encoder
    /// ([`crate::rgce::encode_rgce_with_context_ast`]) for improved grammar coverage. Without
    /// `write`, it falls back to [`crate::rgce::encode_rgce_with_context`].
    ///
    /// Note: table-less structured references like `[@Qty]` can only be encoded when the table id
    /// is unambiguous. If the workbook context contains multiple tables, prefer
    /// [`Self::with_formula_text_with_context_in_sheet`] so the encoder can infer the correct
    /// table id from the sheet + base cell location.
    ///
    /// `formula` may include a leading `=`.
    pub fn with_formula_text_with_context(
        row: u32,
        col: u32,
        new_value: CellValue,
        formula: &str,
        ctx: &crate::workbook_context::WorkbookContext,
    ) -> Result<Self, crate::rgce::EncodeError> {
        let base = crate::rgce::CellCoord::new(row, col);
        let encoded = {
            #[cfg(feature = "write")]
            {
                crate::rgce::encode_rgce_with_context_ast(formula, ctx, base)?
            }
            #[cfg(not(feature = "write"))]
            {
                crate::rgce::encode_rgce_with_context(formula, ctx, base)?
            }
        };
        Ok(Self {
            row,
            col,
            new_value,
            new_style: None,
            clear_formula: false,
            new_formula: Some(encoded.rgce),
            new_rgcb: Some(encoded.rgcb),
            new_formula_flags: None,
            shared_string_index: None,
        })
    }

    /// Convenience helper for updating a formula cell from Excel formula text using workbook
    /// context, plus a sheet name for table inference.
    ///
    /// This is useful for table-less structured references like `[@Col]` in workbooks that contain
    /// multiple tables; we can infer the correct table id from the base cell location (the cell
    /// must be inside exactly one table range on that sheet).
    ///
    /// `formula` may include a leading `=`.
    pub fn with_formula_text_with_context_in_sheet(
        row: u32,
        col: u32,
        new_value: CellValue,
        formula: &str,
        sheet: &str,
        ctx: &crate::workbook_context::WorkbookContext,
    ) -> Result<Self, crate::rgce::EncodeError> {
        let base = crate::rgce::CellCoord::new(row, col);
        let encoded = {
            #[cfg(feature = "write")]
            {
                crate::rgce::encode_rgce_with_context_ast_in_sheet(formula, ctx, sheet, base)?
            }
            #[cfg(not(feature = "write"))]
            {
                // Without the AST encoder, the sheet name currently only affects structured
                // reference inference (not supported), so fall back to the legacy path.
                let _ = sheet;
                crate::rgce::encode_rgce_with_context(formula, ctx, base)?
            }
        };
        Ok(Self {
            row,
            col,
            new_value,
            new_style: None,
            clear_formula: false,
            new_formula: Some(encoded.rgce),
            new_rgcb: Some(encoded.rgcb),
            new_formula_flags: None,
            shared_string_index: None,
        })
    }

    /// Replace `new_formula` + `new_rgcb` by encoding the provided formula text using workbook
    /// context.
    ///
    /// When the `write` feature is enabled, this prefers the AST-based encoder
    /// ([`crate::rgce::encode_rgce_with_context_ast`]) for improved grammar coverage. Without
    /// `write`, it falls back to [`crate::rgce::encode_rgce_with_context`].
    ///
    /// Note: if `formula` uses table-less structured references like `[@Qty]` and the workbook
    /// context contains multiple tables, prefer [`Self::set_formula_text_with_context_in_sheet`]
    /// so the encoder can infer the correct table id from the sheet + base cell location.
    ///
    /// `formula` may include a leading `=`.
    pub fn set_formula_text_with_context(
        &mut self,
        formula: &str,
        ctx: &crate::workbook_context::WorkbookContext,
    ) -> Result<(), crate::rgce::EncodeError> {
        let base = crate::rgce::CellCoord::new(self.row, self.col);
        let encoded = {
            #[cfg(feature = "write")]
            {
                crate::rgce::encode_rgce_with_context_ast(formula, ctx, base)?
            }
            #[cfg(not(feature = "write"))]
            {
                crate::rgce::encode_rgce_with_context(formula, ctx, base)?
            }
        };
        self.new_formula = Some(encoded.rgce);
        self.new_rgcb = Some(encoded.rgcb);
        self.clear_formula = false;
        Ok(())
    }

    /// Replace `new_formula` + `new_rgcb` by encoding the provided formula text using workbook
    /// context, plus a sheet name for table inference.
    ///
    /// See [`Self::with_formula_text_with_context_in_sheet`] for details.
    ///
    /// `formula` may include a leading `=`.
    pub fn set_formula_text_with_context_in_sheet(
        &mut self,
        formula: &str,
        sheet: &str,
        ctx: &crate::workbook_context::WorkbookContext,
    ) -> Result<(), crate::rgce::EncodeError> {
        let base = crate::rgce::CellCoord::new(self.row, self.col);
        let encoded = {
            #[cfg(feature = "write")]
            {
                crate::rgce::encode_rgce_with_context_ast_in_sheet(formula, ctx, sheet, base)?
            }
            #[cfg(not(feature = "write"))]
            {
                let _ = sheet;
                crate::rgce::encode_rgce_with_context(formula, ctx, base)?
            }
        };
        self.new_formula = Some(encoded.rgce);
        self.new_rgcb = Some(encoded.rgcb);
        self.clear_formula = false;
        Ok(())
    }
}

#[cfg(feature = "write")]
impl CellEdit {
    /// Convenience helper for updating a formula cell from Excel formula text.
    ///
    /// `formula` may include a leading `=`.
    pub fn with_formula_text(
        row: u32,
        col: u32,
        new_value: CellValue,
        formula: &str,
    ) -> Result<Self, formula_biff::EncodeRgceError> {
        let encoded = formula_biff::encode_rgce_with_rgcb(formula)?;
        Ok(Self {
            row,
            col,
            new_value,
            new_style: None,
            clear_formula: false,
            new_formula_flags: None,
            new_formula: Some(encoded.rgce),
            new_rgcb: Some(encoded.rgcb),
            shared_string_index: None,
        })
    }

    /// Replace `new_formula` by encoding the provided formula text.
    pub fn set_formula_text(&mut self, formula: &str) -> Result<(), formula_biff::EncodeRgceError> {
        let encoded = formula_biff::encode_rgce_with_rgcb(formula)?;
        self.new_formula = Some(encoded.rgce);
        self.new_rgcb = Some(encoded.rgcb);
        self.clear_formula = false;
        Ok(())
    }
}

/// Patch a worksheet stream (`xl/worksheets/sheetN.bin`) by rewriting only the targeted
/// cell records inside `BrtSheetData`, while copying every other record byte-for-byte.
///
/// This is a minimal bridge between the current read-only XLSB implementation and a
/// full writer:
/// - inserts missing rows/cells inside `BrtSheetData` (row-major order)
/// - updates existing cells that appear in the stream
/// - rewrites only selected supported cell record types
pub fn patch_sheet_bin(sheet_bin: &[u8], edits: &[CellEdit]) -> Result<Vec<u8>, Error> {
    if edits.is_empty() {
        return Ok(sheet_bin.to_vec());
    }

    validate_cell_edits(edits)?;

    // When BrtWsDim is missing from the input stream, we may need to synthesize it so the
    // worksheet used-range metadata stays consistent after inserting non-blank cells.
    //
    // Mirror the streaming patcher's conservative bounding-box logic: only include edits that
    // materialize a non-blank cell after patching.
    let mut requested_bounds: Option<Bounds> = None;
    for edit in edits {
        if insertion_is_noop(edit) {
            continue;
        }
        if matches!(edit.new_value, CellValue::Blank) {
            continue;
        }
        bounds_include(&mut requested_bounds, edit.row, edit.col);
    }

    let mut edits_by_coord: HashMap<(u32, u32), usize> = HashMap::with_capacity(edits.len());
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

    let mut out = Vec::with_capacity(sheet_bin.len());
    let mut writer = Biff12Writer::new(&mut out);

    let mut offset = 0usize;
    let mut in_sheet_data = false;
    let mut current_row: Option<u32> = None;
    let mut row_template: Option<Vec<u8>> = None;
    let mut dim_record: Option<DimensionRecordInfo> = None;
    let mut dim_additions: Option<Bounds> = None;
    let mut changed = false;
    let mut dim_insert_offset: Option<usize> = None;
    let mut observed_cell_bounds: Option<Bounds> = None;

    while offset < sheet_bin.len() {
        let record_start = offset;
        let id = read_record_id(sheet_bin, &mut offset)?;
        let id_end = offset;
        let len_u32 = read_record_len(sheet_bin, &mut offset)?;
        let len = len_u32 as usize;
        let payload_start = offset;
        let payload_end = payload_start.checked_add(len).ok_or(Error::UnexpectedEof)?;
        let payload = sheet_bin
            .get(payload_start..payload_end)
            .ok_or(Error::UnexpectedEof)?;
        offset = payload_end;

        let record_end = payload_end;
        let header_len = payload_start
            .checked_sub(record_start)
            .ok_or(Error::UnexpectedEof)?;

        match id {
            biff12::DIMENSION => {
                if len < 16 {
                    return Err(Error::UnexpectedEof);
                }
                // BrtWsDim: [r1: u32][r2: u32][c1: u32][c2: u32]
                let r1 = read_u32(payload, 0)?;
                let r2 = read_u32(payload, 4)?;
                let c1 = read_u32(payload, 8)?;
                let c2 = read_u32(payload, 12)?;

                // Record offsets may diverge from the input stream once we start patching and/or
                // inserting records. Capture the output offset before writing the raw DIMENSION
                // record so we can patch the payload in-place later even if BrtWsDim appears
                // after other rewritten records (non-standard but possible in malformed streams).
                let out_record_start = writer.bytes_written();
                dim_record = Some(DimensionRecordInfo {
                    out_payload_offset: out_record_start
                        .checked_add(header_len)
                        .ok_or(Error::UnexpectedEof)?,
                    r1,
                    r2,
                    c1,
                    c2,
                });
                writer.write_raw(&sheet_bin[record_start..record_end])?;
            }
            biff12::SHEETDATA => {
                if dim_insert_offset.is_none() {
                    dim_insert_offset = Some(writer.bytes_written());
                }
                in_sheet_data = true;
                current_row = None;
                writer.write_raw(&sheet_bin[record_start..record_end])?;
            }
            biff12::SHEETDATA_END => {
                if in_sheet_data {
                    // Flush any trailing cell inserts for the final row before leaving SheetData.
                    if let Some(row) = current_row {
                        if flush_remaining_cells_in_row(
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
                    // And append any remaining missing rows/cells.
                    if flush_remaining_rows(
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
                writer.write_raw(&sheet_bin[record_start..record_end])?;
            }
            biff12::ROW if in_sheet_data => {
                // Before advancing to a new row, insert any missing cells for the prior row.
                if let Some(row) = current_row {
                    if flush_remaining_cells_in_row(
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

                if row_template.is_none() && payload.len() >= 4 {
                    row_template = Some(payload.to_vec());
                }

                let row = read_u32(payload, 0)?;
                // Insert any missing rows before this one (row-major order).
                if flush_missing_rows_before(
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
                writer.write_raw(&sheet_bin[record_start..record_end])?;
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
                let row = current_row.unwrap_or(0);
                let col = read_u32(payload, 0)?;
                let style = read_u32(payload, 4)?;

                // Best-effort: if BrtWsDim is missing, treat any existing cell records as part of
                // the used range so a synthesized DIMENSION covers both the edits and any
                // pre-existing cells.
                bounds_include(&mut observed_cell_bounds, row, col);

                // Insert any missing cells that should appear before this one.
                if flush_missing_cells_before(
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
                    writer.write_raw(&sheet_bin[record_start..record_end])?;
                    continue;
                };

                applied[edit_idx] = true;
                let edit = &edits[edit_idx];
                let style_out = edit.new_style.unwrap_or(style);
                advance_insert_cursor(&ordered_edits, &applied, &mut insert_cursor);

                let existing_header = ExistingRecordHeader {
                    in_id: id,
                    in_len: len_u32,
                    id_raw: sheet_bin
                        .get(record_start..id_end)
                        .ok_or(Error::UnexpectedEof)?,
                    len_raw: sheet_bin
                        .get(id_end..payload_start)
                        .ok_or(Error::UnexpectedEof)?,
                };

                // Track used-range expansion for edits that turn an empty cell into a value.
                if id == biff12::BLANK && !matches!(edit.new_value, CellValue::Blank) {
                    bounds_include(&mut dim_additions, row, col);
                }

                match id {
                    biff12::FORMULA_FLOAT => {
                        // Preserve the original bytes when the edit is a no-op. This avoids
                        // unintentionally changing record header encodings or triggering
                        // downstream "edited" heuristics (e.g. calcChain invalidation).
                        if formula_edit_is_noop(payload, edit)? {
                            writer.write_raw(&sheet_bin[record_start..record_end])?;
                        } else {
                            changed = true;
                            patch_fmla_num(
                                &mut writer,
                                payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::FORMULA_STRING => {
                        if formula_string_edit_is_noop(payload, edit)? {
                            writer.write_raw(&sheet_bin[record_start..record_end])?;
                        } else {
                            changed = true;
                            patch_fmla_string(
                                &mut writer,
                                payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::FORMULA_BOOL => {
                        if formula_bool_edit_is_noop(payload, edit)? {
                            writer.write_raw(&sheet_bin[record_start..record_end])?;
                        } else {
                            changed = true;
                            patch_fmla_bool(
                                &mut writer,
                                payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::FORMULA_BOOLERR => {
                        if formula_error_edit_is_noop(payload, edit)? {
                            writer.write_raw(&sheet_bin[record_start..record_end])?;
                        } else {
                            changed = true;
                            patch_fmla_error(
                                &mut writer,
                                payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::FLOAT => {
                        if value_edit_is_noop_float(payload, edit)? {
                            writer.write_raw(&sheet_bin[record_start..record_end])?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            convert_value_record_to_formula(
                                &mut writer,
                                id,
                                payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            changed = true;
                            patch_fixed_value_cell_preserving_trailing_bytes(
                                &mut writer,
                                id,
                                payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::NUM => {
                        if value_edit_is_noop_rk(payload, edit)? {
                            writer.write_raw(&sheet_bin[record_start..record_end])?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            convert_value_record_to_formula(
                                &mut writer,
                                id,
                                payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            changed = true;
                            patch_rk_cell(
                                &mut writer,
                                col,
                                style_out,
                                payload,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::CELL_ST => {
                        if value_edit_is_noop_inline_string(payload, edit)? {
                            writer.write_raw(&sheet_bin[record_start..record_end])?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            convert_value_record_to_formula(
                                &mut writer,
                                id,
                                payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            reject_formula_payload_edit(edit, row, col)?;
                            changed = true;
                            patch_cell_st(
                                &mut writer,
                                payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::STRING => {
                        if value_edit_is_noop_shared_string(payload, edit)? {
                            writer.write_raw(&sheet_bin[record_start..record_end])?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            convert_value_record_to_formula(
                                &mut writer,
                                id,
                                payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else if let (CellValue::Text(_), Some(isst)) =
                            (&edit.new_value, edit.shared_string_index)
                        {
                            reject_formula_payload_edit(edit, row, col)?;
                            changed = true;
                            if payload.len() < 12 {
                                return Err(Error::UnexpectedEof);
                            }

                            // BrtCellIsst: [col: u32][style: u32][isst: u32]
                            //
                            // Preserve the original record header varint bytes (including
                            // non-canonical encodings) and any unknown trailing payload bytes.
                            // This keeps diffs minimal while still updating the referenced `isst`.
                            let mut patched = payload.to_vec();
                            patched[4..8].copy_from_slice(&style_out.to_le_bytes());
                            patched[8..12].copy_from_slice(&isst.to_le_bytes());
                            writer.write_raw(&sheet_bin[record_start..payload_start])?;
                            writer.write_raw(&patched)?;
                        } else {
                            // No shared-string index provided: fall back to the generic writer
                            // (FLOAT / inline string).
                            //
                            // NOTE: This converts `BrtCellIsst` (shared string reference) cells
                            // into `BrtCellSt` (inline string) when writing a text value. Use the
                            // shared-strings-aware workbook APIs
                            // (`XlsbWorkbook::save_with_cell_edits_shared_strings` or
                            // `XlsbWorkbook::save_with_cell_edits_streaming_shared_strings`) to
                            // keep shared-string semantics.
                            reject_formula_payload_edit(edit, row, col)?;
                            changed = true;
                            patch_value_cell(
                                &mut writer,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::BOOL => {
                        if value_edit_is_noop_bool(payload, edit)? {
                            writer.write_raw(&sheet_bin[record_start..record_end])?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            convert_value_record_to_formula(
                                &mut writer,
                                id,
                                payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            changed = true;
                            patch_fixed_value_cell_preserving_trailing_bytes(
                                &mut writer,
                                id,
                                payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::BOOLERR => {
                        if value_edit_is_noop_error(payload, edit)? {
                            writer.write_raw(&sheet_bin[record_start..record_end])?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            convert_value_record_to_formula(
                                &mut writer,
                                id,
                                payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            changed = true;
                            patch_fixed_value_cell_preserving_trailing_bytes(
                                &mut writer,
                                id,
                                payload,
                                col,
                                style_out,
                                edit,
                                Some(existing_header),
                            )?;
                        }
                    }
                    biff12::BLANK => {
                        if value_edit_is_noop_blank(style, edit) {
                            writer.write_raw(&sheet_bin[record_start..record_end])?;
                        } else if edit.new_formula.is_some() {
                            changed = true;
                            convert_value_record_to_formula(
                                &mut writer,
                                id,
                                payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            changed = true;
                            patch_fixed_value_cell_preserving_trailing_bytes(
                                &mut writer,
                                id,
                                payload,
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
                            convert_value_record_to_formula(
                                &mut writer,
                                id,
                                payload,
                                col,
                                style_out,
                                edit,
                            )?;
                        } else {
                            reject_formula_payload_edit(edit, row, col)?;
                            changed = true;
                            patch_value_cell(
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
            biff12::WORKSHEET_END => {
                if dim_insert_offset.is_none() {
                    dim_insert_offset = Some(writer.bytes_written());
                }
                writer.write_raw(&sheet_bin[record_start..record_end])?;
            }
            _ => {
                writer.write_raw(&sheet_bin[record_start..record_end])?;
            }
        }
    }

    if in_sheet_data {
        // Worksheet streams should always close `BrtSheetData`, but if they don't, still make a
        // best-effort attempt to apply all edits.
        if let Some(row) = current_row {
            if flush_remaining_cells_in_row(
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
        if flush_remaining_rows(
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

    drop(writer);
    if let Some(dim_record) = dim_record {
        if let Some(additions) = dim_additions {
            let mut r1 = dim_record.r1;
            let mut r2 = dim_record.r2;
            let mut c1 = dim_record.c1;
            let mut c2 = dim_record.c2;

            r1 = r1.min(additions.min_row);
            r2 = r2.max(additions.max_row);
            c1 = c1.min(additions.min_col);
            c2 = c2.max(additions.max_col);

            if (r1, r2, c1, c2) != (dim_record.r1, dim_record.r2, dim_record.c1, dim_record.c2) {
                let start = dim_record.out_payload_offset;
                let end = start.checked_add(16).ok_or(Error::UnexpectedEof)?;
                let payload = out.get_mut(start..end).ok_or(Error::UnexpectedEof)?;
                payload[0..4].copy_from_slice(&r1.to_le_bytes());
                payload[4..8].copy_from_slice(&r2.to_le_bytes());
                payload[8..12].copy_from_slice(&c1.to_le_bytes());
                payload[12..16].copy_from_slice(&c2.to_le_bytes());
            }
        }
    } else if changed {
        if let Some(mut bounds) = requested_bounds {
            if let Some(additions) = dim_additions {
                bounds.min_row = bounds.min_row.min(additions.min_row);
                bounds.max_row = bounds.max_row.max(additions.max_row);
                bounds.min_col = bounds.min_col.min(additions.min_col);
                bounds.max_col = bounds.max_col.max(additions.max_col);
            }
            if let Some(observed) = observed_cell_bounds {
                bounds.min_row = bounds.min_row.min(observed.min_row);
                bounds.max_row = bounds.max_row.max(observed.max_row);
                bounds.min_col = bounds.min_col.min(observed.min_col);
                bounds.max_col = bounds.max_col.max(observed.max_col);
            }

            // No BrtWsDim record was present in the input. Synthesize a new one so consumers that
            // rely on the worksheet used-range metadata (including our own parser) can observe the
            // edits.
            //
            // Only do this when we actually changed the sheet stream; for a complete no-op patch we
            // keep the output byte-identical to the input.
            let insert_at = dim_insert_offset.unwrap_or(out.len());
            if insert_at > out.len() {
                return Err(Error::UnexpectedEof);
            }

            let mut dim_bytes = Vec::with_capacity(24);
            {
                let mut w = Biff12Writer::new(&mut dim_bytes);
                w.write_record_header(biff12::DIMENSION, 16)?;
                w.write_u32(bounds.min_row)?;
                w.write_u32(bounds.max_row)?;
                w.write_u32(bounds.min_col)?;
                w.write_u32(bounds.max_col)?;
            }
            out.splice(insert_at..insert_at, dim_bytes);
        }
    }

    Ok(out)
}

#[derive(Debug, Clone, Copy)]
struct DimensionRecordInfo {
    out_payload_offset: usize,
    r1: u32,
    r2: u32,
    c1: u32,
    c2: u32,
}

#[derive(Debug, Clone, Copy)]
struct Bounds {
    min_row: u32,
    max_row: u32,
    min_col: u32,
    max_col: u32,
}

fn bounds_include(bounds: &mut Option<Bounds>, row: u32, col: u32) {
    match bounds {
        Some(existing) => {
            existing.min_row = existing.min_row.min(row);
            existing.max_row = existing.max_row.max(row);
            existing.min_col = existing.min_col.min(col);
            existing.max_col = existing.max_col.max(col);
        }
        None => {
            *bounds = Some(Bounds {
                min_row: row,
                max_row: row,
                min_col: col,
                max_col: col,
            });
        }
    }
}

fn advance_insert_cursor(ordered: &[usize], applied: &[bool], cursor: &mut usize) {
    while *cursor < ordered.len() && applied[ordered[*cursor]] {
        *cursor += 1;
    }
}

fn insertion_is_noop(edit: &CellEdit) -> bool {
    edit.new_formula.is_none()
        && edit.new_rgcb.is_none()
        && edit.new_style.is_none()
        && matches!(edit.new_value, CellValue::Blank)
}

fn write_row_record<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    row: u32,
    template: Option<&[u8]>,
) -> Result<(), Error> {
    if let Some(template) = template {
        if template.len() >= 4 {
            let len = u32::try_from(template.len()).map_err(|_| {
                Error::Io(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "ROW record template payload length does not fit in u32",
                ))
            })?;
            writer.write_record_header(biff12::ROW, len)?;
            writer.write_u32(row)?;
            writer.write_raw(&template[4..])?;
            return Ok(());
        }
    }

    writer.write_record(biff12::ROW, &row.to_le_bytes())?;
    Ok(())
}

fn flush_missing_rows_before<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    row_template: Option<&[u8]>,
    edits: &[CellEdit],
    applied: &mut [bool],
    ordered: &[usize],
    cursor: &mut usize,
    before_row: u32,
    dim_additions: &mut Option<Bounds>,
) -> Result<bool, Error> {
    advance_insert_cursor(ordered, applied, cursor);
    let mut wrote_any = false;

    while *cursor < ordered.len() {
        let idx = ordered[*cursor];
        if applied[idx] {
            *cursor += 1;
            continue;
        }

        let edit = &edits[idx];
        if edit.row >= before_row {
            break;
        }

        let row = edit.row;
        // If the row doesn't exist in the stream and all pending edits for this row are
        // "clear" operations, treat them as no-ops and avoid materializing an empty row.
        let mut should_write_row = false;
        {
            let mut scan = *cursor;
            while scan < ordered.len() {
                let idx = ordered[scan];
                if applied[idx] {
                    scan += 1;
                    continue;
                }
                let edit = &edits[idx];
                if edit.row != row {
                    break;
                }
                if !insertion_is_noop(edit) {
                    should_write_row = true;
                    break;
                }
                scan += 1;
            }
        }

        if !should_write_row {
            while *cursor < ordered.len() {
                let idx = ordered[*cursor];
                if applied[idx] {
                    *cursor += 1;
                    continue;
                }
                let edit = &edits[idx];
                if edit.row != row {
                    break;
                }
                applied[idx] = true;
                *cursor += 1;
            }
            advance_insert_cursor(ordered, applied, cursor);
            continue;
        }

        write_row_record(writer, row, row_template)?;
        wrote_any = true;

        while *cursor < ordered.len() {
            let idx = ordered[*cursor];
            if applied[idx] {
                *cursor += 1;
                continue;
            }
            let edit = &edits[idx];
            if edit.row != row {
                break;
            }

            if insertion_is_noop(edit) {
                applied[idx] = true;
                *cursor += 1;
                continue;
            }

            write_new_cell_record(writer, edit.col, edit)?;
            wrote_any = true;
            applied[idx] = true;
            if !matches!(edit.new_value, CellValue::Blank) {
                bounds_include(dim_additions, edit.row, edit.col);
            }
            *cursor += 1;
        }

        advance_insert_cursor(ordered, applied, cursor);
    }

    Ok(wrote_any)
}

fn flush_missing_cells_before<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    edits: &[CellEdit],
    applied: &mut [bool],
    ordered: &[usize],
    cursor: &mut usize,
    row: u32,
    before_col: u32,
    dim_additions: &mut Option<Bounds>,
) -> Result<bool, Error> {
    advance_insert_cursor(ordered, applied, cursor);
    let mut wrote_any = false;

    while *cursor < ordered.len() {
        let idx = ordered[*cursor];
        if applied[idx] {
            *cursor += 1;
            continue;
        }

        let edit = &edits[idx];
        if edit.row != row || edit.col >= before_col {
            break;
        }

        if insertion_is_noop(edit) {
            applied[idx] = true;
            *cursor += 1;
            advance_insert_cursor(ordered, applied, cursor);
            continue;
        }

        write_new_cell_record(writer, edit.col, edit)?;
        wrote_any = true;
        applied[idx] = true;
        if !matches!(edit.new_value, CellValue::Blank) {
            bounds_include(dim_additions, edit.row, edit.col);
        }
        *cursor += 1;
        advance_insert_cursor(ordered, applied, cursor);
    }

    Ok(wrote_any)
}

fn flush_remaining_cells_in_row<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    edits: &[CellEdit],
    applied: &mut [bool],
    ordered: &[usize],
    cursor: &mut usize,
    row: u32,
    dim_additions: &mut Option<Bounds>,
) -> Result<bool, Error> {
    advance_insert_cursor(ordered, applied, cursor);
    let mut wrote_any = false;

    while *cursor < ordered.len() {
        let idx = ordered[*cursor];
        if applied[idx] {
            *cursor += 1;
            continue;
        }

        let edit = &edits[idx];
        if edit.row != row {
            break;
        }

        if insertion_is_noop(edit) {
            applied[idx] = true;
            *cursor += 1;
            advance_insert_cursor(ordered, applied, cursor);
            continue;
        }

        write_new_cell_record(writer, edit.col, edit)?;
        wrote_any = true;
        applied[idx] = true;
        if !matches!(edit.new_value, CellValue::Blank) {
            bounds_include(dim_additions, edit.row, edit.col);
        }
        *cursor += 1;
        advance_insert_cursor(ordered, applied, cursor);
    }

    Ok(wrote_any)
}

fn flush_remaining_rows<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    row_template: Option<&[u8]>,
    edits: &[CellEdit],
    applied: &mut [bool],
    ordered: &[usize],
    cursor: &mut usize,
    dim_additions: &mut Option<Bounds>,
) -> Result<bool, Error> {
    advance_insert_cursor(ordered, applied, cursor);
    let mut wrote_any = false;

    while *cursor < ordered.len() {
        let idx = ordered[*cursor];
        if applied[idx] {
            *cursor += 1;
            continue;
        }

        let row = edits[idx].row;
        let mut should_write_row = false;
        {
            let mut scan = *cursor;
            while scan < ordered.len() {
                let idx = ordered[scan];
                if applied[idx] {
                    scan += 1;
                    continue;
                }
                let edit = &edits[idx];
                if edit.row != row {
                    break;
                }
                if !insertion_is_noop(edit) {
                    should_write_row = true;
                    break;
                }
                scan += 1;
            }
        }

        if !should_write_row {
            while *cursor < ordered.len() {
                let idx = ordered[*cursor];
                if applied[idx] {
                    *cursor += 1;
                    continue;
                }
                let edit = &edits[idx];
                if edit.row != row {
                    break;
                }
                applied[idx] = true;
                *cursor += 1;
            }
            advance_insert_cursor(ordered, applied, cursor);
            continue;
        }

        write_row_record(writer, row, row_template)?;
        wrote_any = true;

        while *cursor < ordered.len() {
            let idx = ordered[*cursor];
            if applied[idx] {
                *cursor += 1;
                continue;
            }

            let edit = &edits[idx];
            if edit.row != row {
                break;
            }

            if insertion_is_noop(edit) {
                applied[idx] = true;
                *cursor += 1;
                continue;
            }

            write_new_cell_record(writer, edit.col, edit)?;
            wrote_any = true;
            applied[idx] = true;
            if !matches!(edit.new_value, CellValue::Blank) {
                bounds_include(dim_additions, edit.row, edit.col);
            }
            *cursor += 1;
        }

        advance_insert_cursor(ordered, applied, cursor);
    }

    Ok(wrote_any)
}

fn write_new_cell_record<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    col: u32,
    edit: &CellEdit,
) -> Result<(), Error> {
    let style = edit.new_style.unwrap_or(0u32);
    let rgcb = edit.new_rgcb.as_deref().unwrap_or(&[]);
    if let Some(rgce) = edit.new_formula.as_deref() {
        if rgcb.is_empty() && rgce_references_rgcb(rgce) {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!(
                    "formula update for cell at ({}, {}) requires rgcb bytes (PtgArray present); set CellEdit.new_rgcb",
                    edit.row, edit.col
                ),
            )));
        }
    }
    let flags = edit.new_formula_flags.unwrap_or(0);
    match (&edit.new_formula, &edit.new_value) {
        (Some(rgce), CellValue::Number(v)) => {
            write_new_fmla_num(writer, col, style, *v, flags, rgce, rgcb)
        }
        (Some(rgce), CellValue::Bool(v)) => {
            write_new_fmla_bool(writer, col, style, *v, flags, rgce, rgcb)
        }
        (Some(rgce), CellValue::Error(v)) => {
            write_new_fmla_error(writer, col, style, *v, flags, rgce, rgcb)
        }
        (Some(rgce), CellValue::Text(s)) => {
            write_new_fmla_string(writer, col, style, s, flags, rgce, rgcb)
        }
        (Some(_), CellValue::Blank) => Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "cannot write formula cell with blank cached value at ({}, {})",
                edit.row, edit.col
            ),
        ))),
        (None, _) => patch_value_cell(writer, col, style, edit, None),
    }
}

fn write_new_fmla_num<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    col: u32,
    style: u32,
    cached: f64,
    flags: u16,
    rgce: &[u8],
    rgcb: &[u8],
) -> Result<(), Error> {
    let rgce_len = u32::try_from(rgce.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula token stream is too large",
        ))
    })?;
    let rgcb_len = u32::try_from(rgcb.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula trailing payload is too large",
        ))
    })?;
    let payload_len = 22u32
        .checked_add(rgce_len)
        .and_then(|v| v.checked_add(rgcb_len))
        .ok_or(Error::UnexpectedEof)?;
    writer.write_record_header(biff12::FORMULA_FLOAT, payload_len)?;
    writer.write_u32(col)?;
    writer.write_u32(style)?;
    writer.write_f64(cached)?;
    writer.write_u16(flags)?;
    writer.write_u32(rgce_len)?;
    writer.write_raw(rgce)?;
    writer.write_raw(rgcb)?;
    Ok(())
}

fn write_new_fmla_bool<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    col: u32,
    style: u32,
    cached: bool,
    flags: u16,
    rgce: &[u8],
    rgcb: &[u8],
) -> Result<(), Error> {
    let rgce_len = u32::try_from(rgce.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula token stream is too large",
        ))
    })?;
    let rgcb_len = u32::try_from(rgcb.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula trailing payload is too large",
        ))
    })?;
    let payload_len = 15u32
        .checked_add(rgce_len)
        .and_then(|v| v.checked_add(rgcb_len))
        .ok_or(Error::UnexpectedEof)?;
    writer.write_record_header(biff12::FORMULA_BOOL, payload_len)?;
    writer.write_u32(col)?;
    writer.write_u32(style)?;
    writer.write_raw(&[u8::from(cached)])?;
    writer.write_u16(flags)?;
    writer.write_u32(rgce_len)?;
    writer.write_raw(rgce)?;
    writer.write_raw(rgcb)?;
    Ok(())
}

fn write_new_fmla_error<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    col: u32,
    style: u32,
    cached: u8,
    flags: u16,
    rgce: &[u8],
    rgcb: &[u8],
) -> Result<(), Error> {
    let rgce_len = u32::try_from(rgce.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula token stream is too large",
        ))
    })?;
    let rgcb_len = u32::try_from(rgcb.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula trailing payload is too large",
        ))
    })?;
    let payload_len = 15u32
        .checked_add(rgce_len)
        .and_then(|v| v.checked_add(rgcb_len))
        .ok_or(Error::UnexpectedEof)?;
    writer.write_record_header(biff12::FORMULA_BOOLERR, payload_len)?;
    writer.write_u32(col)?;
    writer.write_u32(style)?;
    writer.write_raw(&[cached])?;
    writer.write_u16(flags)?;
    writer.write_u32(rgce_len)?;
    writer.write_raw(rgce)?;
    writer.write_raw(rgcb)?;
    Ok(())
}

fn write_new_fmla_string<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    col: u32,
    style: u32,
    cached: &str,
    flags: u16,
    rgce: &[u8],
    rgcb: &[u8],
) -> Result<(), Error> {
    let rgce_len = u32::try_from(rgce.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula token stream is too large",
        ))
    })?;
    let rgcb_len = u32::try_from(rgcb.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula trailing payload is too large",
        ))
    })?;

    let cch = cached.encode_utf16().count();
    let cch = u32::try_from(cch).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "string is too large",
        ))
    })?;
    let utf16_len = cch.checked_mul(2).ok_or(Error::UnexpectedEof)?;
    // [cch:u32][flags:u16][utf16 chars...][optional rich/phonetic headers...]
    let mut cached_len = 6u32.checked_add(utf16_len).ok_or(Error::UnexpectedEof)?;
    if flags & FLAG_RICH != 0 {
        cached_len = cached_len.checked_add(4).ok_or(Error::UnexpectedEof)?;
    }
    if flags & FLAG_PHONETIC != 0 {
        cached_len = cached_len.checked_add(4).ok_or(Error::UnexpectedEof)?;
    }
    let payload_len = 8u32
        .checked_add(cached_len)
        .and_then(|v| v.checked_add(4)) // cce
        .and_then(|v| v.checked_add(rgce_len))
        .and_then(|v| v.checked_add(rgcb_len))
        .ok_or(Error::UnexpectedEof)?;

    writer.write_record_header(biff12::FORMULA_STRING, payload_len)?;
    writer.write_u32(col)?;
    writer.write_u32(style)?;
    writer.write_u32(cch)?;
    writer.write_u16(flags)?;
    for unit in cached.encode_utf16() {
        writer.write_raw(&unit.to_le_bytes())?;
    }
    if flags & FLAG_RICH != 0 {
        writer.write_u32(0)?; // cRun
    }
    if flags & FLAG_PHONETIC != 0 {
        writer.write_u32(0)?; // cb
    }
    writer.write_u32(rgce_len)?;
    writer.write_raw(rgce)?;
    writer.write_raw(rgcb)?;
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct WideStringOffsets {
    cch: usize,
    flags: u16,
    utf16_start: usize,
    utf16_end: usize,
    end: usize,
}

fn parse_wide_string_offsets(
    data: &[u8],
    start: usize,
    flags_width: FlagsWidth,
) -> Result<WideStringOffsets, Error> {
    let cch = read_u32(data, start)? as usize;
    let mut offset = start.checked_add(4).ok_or(Error::UnexpectedEof)?;
    let flags = match flags_width {
        FlagsWidth::U8 => {
            let v = read_u8(data, offset)? as u16;
            offset = offset.checked_add(1).ok_or(Error::UnexpectedEof)?;
            v
        }
        FlagsWidth::U16 => {
            let v = read_u16(data, offset)?;
            offset = offset.checked_add(2).ok_or(Error::UnexpectedEof)?;
            v
        }
    };

    let utf16_start = offset;
    let utf16_len = cch.checked_mul(2).ok_or(Error::UnexpectedEof)?;
    let utf16_end = utf16_start
        .checked_add(utf16_len)
        .ok_or(Error::UnexpectedEof)?;
    data.get(utf16_start..utf16_end)
        .ok_or(Error::UnexpectedEof)?;
    offset = utf16_end;

    if flags & FLAG_RICH != 0 {
        let c_run = read_u32(data, offset)? as usize;
        offset = offset.checked_add(4).ok_or(Error::UnexpectedEof)?;
        let run_bytes = c_run
            .checked_mul(RICH_RUN_BYTE_LEN)
            .ok_or(Error::UnexpectedEof)?;
        let end = offset.checked_add(run_bytes).ok_or(Error::UnexpectedEof)?;
        data.get(offset..end).ok_or(Error::UnexpectedEof)?;
        offset = end;
    }

    if flags & FLAG_PHONETIC != 0 {
        let cb = read_u32(data, offset)? as usize;
        offset = offset.checked_add(4).ok_or(Error::UnexpectedEof)?;
        let end = offset.checked_add(cb).ok_or(Error::UnexpectedEof)?;
        data.get(offset..end).ok_or(Error::UnexpectedEof)?;
        offset = end;
    }

    Ok(WideStringOffsets {
        cch,
        flags,
        utf16_start,
        utf16_end,
        end: offset,
    })
}

fn reject_formula_payload_edit(edit: &CellEdit, row: u32, col: u32) -> Result<(), Error> {
    if edit.new_formula.is_none() && edit.new_rgcb.is_none() {
        return Ok(());
    }

    let kind = match (edit.new_formula.is_some(), edit.new_rgcb.is_some()) {
        (true, true) => "formula (rgce + rgcb)",
        (true, false) => "formula",
        (false, true) => "formula rgcb",
        (false, false) => unreachable!(),
    };

    Err(Error::Io(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("attempted to set {kind} for non-formula cell at ({row}, {col})"),
    )))
}

fn value_record_expected_payload_len(record_id: u32, payload: &[u8]) -> Result<usize, Error> {
    match record_id {
        biff12::BLANK => Ok(8),
        biff12::BOOL | biff12::BOOLERR => Ok(9),
        biff12::NUM | biff12::STRING => Ok(12),
        biff12::FLOAT => Ok(16),
        biff12::CELL_ST => {
            if payload.len() < 12 {
                return Err(Error::UnexpectedEof);
            }

            // BrtCellSt has two observed layouts:
            // - Simple: [col][style][cch][utf16 chars...]
            // - Flagged: [col][style][cch][flags:u8][utf16 chars...][optional rich/phonetic...]
            let cch = read_u32(payload, 8)? as usize;
            let utf16_len = cch.checked_mul(2).ok_or(Error::UnexpectedEof)?;
            let expected_simple = 12usize.checked_add(utf16_len).ok_or(Error::UnexpectedEof)?;
            if payload.len() == expected_simple {
                return Ok(expected_simple);
            }

            let ws = parse_wide_string_offsets(payload, 8, FlagsWidth::U8)?;
            Ok(ws.end)
        }
        _ => Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!("unsupported value record type 0x{record_id:04X}"),
        ))),
    }
}

fn convert_value_record_to_formula<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    record_id: u32,
    payload: &[u8],
    col: u32,
    style: u32,
    edit: &CellEdit,
) -> Result<(), Error> {
    let Some(rgce) = edit.new_formula.as_deref() else {
        return Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "convert_value_record_to_formula called without edit.new_formula",
        )));
    };

    // Value-record payloads are fully specified by the MS-XLSB spec. If we encounter unexpected
    // trailing bytes, do not drop them silently when converting the record type — require the
    // caller to explicitly provide `new_rgcb` to confirm the replacement.
    let expected_len = value_record_expected_payload_len(record_id, payload)?;
    if payload.len() < expected_len {
        return Err(Error::UnexpectedEof);
    }
    if payload.len() > expected_len && edit.new_rgcb.is_none() {
        return Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "cannot convert value cell at ({}, {}) to formula: existing record 0x{record_id:04X} has unexpected trailing bytes; provide CellEdit.new_rgcb (even empty) to replace them",
                edit.row, edit.col
            ),
        )));
    }

    let rgcb = edit.new_rgcb.as_deref().unwrap_or(&[]);
    if rgcb.is_empty() && rgce_references_rgcb(rgce) {
        return Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "formula update for cell at ({}, {}) requires rgcb bytes (PtgArray present); set CellEdit.new_rgcb",
                edit.row, edit.col
            ),
        )));
    }
    // Value records do not carry `BrtFmla*` flags. When converting a value record into a formula
    // record, default to `0` unless the caller explicitly overrides. This mirrors Excel's behavior
    // for newly-created formula cells.
    let flags = edit.new_formula_flags.unwrap_or(0);
    match &edit.new_value {
        CellValue::Number(v) => write_new_fmla_num(writer, col, style, *v, flags, rgce, rgcb),
        CellValue::Text(s) => write_new_fmla_string(writer, col, style, s, flags, rgce, rgcb),
        CellValue::Bool(v) => write_new_fmla_bool(writer, col, style, *v, flags, rgce, rgcb),
        CellValue::Error(v) => write_new_fmla_error(writer, col, style, *v, flags, rgce, rgcb),
        CellValue::Blank => Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "cannot write formula cell with blank cached value at ({}, {})",
                edit.row, edit.col
            ),
        ))),
    }
}

fn formula_edit_is_noop(payload: &[u8], edit: &CellEdit) -> Result<bool, Error> {
    if edit.clear_formula {
        return Ok(false);
    }
    if let Some(desired_style) = edit.new_style {
        let existing_style = read_u32(payload, 4)?;
        if existing_style != desired_style {
            return Ok(false);
        }
    }
    let existing_cached = read_f64(payload, 8)?;
    let existing_flags = read_u16(payload, 16)?;
    let desired_flags = edit.new_formula_flags.unwrap_or(existing_flags);
    let cce = read_u32(payload, 18)? as usize;
    let rgce_offset = 22usize;
    let rgce_end = rgce_offset.checked_add(cce).ok_or(Error::UnexpectedEof)?;
    let rgce = payload
        .get(rgce_offset..rgce_end)
        .ok_or(Error::UnexpectedEof)?;
    let extra = payload.get(rgce_end..).unwrap_or(&[]);

    let desired_cached = match &edit.new_value {
        CellValue::Number(v) => *v,
        _ => return Ok(false),
    };

    let desired_rgce: &[u8] = edit.new_formula.as_deref().unwrap_or(rgce);
    let desired_extra: &[u8] = edit.new_rgcb.as_deref().unwrap_or(extra);

    Ok(existing_cached.to_bits() == desired_cached.to_bits()
        && existing_flags == desired_flags
        && rgce == desired_rgce
        && extra == desired_extra)
}

fn formula_string_edit_is_noop(payload: &[u8], edit: &CellEdit) -> Result<bool, Error> {
    if edit.clear_formula {
        return Ok(false);
    }
    if let Some(desired_style) = edit.new_style {
        let existing_style = read_u32(payload, 4)?;
        if existing_style != desired_style {
            return Ok(false);
        }
    }
    let desired_cached = match &edit.new_value {
        CellValue::Text(s) => s,
        _ => return Ok(false),
    };

    let ws = parse_fmla_string_cached_value_offsets(payload)?;
    let desired_flags = edit.new_formula_flags.unwrap_or(ws.flags);
    if desired_flags != ws.flags {
        return Ok(false);
    }
    let raw = payload
        .get(ws.utf16_start..ws.utf16_end)
        .ok_or(Error::UnexpectedEof)?;

    if desired_cached.encode_utf16().count() != ws.cch {
        return Ok(false);
    }

    let str_bytes_len = ws.cch.checked_mul(2).ok_or(Error::UnexpectedEof)?;
    let mut desired_bytes = Vec::with_capacity(str_bytes_len);
    for unit in desired_cached.encode_utf16() {
        desired_bytes.extend_from_slice(&unit.to_le_bytes());
    }
    if desired_bytes != raw {
        return Ok(false);
    }

    let cce = read_u32(payload, ws.end)? as usize;
    let rgce_offset = ws.end.checked_add(4).ok_or(Error::UnexpectedEof)?;
    let rgce_end = rgce_offset.checked_add(cce).ok_or(Error::UnexpectedEof)?;
    let rgce = payload
        .get(rgce_offset..rgce_end)
        .ok_or(Error::UnexpectedEof)?;
    let extra = payload.get(rgce_end..).unwrap_or(&[]);

    let desired_rgce: &[u8] = edit.new_formula.as_deref().unwrap_or(rgce);
    let desired_extra: &[u8] = edit.new_rgcb.as_deref().unwrap_or(extra);
    Ok(rgce == desired_rgce && extra == desired_extra)
}

fn formula_bool_edit_is_noop(payload: &[u8], edit: &CellEdit) -> Result<bool, Error> {
    if edit.clear_formula {
        return Ok(false);
    }
    if let Some(desired_style) = edit.new_style {
        let existing_style = read_u32(payload, 4)?;
        if existing_style != desired_style {
            return Ok(false);
        }
    }
    let desired_cached = match &edit.new_value {
        CellValue::Bool(v) => *v,
        _ => return Ok(false),
    };
    let existing_cached = read_u8(payload, 8)? != 0;
    let existing_flags = read_u16(payload, 9)?;
    let desired_flags = edit.new_formula_flags.unwrap_or(existing_flags);
    let cce = read_u32(payload, 11)? as usize;
    let rgce_offset = 15usize;
    let rgce_end = rgce_offset.checked_add(cce).ok_or(Error::UnexpectedEof)?;
    let rgce = payload
        .get(rgce_offset..rgce_end)
        .ok_or(Error::UnexpectedEof)?;
    let extra = payload.get(rgce_end..).unwrap_or(&[]);
    let desired_rgce: &[u8] = edit.new_formula.as_deref().unwrap_or(rgce);
    let desired_extra: &[u8] = edit.new_rgcb.as_deref().unwrap_or(extra);
    Ok(existing_cached == desired_cached
        && existing_flags == desired_flags
        && rgce == desired_rgce
        && extra == desired_extra)
}

fn formula_error_edit_is_noop(payload: &[u8], edit: &CellEdit) -> Result<bool, Error> {
    if edit.clear_formula {
        return Ok(false);
    }
    if let Some(desired_style) = edit.new_style {
        let existing_style = read_u32(payload, 4)?;
        if existing_style != desired_style {
            return Ok(false);
        }
    }
    let desired_cached = match &edit.new_value {
        CellValue::Error(v) => *v,
        _ => return Ok(false),
    };
    let existing_cached = read_u8(payload, 8)?;
    let existing_flags = read_u16(payload, 9)?;
    let desired_flags = edit.new_formula_flags.unwrap_or(existing_flags);
    let cce = read_u32(payload, 11)? as usize;
    let rgce_offset = 15usize;
    let rgce_end = rgce_offset.checked_add(cce).ok_or(Error::UnexpectedEof)?;
    let rgce = payload
        .get(rgce_offset..rgce_end)
        .ok_or(Error::UnexpectedEof)?;
    let extra = payload.get(rgce_end..).unwrap_or(&[]);
    let desired_rgce: &[u8] = edit.new_formula.as_deref().unwrap_or(rgce);
    let desired_extra: &[u8] = edit.new_rgcb.as_deref().unwrap_or(extra);
    Ok(existing_cached == desired_cached
        && existing_flags == desired_flags
        && rgce == desired_rgce
        && extra == desired_extra)
}

fn value_edit_is_noop_float(payload: &[u8], edit: &CellEdit) -> Result<bool, Error> {
    if edit.new_formula.is_some() || edit.new_rgcb.is_some() {
        return Ok(false);
    }
    if let Some(desired_style) = edit.new_style {
        let existing_style = read_u32(payload, 4)?;
        if existing_style != desired_style {
            return Ok(false);
        }
    }
    let existing = read_f64(payload, 8)?;
    let desired = match &edit.new_value {
        CellValue::Number(v) => *v,
        _ => return Ok(false),
    };
    Ok(existing.to_bits() == desired.to_bits())
}

fn value_edit_is_noop_rk(payload: &[u8], edit: &CellEdit) -> Result<bool, Error> {
    if edit.new_formula.is_some() || edit.new_rgcb.is_some() {
        return Ok(false);
    }
    if let Some(desired_style) = edit.new_style {
        let existing_style = read_u32(payload, 4)?;
        if existing_style != desired_style {
            return Ok(false);
        }
    }
    let existing_rk = read_u32(payload, 8)?;
    let desired = match &edit.new_value {
        CellValue::Number(v) => *v,
        _ => return Ok(false),
    };
    let existing = decode_rk_number(existing_rk);
    Ok(existing.to_bits() == desired.to_bits())
}

fn value_edit_is_noop_shared_string(payload: &[u8], edit: &CellEdit) -> Result<bool, Error> {
    if edit.new_formula.is_some() || edit.new_rgcb.is_some() {
        return Ok(false);
    }
    if let Some(desired_style) = edit.new_style {
        let existing_style = read_u32(payload, 4)?;
        if existing_style != desired_style {
            return Ok(false);
        }
    }
    let Some(isst) = edit.shared_string_index else {
        return Ok(false);
    };
    if !matches!(edit.new_value, CellValue::Text(_)) {
        return Ok(false);
    }
    Ok(read_u32(payload, 8)? == isst)
}

pub(crate) fn value_edit_is_noop_inline_string(
    payload: &[u8],
    edit: &CellEdit,
) -> Result<bool, Error> {
    if edit.new_formula.is_some() || edit.new_rgcb.is_some() {
        return Ok(false);
    }
    if let Some(desired_style) = edit.new_style {
        let existing_style = read_u32(payload, 4)?;
        if existing_style != desired_style {
            return Ok(false);
        }
    }
    let desired = match &edit.new_value {
        CellValue::Text(s) => s,
        _ => return Ok(false),
    };

    // BrtCellSt inline strings appear in at least two layouts:
    // - simple:  [col][style][cch:u32][utf16 bytes...]
    // - flagged: [col][style][cch:u32][flags:u8][utf16 bytes...][optional extras...]
    //
    // Some producers set the wide-string flags byte such that it implies rich/phonetic extras,
    // but do not actually emit those blocks. For the purpose of *no-op detection*, be lenient:
    // never hard-fail on inconsistent flags/extras. Instead, compare the raw UTF-16 bytes against
    // the desired string, trying both plausible start offsets and allowing trailing bytes.
    let Ok(len_chars_u32) = read_u32(payload, 8) else {
        // Malformed record: treat as non-noop so the patcher can rewrite it.
        return Ok(false);
    };
    let len_chars = len_chars_u32 as usize;
    let desired_len_chars = desired.encode_utf16().count();
    if desired_len_chars != len_chars {
        return Ok(false);
    }

    let Some(byte_len) = len_chars.checked_mul(2) else {
        return Ok(false);
    };

    // Avoid parsing rich/phonetic blocks; just compare the UTF-16LE bytes directly.
    let mut desired_bytes = Vec::with_capacity(byte_len);
    for unit in desired.encode_utf16() {
        desired_bytes.extend_from_slice(&unit.to_le_bytes());
    }

    // Flagged layout: UTF-16 begins after the flags byte.
    if let Some(end) = 13usize.checked_add(byte_len) {
        if let Some(raw) = payload.get(13..end) {
            if raw == desired_bytes.as_slice() {
                return Ok(true);
            }
        }
    }

    // Simple layout: UTF-16 begins immediately after the `cch` field.
    if let Some(end) = 12usize.checked_add(byte_len) {
        if let Some(raw) = payload.get(12..end) {
            if raw == desired_bytes.as_slice() {
                return Ok(true);
            }
        }
    }

    Ok(false)
}

fn value_edit_is_noop_bool(payload: &[u8], edit: &CellEdit) -> Result<bool, Error> {
    if edit.new_formula.is_some() || edit.new_rgcb.is_some() {
        return Ok(false);
    }
    if let Some(desired_style) = edit.new_style {
        let existing_style = read_u32(payload, 4)?;
        if existing_style != desired_style {
            return Ok(false);
        }
    }
    let desired = match &edit.new_value {
        CellValue::Bool(v) => *v,
        _ => return Ok(false),
    };
    let existing = read_u8(payload, 8)? != 0;
    Ok(existing == desired)
}

fn value_edit_is_noop_error(payload: &[u8], edit: &CellEdit) -> Result<bool, Error> {
    if edit.new_formula.is_some() || edit.new_rgcb.is_some() {
        return Ok(false);
    }
    if let Some(desired_style) = edit.new_style {
        let existing_style = read_u32(payload, 4)?;
        if existing_style != desired_style {
            return Ok(false);
        }
    }
    let desired = match &edit.new_value {
        CellValue::Error(v) => *v,
        _ => return Ok(false),
    };
    Ok(read_u8(payload, 8)? == desired)
}

fn value_edit_is_noop_blank(existing_style: u32, edit: &CellEdit) -> bool {
    let style_unchanged = match edit.new_style {
        Some(desired) => desired == existing_style,
        None => true,
    };
    style_unchanged
        && edit.new_formula.is_none()
        && edit.new_rgcb.is_none()
        && matches!(edit.new_value, CellValue::Blank)
}

fn patch_value_cell<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    col: u32,
    style: u32,
    edit: &CellEdit,
    existing: Option<ExistingRecordHeader<'_>>,
) -> Result<(), Error> {
    reject_formula_payload_edit(edit, edit.row, edit.col)?;
    match &edit.new_value {
        CellValue::Blank => {
            let mut payload = [0u8; 8];
            payload[0..4].copy_from_slice(&col.to_le_bytes());
            payload[4..8].copy_from_slice(&style.to_le_bytes());
            write_record_preserving_varints(writer, biff12::BLANK, &payload, existing)?;
        }
        CellValue::Number(v) => {
            let mut payload = [0u8; 16];
            payload[0..4].copy_from_slice(&col.to_le_bytes());
            payload[4..8].copy_from_slice(&style.to_le_bytes());
            payload[8..16].copy_from_slice(&v.to_le_bytes());
            write_record_preserving_varints(writer, biff12::FLOAT, &payload, existing)?;
        }
        CellValue::Bool(v) => {
            let mut payload = [0u8; 9];
            payload[0..4].copy_from_slice(&col.to_le_bytes());
            payload[4..8].copy_from_slice(&style.to_le_bytes());
            payload[8] = u8::from(*v);
            write_record_preserving_varints(writer, biff12::BOOL, &payload, existing)?;
        }
        CellValue::Error(v) => {
            let mut payload = [0u8; 9];
            payload[0..4].copy_from_slice(&col.to_le_bytes());
            payload[4..8].copy_from_slice(&style.to_le_bytes());
            payload[8] = *v;
            write_record_preserving_varints(writer, biff12::BOOLERR, &payload, existing)?;
        }
        CellValue::Text(s) => {
            if let Some(isst) = edit.shared_string_index {
                // BrtCellIsst: [col: u32][style: u32][isst: u32]
                let mut payload = [0u8; 12];
                payload[0..4].copy_from_slice(&col.to_le_bytes());
                payload[4..8].copy_from_slice(&style.to_le_bytes());
                payload[8..12].copy_from_slice(&isst.to_le_bytes());
                write_record_preserving_varints(writer, biff12::STRING, &payload, existing)?;
                return Ok(());
            }

            let char_len = s.encode_utf16().count();
            let char_len = u32::try_from(char_len).map_err(|_| {
                Error::Io(io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "string is too large",
                ))
            })?;
            let bytes_len = char_len.checked_mul(2).ok_or(Error::UnexpectedEof)?;
            let payload_len = 12u32.checked_add(bytes_len).ok_or(Error::UnexpectedEof)?;

            write_record_header_preserving_varints(writer, biff12::CELL_ST, payload_len, existing)?;
            writer.write_u32(col)?;
            writer.write_u32(style)?;
            writer.write_utf16_string(s)?;
        }
    }
    Ok(())
}

fn patch_fixed_value_cell_preserving_trailing_bytes<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    record_id: u32,
    payload: &[u8],
    col: u32,
    style: u32,
    edit: &CellEdit,
    existing: Option<ExistingRecordHeader<'_>>,
) -> Result<(), Error> {
    // Some fixed-size value records are fully specified by the MS-XLSB spec, but in practice we
    // may encounter malformed streams or future-compatible extensions that append unexpected
    // trailing bytes. When patching *within the same record type*, preserve those trailing bytes
    // by rewriting the payload in-place instead of truncating to the spec-defined length.
    reject_formula_payload_edit(edit, edit.row, edit.col)?;

    match (record_id, &edit.new_value) {
        (biff12::FLOAT, CellValue::Number(v)) => {
            // Fast path: the spec-defined payload is 16 bytes.
            if payload.len() == 16 {
                return patch_value_cell(writer, col, style, edit, existing);
            }
            if payload.len() < 16 {
                return Err(Error::UnexpectedEof);
            }
            let mut patched = payload.to_vec();
            patched[0..4].copy_from_slice(&col.to_le_bytes());
            patched[4..8].copy_from_slice(&style.to_le_bytes());
            patched[8..16].copy_from_slice(&v.to_le_bytes());
            write_record_preserving_varints(writer, biff12::FLOAT, &patched, existing)?;
            Ok(())
        }
        (biff12::BOOL, CellValue::Bool(v)) => {
            // Fast path: the spec-defined payload is 9 bytes.
            if payload.len() == 9 {
                return patch_value_cell(writer, col, style, edit, existing);
            }
            if payload.len() < 9 {
                return Err(Error::UnexpectedEof);
            }
            let mut patched = payload.to_vec();
            patched[0..4].copy_from_slice(&col.to_le_bytes());
            patched[4..8].copy_from_slice(&style.to_le_bytes());
            patched[8] = u8::from(*v);
            write_record_preserving_varints(writer, biff12::BOOL, &patched, existing)?;
            Ok(())
        }
        (biff12::BOOLERR, CellValue::Error(v)) => {
            // Fast path: the spec-defined payload is 9 bytes.
            if payload.len() == 9 {
                return patch_value_cell(writer, col, style, edit, existing);
            }
            if payload.len() < 9 {
                return Err(Error::UnexpectedEof);
            }
            let mut patched = payload.to_vec();
            patched[0..4].copy_from_slice(&col.to_le_bytes());
            patched[4..8].copy_from_slice(&style.to_le_bytes());
            patched[8] = *v;
            write_record_preserving_varints(writer, biff12::BOOLERR, &patched, existing)?;
            Ok(())
        }
        (biff12::BLANK, CellValue::Blank) => {
            // Fast path: the spec-defined payload is 8 bytes.
            if payload.len() == 8 {
                return patch_value_cell(writer, col, style, edit, existing);
            }
            if payload.len() < 8 {
                return Err(Error::UnexpectedEof);
            }
            let mut patched = payload.to_vec();
            patched[0..4].copy_from_slice(&col.to_le_bytes());
            patched[4..8].copy_from_slice(&style.to_le_bytes());
            write_record_preserving_varints(writer, biff12::BLANK, &patched, existing)?;
            Ok(())
        }
        _ => patch_value_cell(writer, col, style, edit, existing),
    }
}

fn patch_cell_st<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    payload: &[u8],
    col: u32,
    style: u32,
    edit: &CellEdit,
    existing: Option<ExistingRecordHeader<'_>>,
) -> Result<(), Error> {
    // `BrtCellSt` stores inline strings. Most producers use the standard "wide string" layout:
    //
    //   [cch:u32][flags:u8][utf16 chars...][optional rich/phonetic blocks]
    //
    // but some emit a simplified form without the flags byte when no extras are present:
    //
    //   [cch:u32][utf16 chars...]
    //
    // When patching an existing inline string cell, preserve the original layout class so
    // round-trip diffs stay minimal.
    let CellValue::Text(text) = &edit.new_value else {
        return patch_value_cell(writer, col, style, edit, existing);
    };
    if edit.shared_string_index.is_some() {
        // Caller explicitly requested a shared string reference; defer to the generic writer.
        return patch_value_cell(writer, col, style, edit, existing);
    }

    // Detect whether the original record used the "simple" inline string encoding. The
    // simplified form is byte-for-byte:
    //   [col:u32][style:u32][cch:u32][utf16 chars...]
    let existing_cch = read_u32(payload, 8)? as usize;
    let existing_utf16_len = existing_cch.checked_mul(2).ok_or(Error::UnexpectedEof)?;
    let expected_simple_len = 12usize
        .checked_add(existing_utf16_len)
        .ok_or(Error::UnexpectedEof)?;

    if payload.len() == expected_simple_len {
        return patch_value_cell(writer, col, style, edit, existing);
    }

    // Flagged wide-string layout: parse to determine the original flags and whether rich /
    // phonetic blocks were present.
    let ws = parse_wide_string_offsets(payload, 8, FlagsWidth::U8)?;
    let flags_u8 = ws.flags as u8;

    let desired_cch = text.encode_utf16().count();
    let desired_cch_u32 = u32::try_from(desired_cch).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "string is too large",
        ))
    })?;
    let desired_utf16_len = desired_cch_u32
        .checked_mul(2)
        .ok_or(Error::UnexpectedEof)?;

    // If the string contents are unchanged (e.g. style-only edits), preserve the original
    // wide-string payload bytes (including any rich/phonetic blocks) so we don't accidentally
    // strip formatting metadata.
    let existing_utf16 = payload
        .get(ws.utf16_start..ws.utf16_end)
        .ok_or(Error::UnexpectedEof)?;
    let mut desired_utf16 = Vec::with_capacity(desired_utf16_len as usize);
    for unit in text.encode_utf16() {
        desired_utf16.extend_from_slice(&unit.to_le_bytes());
    }
    let preserve_wide_string_bytes = desired_cch == ws.cch && desired_utf16.as_slice() == existing_utf16;
    if preserve_wide_string_bytes {
        // Preserve the full payload bytes (including any unknown trailing bytes) and patch style
        // in-place. This keeps diffs minimal and retains rich/phonetic payloads.
        let mut patched = payload.to_vec();
        patched[0..4].copy_from_slice(&col.to_le_bytes());
        patched[4..8].copy_from_slice(&style.to_le_bytes());
        write_record_preserving_varints(writer, biff12::CELL_ST, &patched, existing)?;
        return Ok(());
    }

    // [col:u32][style:u32] + [cch:u32][flags:u8][utf16...] + optional empty rich/phonetic blocks.
    let mut payload_len = 13u32
        .checked_add(desired_utf16_len)
        .ok_or(Error::UnexpectedEof)?;
    if ws.flags & FLAG_RICH != 0 {
        payload_len = payload_len.checked_add(4).ok_or(Error::UnexpectedEof)?;
    }
    if ws.flags & FLAG_PHONETIC != 0 {
        payload_len = payload_len.checked_add(4).ok_or(Error::UnexpectedEof)?;
    }

    write_record_header_preserving_varints(writer, biff12::CELL_ST, payload_len, existing)?;
    writer.write_u32(col)?;
    writer.write_u32(style)?;
    writer.write_u32(desired_cch_u32)?;
    writer.write_raw(&[flags_u8])?;
    writer.write_raw(&desired_utf16)?;
    if ws.flags & FLAG_RICH != 0 {
        writer.write_u32(0)?; // cRun
    }
    if ws.flags & FLAG_PHONETIC != 0 {
        writer.write_u32(0)?; // cb
    }
    Ok(())
}

fn patch_rk_cell<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    col: u32,
    style: u32,
    payload: &[u8],
    edit: &CellEdit,
    existing: Option<ExistingRecordHeader<'_>>,
) -> Result<(), Error> {
    reject_formula_payload_edit(edit, edit.row, edit.col)?;
    match &edit.new_value {
        CellValue::Number(v) => {
            if let Some(rk) = encode_rk_number(*v) {
                if payload.len() < 12 {
                    return Err(Error::UnexpectedEof);
                }
                if payload.len() == 12 {
                    let mut payload = [0u8; 12];
                    payload[0..4].copy_from_slice(&col.to_le_bytes());
                    payload[4..8].copy_from_slice(&style.to_le_bytes());
                    payload[8..12].copy_from_slice(&rk.to_le_bytes());
                    write_record_preserving_varints(writer, biff12::NUM, &payload, existing)?;
                } else {
                    let mut patched = payload.to_vec();
                    patched[0..4].copy_from_slice(&col.to_le_bytes());
                    patched[4..8].copy_from_slice(&style.to_le_bytes());
                    patched[8..12].copy_from_slice(&rk.to_le_bytes());
                    write_record_preserving_varints(writer, biff12::NUM, &patched, existing)?;
                }
                return Ok(());
            }
        }
        _ => {}
    }

    // Fall back to the generic (FLOAT / inline string) writer.
    patch_value_cell(writer, col, style, edit, existing)
}

fn patch_fmla_num<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    payload: &[u8],
    col: u32,
    style: u32,
    edit: &CellEdit,
    existing: Option<ExistingRecordHeader<'_>>,
) -> Result<(), Error> {
    if edit.clear_formula {
        // Convert an existing formula cell into a plain value cell ("paste values").
        return patch_value_cell(writer, col, style, edit, existing);
    }
    if matches!(edit.new_value, CellValue::Blank) && edit.new_formula.is_none() {
        // Allow clearing formula cells by rewriting the record as a plain blank cell while
        // preserving the original `style` index.
        return patch_value_cell(writer, col, style, edit, existing);
    }

    // BrtFmlaNum: [col: u32][style: u32][value: f64][flags: u16][cce: u32][rgce bytes...]
    let existing_flags = read_u16(payload, 16)?;
    let flags = edit.new_formula_flags.unwrap_or(existing_flags);
    let cce = read_u32(payload, 18)? as usize;
    let rgce_offset = 22usize;
    let rgce_end = rgce_offset.checked_add(cce).ok_or(Error::UnexpectedEof)?;
    let rgce = payload
        .get(rgce_offset..rgce_end)
        .ok_or(Error::UnexpectedEof)?;
    let extra = payload.get(rgce_end..).unwrap_or(&[]);

    let new_rgce: &[u8] = edit.new_formula.as_deref().unwrap_or(rgce);
    let new_extra: &[u8] = edit.new_rgcb.as_deref().unwrap_or(extra);
    if edit.new_formula.is_some()
        && new_rgce != rgce
        && !extra.is_empty()
        && edit.new_rgcb.is_none()
    {
        return Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "cannot replace formula rgce for BrtFmlaNum at ({}, {}) with existing trailing rgcb bytes; provide CellEdit.new_rgcb (even empty) to replace them",
                edit.row, edit.col
            ),
        )));
    }
    if new_extra.is_empty() && rgce_references_rgcb(new_rgce) {
        return Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "formula update for BrtFmlaNum at ({}, {}) requires rgcb bytes (PtgArray present); set CellEdit.new_rgcb",
                edit.row, edit.col
            ),
        )));
    }
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

    let new_rgce_len = u32::try_from(new_rgce.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula token stream is too large",
        ))
    })?;
    let extra_len = u32::try_from(new_extra.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula trailing payload is too large",
        ))
    })?;
    let payload_len = 22u32
        .checked_add(new_rgce_len)
        .and_then(|v| v.checked_add(extra_len))
        .ok_or(Error::UnexpectedEof)?;

    write_record_header_preserving_varints(writer, biff12::FORMULA_FLOAT, payload_len, existing)?;
    writer.write_u32(col)?;
    writer.write_u32(style)?;
    writer.write_f64(cached)?;
    writer.write_u16(flags)?;
    writer.write_u32(new_rgce_len)?;
    writer.write_raw(new_rgce)?;
    writer.write_raw(new_extra)?;
    Ok(())
}

fn patch_fmla_bool<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    payload: &[u8],
    col: u32,
    style: u32,
    edit: &CellEdit,
    existing: Option<ExistingRecordHeader<'_>>,
) -> Result<(), Error> {
    if edit.clear_formula {
        return patch_value_cell(writer, col, style, edit, existing);
    }
    if matches!(edit.new_value, CellValue::Blank) && edit.new_formula.is_none() {
        return patch_value_cell(writer, col, style, edit, existing);
    }

    // BrtFmlaBool: [col: u32][style: u32][value: u8][flags: u16][cce: u32][rgce bytes...][extra...]
    let existing_flags = read_u16(payload, 9)?;
    let flags = edit.new_formula_flags.unwrap_or(existing_flags);
    let cce = read_u32(payload, 11)? as usize;
    let rgce_offset = 15usize;
    let rgce_end = rgce_offset.checked_add(cce).ok_or(Error::UnexpectedEof)?;
    let rgce = payload
        .get(rgce_offset..rgce_end)
        .ok_or(Error::UnexpectedEof)?;
    let extra = payload.get(rgce_end..).unwrap_or(&[]);

    let new_rgce: &[u8] = edit.new_formula.as_deref().unwrap_or(rgce);
    let new_extra: &[u8] = edit.new_rgcb.as_deref().unwrap_or(extra);
    if edit.new_formula.is_some()
        && new_rgce != rgce
        && !extra.is_empty()
        && edit.new_rgcb.is_none()
    {
        return Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "cannot replace formula rgce for BrtFmlaBool at ({}, {}) with existing trailing rgcb bytes; provide CellEdit.new_rgcb (even empty) to replace them",
                edit.row, edit.col
            ),
        )));
    }
    if new_extra.is_empty() && rgce_references_rgcb(new_rgce) {
        return Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "formula update for BrtFmlaBool at ({}, {}) requires rgcb bytes (PtgArray present); set CellEdit.new_rgcb",
                edit.row, edit.col
            ),
        )));
    }

    let cached = match &edit.new_value {
        CellValue::Bool(v) => *v,
        _ => {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!(
                    "BrtFmlaBool edit requires boolean cached value at ({}, {})",
                    edit.row, edit.col
                ),
            )));
        }
    };

    let new_rgce_len = u32::try_from(new_rgce.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula token stream is too large",
        ))
    })?;
    let extra_len = u32::try_from(new_extra.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula trailing payload is too large",
        ))
    })?;
    let payload_len = 15u32
        .checked_add(new_rgce_len)
        .and_then(|v| v.checked_add(extra_len))
        .ok_or(Error::UnexpectedEof)?;

    write_record_header_preserving_varints(writer, biff12::FORMULA_BOOL, payload_len, existing)?;
    writer.write_u32(col)?;
    writer.write_u32(style)?;
    writer.write_raw(&[u8::from(cached)])?;
    writer.write_u16(flags)?;
    writer.write_u32(new_rgce_len)?;
    writer.write_raw(new_rgce)?;
    writer.write_raw(new_extra)?;
    Ok(())
}

fn patch_fmla_error<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    payload: &[u8],
    col: u32,
    style: u32,
    edit: &CellEdit,
    existing: Option<ExistingRecordHeader<'_>>,
) -> Result<(), Error> {
    if edit.clear_formula {
        return patch_value_cell(writer, col, style, edit, existing);
    }
    if matches!(edit.new_value, CellValue::Blank) && edit.new_formula.is_none() {
        return patch_value_cell(writer, col, style, edit, existing);
    }

    // BrtFmlaError: [col: u32][style: u32][value: u8][flags: u16][cce: u32][rgce bytes...][extra...]
    let existing_flags = read_u16(payload, 9)?;
    let flags = edit.new_formula_flags.unwrap_or(existing_flags);
    let cce = read_u32(payload, 11)? as usize;
    let rgce_offset = 15usize;
    let rgce_end = rgce_offset.checked_add(cce).ok_or(Error::UnexpectedEof)?;
    let rgce = payload
        .get(rgce_offset..rgce_end)
        .ok_or(Error::UnexpectedEof)?;
    let extra = payload.get(rgce_end..).unwrap_or(&[]);

    let new_rgce: &[u8] = edit.new_formula.as_deref().unwrap_or(rgce);
    let new_extra: &[u8] = edit.new_rgcb.as_deref().unwrap_or(extra);
    if edit.new_formula.is_some()
        && new_rgce != rgce
        && !extra.is_empty()
        && edit.new_rgcb.is_none()
    {
        return Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "cannot replace formula rgce for BrtFmlaError at ({}, {}) with existing trailing rgcb bytes; provide CellEdit.new_rgcb (even empty) to replace them",
                edit.row, edit.col
            ),
        )));
    }
    if new_extra.is_empty() && rgce_references_rgcb(new_rgce) {
        return Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "formula update for BrtFmlaError at ({}, {}) requires rgcb bytes (PtgArray present); set CellEdit.new_rgcb",
                edit.row, edit.col
            ),
        )));
    }

    let cached = match &edit.new_value {
        CellValue::Error(v) => *v,
        _ => {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!(
                    "BrtFmlaError edit requires error cached value at ({}, {})",
                    edit.row, edit.col
                ),
            )));
        }
    };

    let new_rgce_len = u32::try_from(new_rgce.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula token stream is too large",
        ))
    })?;
    let extra_len = u32::try_from(new_extra.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula trailing payload is too large",
        ))
    })?;
    let payload_len = 15u32
        .checked_add(new_rgce_len)
        .and_then(|v| v.checked_add(extra_len))
        .ok_or(Error::UnexpectedEof)?;

    write_record_header_preserving_varints(writer, biff12::FORMULA_BOOLERR, payload_len, existing)?;
    writer.write_u32(col)?;
    writer.write_u32(style)?;
    writer.write_raw(&[cached])?;
    writer.write_u16(flags)?;
    writer.write_u32(new_rgce_len)?;
    writer.write_raw(new_rgce)?;
    writer.write_raw(new_extra)?;
    Ok(())
}

fn patch_fmla_string<W: io::Write>(
    writer: &mut Biff12Writer<W>,
    payload: &[u8],
    col: u32,
    style: u32,
    edit: &CellEdit,
    existing: Option<ExistingRecordHeader<'_>>,
) -> Result<(), Error> {
    if edit.clear_formula {
        return patch_value_cell(writer, col, style, edit, existing);
    }
    if matches!(edit.new_value, CellValue::Blank) && edit.new_formula.is_none() {
        return patch_value_cell(writer, col, style, edit, existing);
    }

    // BrtFmlaString:
    //   [col: u32][style: u32]
    //   [cached value: u32 cch + u16 flags + utf16 chars...]
    //   [cce: u32][rgce bytes...][extra...]
    let ws = parse_fmla_string_cached_value_offsets(payload)?;
    let existing_flags = ws.flags;
    let flags = edit.new_formula_flags.unwrap_or(existing_flags);
    let layout_changed = ((existing_flags ^ flags) & (FLAG_RICH | FLAG_PHONETIC)) != 0;

    let cce = read_u32(payload, ws.end)? as usize;
    let rgce_offset = ws.end.checked_add(4).ok_or(Error::UnexpectedEof)?;
    let rgce_end = rgce_offset.checked_add(cce).ok_or(Error::UnexpectedEof)?;
    let rgce = payload
        .get(rgce_offset..rgce_end)
        .ok_or(Error::UnexpectedEof)?;
    let extra = payload.get(rgce_end..).unwrap_or(&[]);

    let new_rgce: &[u8] = edit.new_formula.as_deref().unwrap_or(rgce);
    let new_extra: &[u8] = edit.new_rgcb.as_deref().unwrap_or(extra);
    if edit.new_formula.is_some()
        && new_rgce != rgce
        && !extra.is_empty()
        && edit.new_rgcb.is_none()
    {
        return Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "cannot replace formula rgce for BrtFmlaString at ({}, {}) with existing trailing rgcb bytes; provide CellEdit.new_rgcb (even empty) to replace them",
                edit.row, edit.col
            ),
        )));
    }
    if new_extra.is_empty() && rgce_references_rgcb(new_rgce) {
        return Err(Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            format!(
                "formula update for BrtFmlaString at ({}, {}) requires rgcb bytes (PtgArray present); set CellEdit.new_rgcb",
                edit.row, edit.col
            ),
        )));
    }

    let cached = match &edit.new_value {
        CellValue::Text(s) => s,
        _ => {
            return Err(Error::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!(
                    "BrtFmlaString edit requires text cached value at ({}, {})",
                    edit.row, edit.col
                ),
            )));
        }
    };

    let desired_cch = cached.encode_utf16().count();
    let desired_cch_u32 = u32::try_from(desired_cch).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "string is too large",
        ))
    })?;
    let desired_str_len = desired_cch_u32.checked_mul(2).ok_or(Error::UnexpectedEof)?;

    let existing_utf16 = payload
        .get(ws.utf16_start..ws.utf16_end)
        .ok_or(Error::UnexpectedEof)?;
    let mut desired_utf16 = Vec::with_capacity(desired_str_len as usize);
    for unit in cached.encode_utf16() {
        desired_utf16.extend_from_slice(&unit.to_le_bytes());
    }
    let preserve_cached_bytes =
        desired_cch == ws.cch && desired_utf16.as_slice() == existing_utf16 && !layout_changed;
    let cached_bytes_len = if preserve_cached_bytes {
        u32::try_from(ws.end.saturating_sub(8)).map_err(|_| {
            Error::Io(io::Error::new(
                io::ErrorKind::InvalidInput,
                "cached string payload is too large",
            ))
        })?
    } else {
        // [cch:u32][flags:u16][utf16 bytes...][optional rich/phonetic headers...]
        //
        // If the flags indicate rich-text runs or phonetic blocks, emit a minimal empty payload so
        // the stream stays parseable even when we dropped the original formatting bytes.
        let mut len = 6u32
            .checked_add(desired_str_len)
            .ok_or(Error::UnexpectedEof)?;
        if flags & FLAG_RICH != 0 {
            len = len.checked_add(4).ok_or(Error::UnexpectedEof)?;
        }
        if flags & FLAG_PHONETIC != 0 {
            len = len.checked_add(4).ok_or(Error::UnexpectedEof)?;
        }
        len
    };

    let new_rgce_len = u32::try_from(new_rgce.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula token stream is too large",
        ))
    })?;
    let extra_len = u32::try_from(new_extra.len()).map_err(|_| {
        Error::Io(io::Error::new(
            io::ErrorKind::InvalidInput,
            "formula trailing payload is too large",
        ))
    })?;
    let payload_len = 8u32
        .checked_add(cached_bytes_len)
        .and_then(|v| v.checked_add(4)) // cce
        .and_then(|v| v.checked_add(new_rgce_len))
        .and_then(|v| v.checked_add(extra_len))
        .ok_or(Error::UnexpectedEof)?;

    write_record_header_preserving_varints(writer, biff12::FORMULA_STRING, payload_len, existing)?;
    writer.write_u32(col)?;
    writer.write_u32(style)?;
    if preserve_cached_bytes {
        // Preserve the cached string payload bytes, but allow callers to override the u16 flags
        // field in-place.
        writer.write_raw(payload.get(8..12).ok_or(Error::UnexpectedEof)?)?;
        writer.write_u16(flags)?;
        writer.write_raw(
            payload
                .get(ws.utf16_start..ws.end)
                .ok_or(Error::UnexpectedEof)?,
        )?;
    } else {
        writer.write_u32(desired_cch_u32)?;
        writer.write_u16(flags)?;
        writer.write_raw(&desired_utf16)?;
        if flags & FLAG_RICH != 0 {
            writer.write_u32(0)?; // cRun
        }
        if flags & FLAG_PHONETIC != 0 {
            writer.write_u32(0)?; // cb
        }
    }
    writer.write_u32(new_rgce_len)?;
    writer.write_raw(new_rgce)?;
    writer.write_raw(new_extra)?;
    Ok(())
}

fn parse_fmla_string_cached_value_offsets(payload: &[u8]) -> Result<WideStringOffsets, Error> {
    // BrtFmlaString cached values are encoded using `XLWideString` with u16 flags, but some
    // producers appear to set reserved bits that overlap with the rich/phonetic indicators
    // without actually emitting those payload blocks. Be lenient and fall back to the "simple"
    // wide-string layout (no rich/phonetic blocks) when interpreting those bits would shift the
    // subsequent formula fields (`cce`, `rgce`, trailing bytes) out of bounds.

    // Simple layout: [cch:u32][flags:u16][utf16 chars...]
    let cch = read_u32(payload, 8)? as usize;
    let flags = read_u16(payload, 12)?;
    let utf16_start = 14usize;
    let utf16_len = cch.checked_mul(2).ok_or(Error::UnexpectedEof)?;
    let utf16_end = utf16_start
        .checked_add(utf16_len)
        .ok_or(Error::UnexpectedEof)?;
    payload
        .get(utf16_start..utf16_end)
        .ok_or(Error::UnexpectedEof)?;

    let simple = WideStringOffsets {
        cch,
        flags,
        utf16_start,
        utf16_end,
        end: utf16_end,
    };

    let validate_following_formula_fields = |end: usize| -> bool {
        let Ok(cce) = read_u32(payload, end).map(|v| v as usize) else {
            return false;
        };
        let Some(rgce_offset) = end.checked_add(4) else {
            return false;
        };
        let Some(rgce_end) = rgce_offset.checked_add(cce) else {
            return false;
        };
        rgce_end <= payload.len()
    };

    let full = parse_wide_string_offsets(payload, 8, FlagsWidth::U16).ok();
    let full_valid = full
        .as_ref()
        .is_some_and(|ws| validate_following_formula_fields(ws.end));
    let simple_valid = validate_following_formula_fields(simple.end);

    match (full_valid, simple_valid) {
        (true, false) => Ok(full.expect("full offsets present")),
        (false, true) => Ok(simple),
        (true, true) => {
            // If the flags contain only the rich/phonetic indicators, trust the full parse.
            if flags & !(FLAG_RICH | FLAG_PHONETIC) == 0 {
                return Ok(full.expect("full offsets present"));
            }

            // When reserved bits are set, prefer the "simple" layout if the rich/phonetic bits
            // only contributed zero-length blocks (i.e. the parser suggests the cached string is
            // followed solely by `[cRun=0]` / `[cb=0]` fields). This avoids misinterpreting the
            // formula fields as string extras.
            let mut expected_delta = 0usize;
            if flags & FLAG_RICH != 0 {
                expected_delta = expected_delta.saturating_add(4);
            }
            if flags & FLAG_PHONETIC != 0 {
                expected_delta = expected_delta.saturating_add(4);
            }
            if let Some(full) = full {
                if full.end == simple.end.saturating_add(expected_delta) {
                    Ok(simple)
                } else {
                    Ok(full)
                }
            } else {
                Ok(simple)
            }
        }
        (false, false) => Err(Error::UnexpectedEof),
    }
}

/// Best-effort check for whether an `rgce` token stream references trailing `rgcb` bytes.
///
/// XLSB stores a primary formula token stream (`rgce`) and an optional trailing payload (`rgcb`).
/// Some ptgs (notably `PtgArray`) reference data stored in `rgcb`.
///
/// This is used by the worksheet patcher as a minimal safety check: if a caller replaces `rgce`
/// bytes and the formula references `rgcb`, they must supply the corresponding `rgcb` bytes.
///
/// Note: This intentionally does a minimal token-size parse so we don't accidentally match the
/// `0x20` `PtgArray` opcode inside unrelated token payloads (e.g. row indices).
#[doc(hidden)]
pub fn rgce_references_rgcb(rgce: &[u8]) -> bool {
    // Today we only *write* `rgcb` for `PtgArray` (array constants), so treat it as required when
    // `PtgArray` is present.

    fn has_remaining(buf: &[u8], i: usize, needed: usize) -> bool {
        buf.len().saturating_sub(i) >= needed
    }

    // Some ptgs (PtgMem*) embed a nested token stream of known length (`cce`). Scan those
    // sub-streams as well so we can detect `PtgArray` even when it appears inside the mem payload.
    //
    // The top-level `rgce` stream is scanned in "strict" mode: unknown ptgs abort the scan to
    // avoid desync/false positives.
    //
    // Nested mem subexpressions are scanned in a "best-effort" mode: unknown ptgs stop scanning
    // only that subexpression and resume scanning the outer stream. This avoids introducing false
    // negatives when a mem subexpression contains an unsupported token but the outer formula
    // later contains `PtgArray`.
    let mut stack: Vec<(&[u8], usize, bool)> = vec![(rgce, 0, true)];

    while let Some((buf, mut i, strict)) = stack.pop() {
        while i < buf.len() {
            let ptg = buf[i];
            i += 1;
            match ptg {
                // PtgArray (any class)
                0x20 | 0x40 | 0x60 => return true,

                // Binary operators and simple operators with no payload.
                0x03..=0x16 | 0x2F => {}

                // PtgStr: [cch: u16][utf16 chars...]
                0x17 => {
                    if !has_remaining(buf, i, 2) {
                        return false;
                    }
                    let cch = u16::from_le_bytes([buf[i], buf[i + 1]]) as usize;
                    i += 2;
                    let byte_len = match cch.checked_mul(2) {
                        Some(v) => v,
                        None => return false,
                    };
                    if !has_remaining(buf, i, byte_len) {
                        return false;
                    }
                    i += byte_len;
                }

                // PtgExtend / PtgExtendV / PtgExtendA: [etpg: u8][payload...]
                0x18 | 0x38 | 0x58 => {
                    if !has_remaining(buf, i, 1) {
                        return false;
                    }
                    let etpg = buf[i];
                    i += 1;
                    match etpg {
                        // etpg=0x19 is the structured reference payload (PtgList).
                        0x19 => {
                            if !has_remaining(buf, i, 12) {
                                return false;
                            }
                            i += 12;
                        }
                        // Unknown extend subtype: stop scanning to avoid desync/false positives.
                        _ => {
                            if strict {
                                return false;
                            }
                            break;
                        }
                    }
                }

                // PtgAttr: [grbit: u8][wAttr: u16] + optional jump table for tAttrChoose.
                0x19 => {
                    if !has_remaining(buf, i, 3) {
                        return false;
                    }
                    let grbit = buf[i];
                    let w_attr = u16::from_le_bytes([buf[i + 1], buf[i + 2]]) as usize;
                    i += 3;

                    const T_ATTR_CHOOSE: u8 = 0x04;
                    if grbit & T_ATTR_CHOOSE != 0 {
                        let needed = match w_attr.checked_mul(2) {
                            Some(v) => v,
                            None => return false,
                        };
                        if !has_remaining(buf, i, needed) {
                            return false;
                        }
                        i += needed;
                    }
                }

                // PtgErr: [code: u8]
                0x1C => {
                    if !has_remaining(buf, i, 1) {
                        return false;
                    }
                    i += 1;
                }
                // PtgBool: [b: u8]
                0x1D => {
                    if !has_remaining(buf, i, 1) {
                        return false;
                    }
                    i += 1;
                }
                // PtgInt: [u16]
                0x1E => {
                    if !has_remaining(buf, i, 2) {
                        return false;
                    }
                    i += 2;
                }
                // PtgNum: [f64]
                0x1F => {
                    if !has_remaining(buf, i, 8) {
                        return false;
                    }
                    i += 8;
                }

                // PtgFunc: [iftab: u16]
                0x21 | 0x41 | 0x61 => {
                    if !has_remaining(buf, i, 2) {
                        return false;
                    }
                    i += 2;
                }
                // PtgFuncVar: [argc: u8][iftab: u16]
                0x22 | 0x42 | 0x62 => {
                    if !has_remaining(buf, i, 3) {
                        return false;
                    }
                    i += 3;
                }

                // PtgName: [nameIndex: u32][unused: u16]
                0x23 | 0x43 | 0x63 => {
                    if !has_remaining(buf, i, 6) {
                        return false;
                    }
                    i += 6;
                }

                // PtgRef: [row: u32][col: u16]
                0x24 | 0x44 | 0x64 => {
                    if !has_remaining(buf, i, 6) {
                        return false;
                    }
                    i += 6;
                }
                // PtgArea: [rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
                0x25 | 0x45 | 0x65 => {
                    if !has_remaining(buf, i, 12) {
                        return false;
                    }
                    i += 12;
                }

                // PtgMem* tokens: [cce: u16][subexpression bytes...]
                0x26 | 0x46 | 0x66 | 0x27 | 0x47 | 0x67 | 0x28 | 0x48 | 0x68 | 0x29 | 0x49
                | 0x69 | 0x2E | 0x4E | 0x6E => {
                    if !has_remaining(buf, i, 2) {
                        return false;
                    }
                    let cce = u16::from_le_bytes([buf[i], buf[i + 1]]) as usize;
                    i += 2;
                    if !has_remaining(buf, i, cce) {
                        return false;
                    }
                    let subexpr = &buf[i..i + cce];
                    i += cce;

                    // Scan nested stream first, then resume after it.
                    stack.push((buf, i, strict));
                    stack.push((subexpr, 0, false));
                    break;
                }

                // PtgRefErr: [row: u32][col: u16]
                0x2A | 0x4A | 0x6A => {
                    if !has_remaining(buf, i, 6) {
                        return false;
                    }
                    i += 6;
                }
                // PtgAreaErr: [rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
                0x2B | 0x4B | 0x6B => {
                    if !has_remaining(buf, i, 12) {
                        return false;
                    }
                    i += 12;
                }

                // PtgRefN: [row_off: i32][col_off: i16]
                0x2C | 0x4C | 0x6C => {
                    if !has_remaining(buf, i, 6) {
                        return false;
                    }
                    i += 6;
                }
                // PtgAreaN: [rowFirst_off: i32][rowLast_off: i32][colFirst_off: i16][colLast_off: i16]
                0x2D | 0x4D | 0x6D => {
                    if !has_remaining(buf, i, 12) {
                        return false;
                    }
                    i += 12;
                }

                // PtgNameX: [ixti: u16][nameIndex: u16]
                0x39 | 0x59 | 0x79 => {
                    if !has_remaining(buf, i, 4) {
                        return false;
                    }
                    i += 4;
                }

                // PtgRef3d: [ixti: u16][row: u32][col: u16]
                0x3A | 0x5A | 0x7A => {
                    if !has_remaining(buf, i, 8) {
                        return false;
                    }
                    i += 8;
                }
                // PtgArea3d: [ixti: u16][rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
                0x3B | 0x5B | 0x7B => {
                    if !has_remaining(buf, i, 14) {
                        return false;
                    }
                    i += 14;
                }
                // PtgRefErr3d: [ixti: u16][row: u32][col: u16]
                0x3C | 0x5C | 0x7C => {
                    if !has_remaining(buf, i, 8) {
                        return false;
                    }
                    i += 8;
                }
                // PtgAreaErr3d: [ixti: u16][rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
                0x3D | 0x5D | 0x7D => {
                    if !has_remaining(buf, i, 14) {
                        return false;
                    }
                    i += 14;
                }

                // Unknown ptg: stop scanning to avoid desync/false positives.
                _ => {
                    if strict {
                        return false;
                    }
                    break;
                }
            }
        }
    }

    false
}

fn read_u16(data: &[u8], offset: usize) -> Result<u16, Error> {
    let bytes: [u8; 2] = data
        .get(offset..offset + 2)
        .ok_or(Error::UnexpectedEof)?
        .try_into()
        .unwrap();
    Ok(u16::from_le_bytes(bytes))
}

fn read_u8(data: &[u8], offset: usize) -> Result<u8, Error> {
    Ok(*data.get(offset).ok_or(Error::UnexpectedEof)?)
}

fn read_u32(data: &[u8], offset: usize) -> Result<u32, Error> {
    let bytes: [u8; 4] = data
        .get(offset..offset + 4)
        .ok_or(Error::UnexpectedEof)?
        .try_into()
        .unwrap();
    Ok(u32::from_le_bytes(bytes))
}

fn read_f64(data: &[u8], offset: usize) -> Result<f64, Error> {
    let bytes: [u8; 8] = data
        .get(offset..offset + 8)
        .ok_or(Error::UnexpectedEof)?
        .try_into()
        .unwrap();
    Ok(f64::from_le_bytes(bytes))
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

fn encode_rk_number(value: f64) -> Option<u32> {
    if !value.is_finite() {
        return None;
    }

    let int = value.round();
    if (value - int).abs() <= f64::EPSILON && int >= i32::MIN as f64 && int <= i32::MAX as f64 {
        let i = int as i32;
        return Some(((i as u32) << 2) | 0x02);
    }

    let scaled = (value * 100.0).round();
    if ((value * 100.0) - scaled).abs() <= 1e-6
        && scaled >= i32::MIN as f64
        && scaled <= i32::MAX as f64
    {
        let i = scaled as i32;
        return Some(((i as u32) << 2) | 0x03);
    }

    // Non-integer RK numbers store the top 30 bits of the IEEE754 f64 (with the low
    // 34 bits cleared) and set the low two bits to 0b00 (or 0b01 when scaled by 100).
    const LOW_34_MASK: u64 = (1u64 << 34) - 1;
    let bits = value.to_bits();
    if bits & LOW_34_MASK == 0 {
        let raw = (bits >> 32) as u32;
        if raw & 0x03 == 0 {
            return Some(raw);
        }
    }

    let scaled = value * 100.0;
    if scaled.is_finite() {
        let bits = scaled.to_bits();
        if bits & LOW_34_MASK == 0 {
            let raw = (bits >> 32) as u32;
            if raw & 0x03 == 0 {
                return Some(raw | 0x01);
            }
        }
    }

    None
}

fn decode_rk_number(raw: u32) -> f64 {
    let raw_i = raw as i32;
    let mut v = if raw_i & 0x02 != 0 {
        (raw_i >> 2) as f64
    } else {
        let shifted = raw & 0xFFFFFFFC;
        f64::from_bits((shifted as u64) << 32)
    };
    if raw_i & 0x01 != 0 {
        v /= 100.0;
    }
    v
}
