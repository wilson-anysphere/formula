use std::collections::{HashMap, HashSet};
use std::io::Write;

use formula_model::rich_text::{RichText, RichTextRun, RichTextRunStyle, Underline};
use formula_model::{CellRef, CellValue, Color, ErrorValue, Worksheet, WorksheetId};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::SheetMeta;
use crate::package::XlsxError;

use super::{CellValueKind, SharedStringKey, WriteError, XlsxDocument};

#[derive(Debug, Clone)]
struct SheetMlTags {
    row: String,
    c: String,
    f: String,
    v: String,
    is_: String,
    t: String,
    r_ph: String,
}

impl SheetMlTags {
    fn new(prefix: Option<&str>) -> Self {
        Self {
            row: super::prefixed_tag(prefix, "row"),
            c: super::prefixed_tag(prefix, "c"),
            f: super::prefixed_tag(prefix, "f"),
            v: super::prefixed_tag(prefix, "v"),
            is_: super::prefixed_tag(prefix, "is"),
            t: super::prefixed_tag(prefix, "t"),
            r_ph: super::prefixed_tag(prefix, "rPh"),
        }
    }
}

pub(super) fn patch_worksheet_xml(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    sheet: &Worksheet,
    original: &[u8],
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
) -> Result<Vec<u8>, WriteError> {
    let meta_sheet_id = cell_meta_sheet_ids
        .get(&sheet_meta.worksheet_id)
        .copied()
        .unwrap_or(sheet_meta.worksheet_id);
    let shared_formulas = super::shared_formula_groups(doc, meta_sheet_id);

    // Desired cells in row-major order.
    let mut desired_cells: Vec<(CellRef, &formula_model::Cell)> = Vec::new();
    if desired_cells.try_reserve(sheet.cell_count()).is_err() {
        return Err(XlsxError::AllocationFailure("patch_worksheet_xml desired cells").into());
    }
    for (cell_ref, cell) in sheet.iter_cells() {
        desired_cells.push((cell_ref, cell));
    }
    desired_cells.sort_by_key(|(r, _)| (r.row, r.col));
    let mut desired_idx = 0usize;

    let mut reader = Reader::from_reader(original);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    if out.try_reserve_exact(original.len()).is_err() {
        return Err(XlsxError::AllocationFailure("patch_worksheet_xml output").into());
    }
    let mut writer = Writer::new(out);
    let mut worksheet_prefix: Option<String> = None;
    let mut worksheet_has_default_ns = false;
    let mut saw_sheet_data = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if super::local_name(e.name().as_ref()) == b"worksheet" => {
                if worksheet_prefix.is_none() {
                    worksheet_prefix = super::element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                    worksheet_has_default_ns = super::worksheet_has_default_spreadsheetml_ns(&e)?;
                }
                writer.write_event(Event::Start(e.into_owned()))?;
            }
            Event::Start(e) if super::local_name(e.name().as_ref()) == b"sheetData" => {
                saw_sheet_data = true;
                let sheet_prefix = super::element_prefix(e.name().as_ref())
                    .and_then(|p| std::str::from_utf8(p).ok())
                    .map(|s| s.to_string());
                let tags = SheetMlTags::new(sheet_prefix.as_deref());
                writer.write_event(Event::Start(e.into_owned()))?;
                patch_sheet_data_contents(
                    doc,
                    sheet_meta,
                    sheet,
                    shared_lookup,
                    style_to_xf,
                    &shared_formulas,
                    cell_meta_sheet_ids,
                    changed_formula_cells,
                    &mut desired_cells,
                    &mut desired_idx,
                    &tags,
                    &mut reader,
                    &mut writer,
                    &mut buf,
                )?;
            }
            Event::Empty(e) if super::local_name(e.name().as_ref()) == b"sheetData" => {
                saw_sheet_data = true;
                if desired_idx >= desired_cells.len() {
                    writer.write_event(Event::Empty(e.into_owned()))?;
                } else {
                    // Expand <sheetData/> to insert new rows/cells.
                    let sheet_data_tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    let sheet_prefix = super::element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                    let tags = SheetMlTags::new(sheet_prefix.as_deref());
                    writer.write_event(Event::Start(e.into_owned()))?;
                    write_remaining_rows(
                        doc,
                        sheet_meta,
                        sheet,
                        shared_lookup,
                        style_to_xf,
                        &shared_formulas,
                        cell_meta_sheet_ids,
                        changed_formula_cells,
                        &mut desired_cells,
                        &mut desired_idx,
                        &tags,
                        &mut writer,
                    )?;
                    writer.write_event(Event::End(BytesEnd::new(sheet_data_tag.as_str())))?;
                }
            }
            Event::End(e) if super::local_name(e.name().as_ref()) == b"worksheet" => {
                if !saw_sheet_data && desired_idx < desired_cells.len() {
                    let prefix = if worksheet_has_default_ns {
                        None
                    } else {
                        worksheet_prefix.as_deref()
                    };
                    let sheet_data_tag = super::prefixed_tag(prefix, "sheetData");
                    let tags = SheetMlTags::new(prefix);
                    writer.write_event(Event::Start(BytesStart::new(sheet_data_tag.as_str())))?;
                    write_remaining_rows(
                        doc,
                        sheet_meta,
                        sheet,
                        shared_lookup,
                        style_to_xf,
                        &shared_formulas,
                        cell_meta_sheet_ids,
                        changed_formula_cells,
                        &mut desired_cells,
                        &mut desired_idx,
                        &tags,
                        &mut writer,
                    )?;
                    writer.write_event(Event::End(BytesEnd::new(sheet_data_tag.as_str())))?;
                }
                writer.write_event(Event::End(e.into_owned()))?;
            }
            Event::Eof => break,
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn patch_sheet_data_contents<R: std::io::BufRead, W: Write>(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    sheet: &Worksheet,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    shared_formulas: &HashMap<u32, super::SharedFormulaGroup>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
    desired_cells: &mut [(CellRef, &formula_model::Cell)],
    desired_idx: &mut usize,
    tags: &SheetMlTags,
    reader: &mut Reader<R>,
    writer: &mut Writer<W>,
    buf: &mut Vec<u8>,
) -> Result<(), WriteError> {
    loop {
        match reader.read_event_into(buf)? {
            Event::Start(e) if super::local_name(e.name().as_ref()) == b"row" => {
                let e = e.into_owned();
                let row_num = row_num_from_attrs(&e)?;
                let has_unknown_attrs = row_has_unknown_attrs(&e)?;
                let keep_due_to_row_props = sheet.row_properties(row_num.saturating_sub(1)).is_some();
                write_missing_rows_before(
                    doc,
                    sheet_meta,
                    sheet,
                    shared_lookup,
                    style_to_xf,
                    shared_formulas,
                    cell_meta_sheet_ids,
                    changed_formula_cells,
                    desired_cells,
                    desired_idx,
                    tags,
                    writer,
                    row_num,
                )?;

                let mut row_writer = Writer::new(Vec::new());
                write_patched_row_tag(&mut row_writer, e, sheet, style_to_xf, row_num, false)?;
                let outcome = patch_row_contents(
                    doc,
                    sheet_meta,
                    shared_lookup,
                    style_to_xf,
                    shared_formulas,
                    cell_meta_sheet_ids,
                    changed_formula_cells,
                    desired_cells,
                    desired_idx,
                    tags,
                    reader,
                    &mut row_writer,
                    buf,
                    row_num,
                )?;
                let row_bytes = row_writer.into_inner();
                if has_unknown_attrs
                    || keep_due_to_row_props
                    || outcome.wrote_cell
                    || outcome.wrote_other
                {
                    writer.get_mut().write_all(&row_bytes)?;
                }
            }
            Event::Empty(e) if super::local_name(e.name().as_ref()) == b"row" => {
                let e = e.into_owned();
                let row_num = row_num_from_attrs(&e)?;
                let has_unknown_attrs = row_has_unknown_attrs(&e)?;
                let keep_due_to_row_props = sheet.row_properties(row_num.saturating_sub(1)).is_some();
                write_missing_rows_before(
                    doc,
                    sheet_meta,
                    sheet,
                    shared_lookup,
                    style_to_xf,
                    shared_formulas,
                    cell_meta_sheet_ids,
                    changed_formula_cells,
                    desired_cells,
                    desired_idx,
                    tags,
                    writer,
                    row_num,
                )?;

                if peek_row(desired_cells, *desired_idx) != Some(row_num) {
                    if has_unknown_attrs || keep_due_to_row_props {
                        write_patched_row_tag(writer, e, sheet, style_to_xf, row_num, true)?;
                    } else {
                        // Drop empty placeholder rows that were introduced solely to hold cells
                        // that are no longer present in the model.
                        drop(e);
                    }
                } else {
                    // Row existed but was empty; expand and insert any desired cells.
                    let row_tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    let mut row_writer = Writer::new(Vec::new());
                    write_patched_row_tag(&mut row_writer, e, sheet, style_to_xf, row_num, false)?;
                    write_remaining_cells_in_row(
                        doc,
                        sheet_meta,
                        shared_lookup,
                        style_to_xf,
                        shared_formulas,
                        cell_meta_sheet_ids,
                        changed_formula_cells,
                        desired_cells,
                        desired_idx,
                        tags,
                        &mut row_writer,
                        row_num,
                    )?;
                    row_writer.write_event(Event::End(BytesEnd::new(row_tag.as_str())))?;
                    writer.get_mut().write_all(&row_writer.into_inner())?;
                }
            }
            Event::End(e) if super::local_name(e.name().as_ref()) == b"sheetData" => {
                write_remaining_rows(
                    doc,
                    sheet_meta,
                    sheet,
                    shared_lookup,
                    style_to_xf,
                    shared_formulas,
                    cell_meta_sheet_ids,
                    changed_formula_cells,
                    desired_cells,
                    desired_idx,
                    tags,
                    writer,
                )?;
                writer.write_event(Event::End(e.into_owned()))?;
                break;
            }
            Event::Eof => break,
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(())
}

fn patch_row_contents<R: std::io::BufRead, W: Write>(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    shared_formulas: &HashMap<u32, super::SharedFormulaGroup>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
    desired_cells: &mut [(CellRef, &formula_model::Cell)],
    desired_idx: &mut usize,
    tags: &SheetMlTags,
    reader: &mut Reader<R>,
    writer: &mut Writer<W>,
    buf: &mut Vec<u8>,
    row_num: u32,
) -> Result<RowPatchOutcome, WriteError> {
    let mut wrote_cell = false;
    let mut wrote_other = false;
    loop {
        match reader.read_event_into(buf)? {
            Event::Start(e) if super::local_name(e.name().as_ref()) == b"c" => {
                let cell_events = collect_full_element(Event::Start(e.into_owned()), reader, buf)?;
                let (cell_ref, col_1_based) = match cell_ref_from_cell_events(&cell_events)? {
                    Some(v) => v,
                    None => {
                        // Can't address the cell; preserve it.
                        write_events(writer, cell_events)?;
                        continue;
                    }
                };

                let idx_before = *desired_idx;
                write_missing_cells_before_col(
                    doc,
                    sheet_meta,
                    shared_lookup,
                    style_to_xf,
                    shared_formulas,
                    cell_meta_sheet_ids,
                    changed_formula_cells,
                    desired_cells,
                    desired_idx,
                    tags,
                    writer,
                    row_num,
                    col_1_based,
                )?;
                wrote_cell |= *desired_idx > idx_before;
                patch_or_copy_cell(
                    doc,
                    sheet_meta,
                    shared_lookup,
                    style_to_xf,
                    shared_formulas,
                    cell_meta_sheet_ids,
                    changed_formula_cells,
                    desired_cells,
                    desired_idx,
                    tags,
                    writer,
                    row_num,
                    col_1_based,
                    cell_ref,
                    cell_events,
                )
                .map(|wrote| {
                    wrote_cell |= wrote;
                })?;
            }
            Event::Empty(e) if super::local_name(e.name().as_ref()) == b"c" => {
                let cell_events = vec![Event::Empty(e.into_owned())];
                let (cell_ref, col_1_based) = match cell_ref_from_cell_events(&cell_events)? {
                    Some(v) => v,
                    None => {
                        write_events(writer, cell_events)?;
                        continue;
                    }
                };

                let idx_before = *desired_idx;
                write_missing_cells_before_col(
                    doc,
                    sheet_meta,
                    shared_lookup,
                    style_to_xf,
                    shared_formulas,
                    cell_meta_sheet_ids,
                    changed_formula_cells,
                    desired_cells,
                    desired_idx,
                    tags,
                    writer,
                    row_num,
                    col_1_based,
                )?;
                wrote_cell |= *desired_idx > idx_before;
                patch_or_copy_cell(
                    doc,
                    sheet_meta,
                    shared_lookup,
                    style_to_xf,
                    shared_formulas,
                    cell_meta_sheet_ids,
                    changed_formula_cells,
                    desired_cells,
                    desired_idx,
                    tags,
                    writer,
                    row_num,
                    col_1_based,
                    cell_ref,
                    cell_events,
                )
                .map(|wrote| {
                    wrote_cell |= wrote;
                })?;
            }
            Event::End(e) if super::local_name(e.name().as_ref()) == b"row" => {
                // Append any new cells at the end of the row.
                let idx_before = *desired_idx;
                write_remaining_cells_in_row(
                    doc,
                    sheet_meta,
                    shared_lookup,
                    style_to_xf,
                    shared_formulas,
                    cell_meta_sheet_ids,
                    changed_formula_cells,
                    desired_cells,
                    desired_idx,
                    tags,
                    writer,
                    row_num,
                )?;
                wrote_cell |= *desired_idx > idx_before;

                writer.write_event(Event::End(e.into_owned()))?;
                break;
            }
            // Non-cell element inside the row (eg extLst). Ensure any new cells are emitted
            // before it so we keep cells grouped at the start of the row.
            Event::Start(e) if super::local_name(e.name().as_ref()) != b"c" => {
                wrote_other = true;
                let idx_before = *desired_idx;
                write_remaining_cells_in_row(
                    doc,
                    sheet_meta,
                    shared_lookup,
                    style_to_xf,
                    shared_formulas,
                    cell_meta_sheet_ids,
                    changed_formula_cells,
                    desired_cells,
                    desired_idx,
                    tags,
                    writer,
                    row_num,
                )?;
                wrote_cell |= *desired_idx > idx_before;
                writer.write_event(Event::Start(e.into_owned()))?;
            }
            Event::Empty(e) if super::local_name(e.name().as_ref()) != b"c" => {
                wrote_other = true;
                let idx_before = *desired_idx;
                write_remaining_cells_in_row(
                    doc,
                    sheet_meta,
                    shared_lookup,
                    style_to_xf,
                    shared_formulas,
                    cell_meta_sheet_ids,
                    changed_formula_cells,
                    desired_cells,
                    desired_idx,
                    tags,
                    writer,
                    row_num,
                )?;
                wrote_cell |= *desired_idx > idx_before;
                writer.write_event(Event::Empty(e.into_owned()))?;
            }
            Event::Text(e) => {
                let text = e.unescape()?.into_owned();
                if !text.chars().all(|c| c.is_whitespace()) {
                    wrote_other = true;
                }
                writer.write_event(Event::Text(e.into_owned()))?;
            }
            Event::CData(e) => {
                let text = String::from_utf8_lossy(e.as_ref());
                if !text.chars().all(|c| c.is_whitespace()) {
                    wrote_other = true;
                }
                writer.write_event(Event::CData(e.into_owned()))?;
            }
            Event::Eof => break,
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(RowPatchOutcome {
        wrote_cell,
        wrote_other,
    })
}

fn patch_or_copy_cell<W: Write>(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    shared_formulas: &HashMap<u32, super::SharedFormulaGroup>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
    desired_cells: &mut [(CellRef, &formula_model::Cell)],
    desired_idx: &mut usize,
    tags: &SheetMlTags,
    writer: &mut Writer<W>,
    row_num: u32,
    col_1_based: u32,
    cell_ref: CellRef,
    cell_events: Vec<Event<'static>>,
) -> Result<bool, WriteError> {
    let original = parse_cell_semantics(doc, sheet_meta, &cell_events, &doc.shared_strings)?;

    let desired = desired_cells
        .get(*desired_idx)
        .copied()
        .filter(|(r, _)| r.row + 1 == row_num && r.col + 1 == col_1_based);

    match desired {
        Some((_, desired_cell)) => {
            // This cell is present in the model; decide whether it changed.
            let mut desired_semantics = CellSemantics::from_model(desired_cell, style_to_xf);
            if original.formula.is_none() {
                if let Some(model_formula) = desired_semantics.formula.as_deref() {
                    let meta = super::lookup_cell_meta(
                        doc,
                        cell_meta_sheet_ids,
                        sheet_meta.worksheet_id,
                        cell_ref,
                    );
                    if let Some(meta_formula) = meta.and_then(|m| m.formula.as_ref()) {
                        let is_shared_follower = meta_formula.t.as_deref() == Some("shared")
                            && meta_formula.reference.is_none()
                            && meta_formula.shared_index.is_some()
                            && meta_formula.file_text.is_empty();
                        if is_shared_follower {
                            if let Some(shared_index) = meta_formula.shared_index {
                                if let Some(expected) = super::shared_formula_expected(
                                    shared_formulas,
                                    shared_index,
                                    cell_ref,
                                ) {
                                    if expected == model_formula {
                                        desired_semantics.formula = None;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            *desired_idx += 1;

             if desired_semantics == original {
                 write_events(writer, cell_events)?;
             } else {
                // `vm="..."` is a value-metadata pointer into `xl/metadata.xml` (rich values /
                // images-in-cell).
                //
                // We generally preserve unknown SpreadsheetML metadata for fidelity, even when the
                // cached value changes.
                //
                // The exception is the rich-value placeholder semantics used by in-cell images:
                //
                // Excel stores the image itself in `xl/richData/*` and represents the cell's cached
                // value as a placeholder (commonly the error `#VALUE!`, but some producers use a
                // numeric value like `0`) with a `vm` pointer. When the cell is edited away from that
                // placeholder value, we must drop `vm` to avoid leaving dangling rich-value metadata
                // pointers.
                let original_is_rich_placeholder =
                    matches!(&original.value, CellValue::Error(ErrorValue::Value))
                        || matches!(&original.value, CellValue::Number(n) if *n == 0.0);
                let desired_is_rich_placeholder =
                    matches!(&desired_semantics.value, CellValue::Error(ErrorValue::Value))
                        || matches!(&desired_semantics.value, CellValue::Number(n) if *n == 0.0);
                let preserve_vm = !(original_is_rich_placeholder && !desired_is_rich_placeholder);

                // When rewriting an inline string cell without changing the visible text, preserve
                // the original `<is>...</is>` subtree byte-for-byte so we don't drop unsupported
                // structures like phonetic runs (`<rPh>`) or rich run formatting.
                let original_cell_type = cell_type_from_cell_events(&cell_events)?;
                let preserved_inline_is = (original_cell_type.as_deref() == Some("inlineStr")
                    && cell_value_semantics_eq(&desired_semantics.value, &original.value)
                    && desired_semantics.phonetic == original.phonetic)
                .then(|| extract_is_subtree(&cell_events))
                .flatten();
                write_updated_cell(
                    doc,
                    sheet_meta,
                    shared_lookup,
                    style_to_xf,
                    shared_formulas,
                    cell_meta_sheet_ids,
                    changed_formula_cells,
                    tags,
                    writer,
                    cell_ref,
                    desired_cell,
                    preserve_vm,
                    cell_events.first(),
                    preserved_inline_is,
                    extract_preserved_cell_children(&cell_events),
                )?;
              }
              Ok(true)
          }
        None => {
            // Cell not represented in the model.
            if original.is_truly_empty() {
                // Preserve unknown metadata-only cells.
                write_events(writer, cell_events)?;
                Ok(true)
            } else {
                // The cell existed in the original file but was cleared from the model.
                // Drop it from the output.
                Ok(false)
            }
        }
    }
}

#[derive(Debug, Clone, Copy, Default)]
struct RowPatchOutcome {
    wrote_cell: bool,
    wrote_other: bool,
}

fn write_missing_rows_before<W: Write>(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    sheet: &Worksheet,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    shared_formulas: &HashMap<u32, super::SharedFormulaGroup>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
    desired_cells: &mut [(CellRef, &formula_model::Cell)],
    desired_idx: &mut usize,
    tags: &SheetMlTags,
    writer: &mut Writer<W>,
    row_num: u32,
) -> Result<(), WriteError> {
    while let Some(next_row) = peek_row(desired_cells, *desired_idx) {
        if next_row >= row_num {
            break;
        }
        write_new_row(
            doc,
            sheet_meta,
            sheet,
            shared_lookup,
            style_to_xf,
            shared_formulas,
            cell_meta_sheet_ids,
            changed_formula_cells,
            desired_cells,
            desired_idx,
            tags,
            writer,
            next_row,
        )?;
    }
    Ok(())
}

fn write_remaining_rows<W: Write>(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    sheet: &Worksheet,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    shared_formulas: &HashMap<u32, super::SharedFormulaGroup>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
    desired_cells: &mut [(CellRef, &formula_model::Cell)],
    desired_idx: &mut usize,
    tags: &SheetMlTags,
    writer: &mut Writer<W>,
) -> Result<(), WriteError> {
    while let Some(next_row) = peek_row(desired_cells, *desired_idx) {
        write_new_row(
            doc,
            sheet_meta,
            sheet,
            shared_lookup,
            style_to_xf,
            shared_formulas,
            cell_meta_sheet_ids,
            changed_formula_cells,
            desired_cells,
            desired_idx,
            tags,
            writer,
            next_row,
        )?;
    }
    Ok(())
}

fn write_new_row<W: Write>(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    sheet: &Worksheet,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    shared_formulas: &HashMap<u32, super::SharedFormulaGroup>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
    desired_cells: &mut [(CellRef, &formula_model::Cell)],
    desired_idx: &mut usize,
    tags: &SheetMlTags,
    writer: &mut Writer<W>,
    row_num: u32,
) -> Result<(), WriteError> {
    let mut row_start = BytesStart::new(tags.row.as_str());
    let row_str = row_num.to_string();
    row_start.push_attribute(("r", row_str.as_str()));
    let outline_entry = sheet.outline.rows.entry(row_num);
    let row_props = sheet.row_properties(row_num.saturating_sub(1));

    let height_str = row_props.and_then(|props| props.height.map(|h| h.to_string()));
    if let Some(height_str) = &height_str {
        row_start.push_attribute(("ht", height_str.as_str()));
        row_start.push_attribute(("customHeight", "1"));
    }

    let is_hidden = row_props.is_some_and(|p| p.hidden) || outline_entry.hidden.is_hidden();
    if is_hidden {
        row_start.push_attribute(("hidden", "1"));
    }

    let outline_level_str = (outline_entry.level > 0).then(|| outline_entry.level.to_string());
    if let Some(level_str) = &outline_level_str {
        row_start.push_attribute(("outlineLevel", level_str.as_str()));
    }
    if outline_entry.collapsed {
        row_start.push_attribute(("collapsed", "1"));
    }

    let style_xf_str = row_props
        .and_then(|props| props.style_id)
        .filter(|style_id| *style_id != 0)
        .and_then(|style_id| style_to_xf.get(&style_id).copied())
        .map(|xf| xf.to_string());
    if let Some(style_xf_str) = &style_xf_str {
        row_start.push_attribute(("s", style_xf_str.as_str()));
        row_start.push_attribute(("customFormat", "1"));
    }

    writer.write_event(Event::Start(row_start))?;
    write_remaining_cells_in_row(
        doc,
        sheet_meta,
        shared_lookup,
        style_to_xf,
        shared_formulas,
        cell_meta_sheet_ids,
        changed_formula_cells,
        desired_cells,
        desired_idx,
        tags,
        writer,
        row_num,
    )?;
    writer.write_event(Event::End(BytesEnd::new(tags.row.as_str())))?;
    Ok(())
}

fn write_missing_cells_before_col<W: Write>(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    shared_formulas: &HashMap<u32, super::SharedFormulaGroup>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
    desired_cells: &mut [(CellRef, &formula_model::Cell)],
    desired_idx: &mut usize,
    tags: &SheetMlTags,
    writer: &mut Writer<W>,
    row_num: u32,
    col_limit_1_based: u32,
) -> Result<(), WriteError> {
    while let Some((cell_ref, cell)) = desired_cells.get(*desired_idx).copied() {
        if cell_ref.row + 1 != row_num {
            break;
        }
        let col = cell_ref.col + 1;
        if col >= col_limit_1_based {
            break;
        }
        write_updated_cell(
            doc,
            sheet_meta,
            shared_lookup,
            style_to_xf,
            shared_formulas,
            cell_meta_sheet_ids,
            changed_formula_cells,
            tags,
            writer,
            cell_ref,
            cell,
            true,
            None,
            None,
            Vec::new(),
        )?;
        *desired_idx += 1;
    }
    Ok(())
}

fn write_remaining_cells_in_row<W: Write>(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    shared_formulas: &HashMap<u32, super::SharedFormulaGroup>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
    desired_cells: &mut [(CellRef, &formula_model::Cell)],
    desired_idx: &mut usize,
    tags: &SheetMlTags,
    writer: &mut Writer<W>,
    row_num: u32,
) -> Result<(), WriteError> {
    while let Some((cell_ref, cell)) = desired_cells.get(*desired_idx).copied() {
        if cell_ref.row + 1 != row_num {
            break;
        }
        write_updated_cell(
            doc,
            sheet_meta,
            shared_lookup,
            style_to_xf,
            shared_formulas,
            cell_meta_sheet_ids,
            changed_formula_cells,
            tags,
            writer,
            cell_ref,
            cell,
            true,
            None,
            None,
            Vec::new(),
        )?;
        *desired_idx += 1;
    }
    Ok(())
}

fn write_updated_cell<W: Write>(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    shared_formulas: &HashMap<u32, super::SharedFormulaGroup>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
    tags: &SheetMlTags,
    writer: &mut Writer<W>,
    cell_ref: CellRef,
    cell: &formula_model::Cell,
    preserve_vm: bool,
    original_cell_event: Option<&Event<'static>>,
    preserved_inline_is: Option<Vec<Event<'static>>>,
    preserved_children: Vec<Event<'static>>,
) -> Result<(), WriteError> {
    let original_start = original_cell_event.and_then(|ev| match ev {
        Event::Start(e) => Some(e),
        Event::Empty(e) => Some(e),
        _ => None,
    });
    let cell_tag = original_start
        .map(|start| String::from_utf8_lossy(start.name().as_ref()).into_owned())
        .unwrap_or_else(|| tags.c.clone());

    let tags_prefix =
        super::element_prefix(tags.c.as_bytes()).and_then(|p| std::str::from_utf8(p).ok());
    let cell_prefix =
        super::element_prefix(cell_tag.as_bytes()).and_then(|p| std::str::from_utf8(p).ok());
    let cell_tags = (cell_prefix != tags_prefix).then(|| SheetMlTags::new(cell_prefix));
    let f_tag = cell_tags
        .as_ref()
        .map(|t| t.f.as_str())
        .unwrap_or(tags.f.as_str());
    let v_tag = cell_tags
        .as_ref()
        .map(|t| t.v.as_str())
        .unwrap_or(tags.v.as_str());
    let is_tag = cell_tags
        .as_ref()
        .map(|t| t.is_.as_str())
        .unwrap_or(tags.is_.as_str());
    let t_tag = cell_tags
        .as_ref()
        .map(|t| t.t.as_str())
        .unwrap_or(tags.t.as_str());
    let r_ph_tag = cell_tags
        .as_ref()
        .map(|t| t.r_ph.as_str())
        .unwrap_or(tags.r_ph.as_str());

    let mut original_has_vm = false;
    let mut original_has_cm = false;
    let mut preserved_attrs: Vec<(String, String)> = Vec::new();
    if let Some(start) = original_start {
        for attr in start.attributes() {
            let attr = attr?;
            match attr.key.as_ref() {
                b"r" | b"s" | b"t" => {}
                _ => {
                    if attr.key.as_ref() == b"vm" {
                        original_has_vm = true;
                    } else if attr.key.as_ref() == b"cm" {
                        original_has_cm = true;
                    }
                    let key = std::str::from_utf8(attr.key.as_ref())
                        .unwrap_or("")
                        .to_string();
                    let value = attr.unescape_value()?.into_owned();
                    preserved_attrs.push((key, value));
                }
            }
        }
    }

    let mut a1 = String::new();
    formula_model::push_a1_cell_ref(cell_ref.row, cell_ref.col, false, false, &mut a1);
    let meta = super::lookup_cell_meta(doc, cell_meta_sheet_ids, sheet_meta.worksheet_id, cell_ref);
    // `CellMeta.vm` is captured from the original file to allow lossless round-trip (including
    // preserving formatting like leading zeros). For existing cells we *also* have access to the
    // original `vm` attribute via `preserved_attrs`.
    //
    // Treat `meta.vm` as an explicit override only when:
    // - the original cell did not have a `vm` attribute (e.g. inserted cell / synthesized meta), or
    // - the stored value differs from the original attribute (caller mutated metadata).
    //
    // Otherwise, rely on the preserved original attribute so we can still drop `vm` when the cell
    // value changes away from rich-value semantics.
    let meta_vm = meta.and_then(|m| m.vm.clone()).and_then(|vm| {
        if original_has_vm {
            let original_vm = preserved_attrs
                .iter()
                .find(|(k, _)| k == "vm")
                .map(|(_, v)| v.as_str());
            if original_vm.is_some_and(|orig| orig == vm.as_str()) {
                None
            } else {
                Some(vm)
            }
        } else {
            Some(vm)
        }
    });
    let meta_cm = meta.and_then(|m| m.cm.clone());
    let mut preserved_inline_is = preserved_inline_is;
    let mut value_kind = super::effective_value_kind(meta, cell);
    if preserved_inline_is.is_some() {
        // We have an original `<is>...</is>` payload that we want to preserve byte-for-byte.
        // Force `t="inlineStr"` so the output remains a valid inline string cell.
        value_kind = CellValueKind::InlineString;
    }
    let meta_sheet_id = cell_meta_sheet_ids
        .get(&sheet_meta.worksheet_id)
        .copied()
        .unwrap_or(sheet_meta.worksheet_id);
    let clear_cached_value = cell
        .formula
        .as_deref()
        .is_some_and(|f| !super::strip_leading_equals(f).is_empty())
        && changed_formula_cells.contains(&(meta_sheet_id, cell_ref));

    let model_formula = cell.formula.as_deref();
    let mut preserve_textless_shared = false;
    let mut formula_meta = formula_meta_for_cell(meta, cell);
    if let (Some(display), Some(meta)) = (model_formula, formula_meta.as_mut()) {
        if meta.t.as_deref() == Some("shared") && meta.file_text.is_empty() {
            if let Some(shared_index) = meta.shared_index {
                if let Some(expected) =
                    super::shared_formula_expected(shared_formulas, shared_index, cell_ref)
                {
                    if expected == super::strip_leading_equals(display) {
                        preserve_textless_shared = true;
                    } else {
                        meta.t = None;
                        meta.reference = None;
                        meta.shared_index = None;
                        meta.file_text = crate::formula_text::add_xlfn_prefixes(
                            super::strip_leading_equals(display),
                        );
                    }
                } else {
                    meta.t = None;
                    meta.reference = None;
                    meta.shared_index = None;
                    meta.file_text = crate::formula_text::add_xlfn_prefixes(
                        super::strip_leading_equals(display),
                    );
                }
            } else {
                meta.t = None;
                meta.reference = None;
                meta.shared_index = None;
                meta.file_text =
                    crate::formula_text::add_xlfn_prefixes(super::strip_leading_equals(display));
            }
        }
    }

    let has_value = !clear_cached_value && !matches!(cell.value, CellValue::Empty);
    let has_children = formula_meta.is_some() || has_value || !preserved_children.is_empty();

    let style_xf_str = (cell.style_id != 0)
        .then(|| style_to_xf.get(&cell.style_id))
        .flatten()
        .map(|xf_index| xf_index.to_string());

    let mut c_start = BytesStart::new(cell_tag.as_str());
    c_start.push_attribute(("r", a1.as_str()));
    if let Some(style_xf_str) = &style_xf_str {
        c_start.push_attribute(("s", style_xf_str.as_str()));
    }
    if has_value {
        match &value_kind {
            CellValueKind::SharedString { .. } => c_start.push_attribute(("t", "s")),
            CellValueKind::InlineString => c_start.push_attribute(("t", "inlineStr")),
            CellValueKind::Bool => c_start.push_attribute(("t", "b")),
            CellValueKind::Error => c_start.push_attribute(("t", "e")),
            CellValueKind::Str => c_start.push_attribute(("t", "str")),
            CellValueKind::Number => {}
            CellValueKind::Other { t } => c_start.push_attribute(("t", t.as_str())),
        }
    }
    // `vm="..."` is a SpreadsheetML value-metadata pointer (typically into `xl/metadata.xml`).
    //
    // This is used by Excel rich values (linked data types, in-cell images, etc). When patching
    // worksheet XML, drop `vm` when the cell's cached value changes away from rich-value placeholder
    // semantics, to avoid leaving dangling metadata pointers.
    //
    // If the caller explicitly overrides the `vm` value in `CellMeta`, prefer it over what we saw
    // in the original `<c>` element.
    let original_vm = preserved_attrs
        .iter()
        .find(|(k, _)| k == "vm")
        .map(|(_, v)| v.as_str());
    let meta_vm_override = meta_vm
        .as_deref()
        .is_some_and(|vm| !original_has_vm || Some(vm) != original_vm);
    let drop_vm = original_has_vm && !preserve_vm && !meta_vm_override;

    if meta_vm_override {
        if let Some(vm) = meta_vm.as_deref() {
            c_start.push_attribute(("vm", vm));
        }
    } else if !original_has_vm {
        if let Some(vm) = meta_vm.as_deref() {
            c_start.push_attribute(("vm", vm));
        }
    }
    if !original_has_cm {
        if let Some(cm) = meta_cm.as_deref() {
            c_start.push_attribute(("cm", cm));
        }
    }
    for (k, v) in &preserved_attrs {
        if k == "vm" && (drop_vm || meta_vm_override) {
            continue;
        }
        c_start.push_attribute((k.as_str(), v.as_str()));
    }

    if !has_children {
        writer.write_event(Event::Empty(c_start))?;
        return Ok(());
    }

    writer.write_event(Event::Start(c_start))?;

    if let Some(formula_meta) = formula_meta {
        let mut f_start = BytesStart::new(f_tag);
        let si_str = formula_meta.shared_index.map(|si| si.to_string());
        if let Some(t) = &formula_meta.t {
            f_start.push_attribute(("t", t.as_str()));
        }
        if let Some(reference) = &formula_meta.reference {
            f_start.push_attribute(("ref", reference.as_str()));
        }
        if let Some(si_str) = &si_str {
            f_start.push_attribute(("si", si_str.as_str()));
        }
        if let Some(aca) = formula_meta.always_calc {
            f_start.push_attribute(("aca", if aca { "1" } else { "0" }));
        }

        let file_text = if preserve_textless_shared {
            String::new()
        } else {
            super::formula_file_text(&formula_meta, model_formula)
        };
        if file_text.is_empty() {
            writer.write_event(Event::Empty(f_start))?;
        } else {
            writer.write_event(Event::Start(f_start))?;
            writer.write_event(Event::Text(BytesText::new(&file_text)))?;
            writer.write_event(Event::End(BytesEnd::new(f_tag)))?;
        }
    }

    if !clear_cached_value {
        match &cell.value {
            CellValue::Empty => {}
            CellValue::String(s) if matches!(&value_kind, CellValueKind::Other { .. }) => {
                writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                let v = super::raw_or_other(meta, s);
                writer.write_event(Event::Text(BytesText::new(&v)))?;
                writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
            }
            CellValue::Number(n) => {
                writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                let v = super::raw_or_number(meta, *n);
                writer.write_event(Event::Text(BytesText::new(&v)))?;
                writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
            }
            CellValue::Boolean(b) => {
                writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                let v = super::raw_or_bool(meta, *b);
                writer.write_event(Event::Text(BytesText::new(v)))?;
                writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
            }
            CellValue::Error(err) => {
                writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                let v = super::raw_or_error(meta, *err);
                writer.write_event(Event::Text(BytesText::new(&v)))?;
                writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
            }
            value @ CellValue::String(s) => match &value_kind {
                CellValueKind::SharedString { .. } => {
                    let idx = super::shared_string_index(doc, meta, value, shared_lookup);
                    writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                    writer.write_event(Event::Text(BytesText::new(&idx.to_string())))?;
                    writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
                }
                CellValueKind::InlineString => {
                    if let Some(is_events) = preserved_inline_is.take() {
                        for ev in is_events {
                            writer.write_event(ev)?;
                        }
                    } else {
                        writer.write_event(Event::Start(BytesStart::new(is_tag)))?;
                        let mut t_start = BytesStart::new(t_tag);
                        if super::needs_space_preserve(s) {
                            t_start.push_attribute(("xml:space", "preserve"));
                        }
                        writer.write_event(Event::Start(t_start))?;
                        writer.write_event(Event::Text(BytesText::new(s)))?;
                        writer.write_event(Event::End(BytesEnd::new(t_tag)))?;
                        if let Some(phonetic) = cell.phonetic.as_deref() {
                            let base_len = s.chars().count().to_string();
                            let mut rph_start = BytesStart::new(r_ph_tag);
                            rph_start.push_attribute(("sb", "0"));
                            rph_start.push_attribute(("eb", base_len.as_str()));
                            writer.write_event(Event::Start(rph_start))?;

                            let mut ph_t_start = BytesStart::new(t_tag);
                            if super::needs_space_preserve(phonetic) {
                                ph_t_start.push_attribute(("xml:space", "preserve"));
                            }
                            writer.write_event(Event::Start(ph_t_start))?;
                            writer.write_event(Event::Text(BytesText::new(phonetic)))?;
                            writer.write_event(Event::End(BytesEnd::new(t_tag)))?;
                            writer.write_event(Event::End(BytesEnd::new(r_ph_tag)))?;
                        }
                        writer.write_event(Event::End(BytesEnd::new(is_tag)))?;
                    }
                }
                CellValueKind::Str => {
                    writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                    let v = super::raw_or_str(meta, s);
                    writer.write_event(Event::Text(BytesText::new(&v)))?;
                    writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
                }
                _ => {
                    let idx = super::shared_string_index(doc, meta, value, shared_lookup);
                    writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                    writer.write_event(Event::Text(BytesText::new(&idx.to_string())))?;
                    writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
                }
            },
            value @ CellValue::Entity(entity) => {
                let s = entity.display_value.as_str();
                match &value_kind {
                    CellValueKind::SharedString { .. } => {
                        let idx = super::shared_string_index(doc, meta, value, shared_lookup);
                        writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                        writer.write_event(Event::Text(BytesText::new(&idx.to_string())))?;
                        writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
                    }
                    CellValueKind::InlineString => {
                        if let Some(is_events) = preserved_inline_is.take() {
                            for ev in is_events {
                                writer.write_event(ev)?;
                            }
                        } else {
                            writer.write_event(Event::Start(BytesStart::new(is_tag)))?;
                            let mut t_start = BytesStart::new(t_tag);
                            if super::needs_space_preserve(s) {
                                t_start.push_attribute(("xml:space", "preserve"));
                            }
                            writer.write_event(Event::Start(t_start))?;
                            writer.write_event(Event::Text(BytesText::new(s)))?;
                            writer.write_event(Event::End(BytesEnd::new(t_tag)))?;
                            if let Some(phonetic) = cell.phonetic.as_deref() {
                                let base_len = s.chars().count().to_string();
                                let mut rph_start = BytesStart::new(r_ph_tag);
                                rph_start.push_attribute(("sb", "0"));
                                rph_start.push_attribute(("eb", base_len.as_str()));
                                writer.write_event(Event::Start(rph_start))?;

                                let mut ph_t_start = BytesStart::new(t_tag);
                                if super::needs_space_preserve(phonetic) {
                                    ph_t_start.push_attribute(("xml:space", "preserve"));
                                }
                                writer.write_event(Event::Start(ph_t_start))?;
                                writer.write_event(Event::Text(BytesText::new(phonetic)))?;
                                writer.write_event(Event::End(BytesEnd::new(t_tag)))?;
                                writer.write_event(Event::End(BytesEnd::new(r_ph_tag)))?;
                            }
                            writer.write_event(Event::End(BytesEnd::new(is_tag)))?;
                        }
                    }
                    CellValueKind::Str => {
                        writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                        let v = super::raw_or_str(meta, s);
                        writer.write_event(Event::Text(BytesText::new(&v)))?;
                        writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
                    }
                    _ => {
                        let idx = super::shared_string_index(doc, meta, value, shared_lookup);
                        writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                        writer.write_event(Event::Text(BytesText::new(&idx.to_string())))?;
                        writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
                    }
                }
            }
            value @ CellValue::Record(record) => {
                let display = record.to_string();
                let s = display.as_str();
                match &value_kind {
                    CellValueKind::SharedString { .. } => {
                        let idx = super::shared_string_index(doc, meta, value, shared_lookup);
                        writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                        writer.write_event(Event::Text(BytesText::new(&idx.to_string())))?;
                        writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
                    }
                    CellValueKind::InlineString => {
                        if let Some(is_events) = preserved_inline_is.take() {
                            for ev in is_events {
                                writer.write_event(ev)?;
                            }
                        } else {
                            writer.write_event(Event::Start(BytesStart::new(is_tag)))?;
                            let mut t_start = BytesStart::new(t_tag);
                            if super::needs_space_preserve(s) {
                                t_start.push_attribute(("xml:space", "preserve"));
                            }
                            writer.write_event(Event::Start(t_start))?;
                            writer.write_event(Event::Text(BytesText::new(s)))?;
                            writer.write_event(Event::End(BytesEnd::new(t_tag)))?;
                            if let Some(phonetic) = cell.phonetic.as_deref() {
                                let base_len = s.chars().count().to_string();
                                let mut rph_start = BytesStart::new(r_ph_tag);
                                rph_start.push_attribute(("sb", "0"));
                                rph_start.push_attribute(("eb", base_len.as_str()));
                                writer.write_event(Event::Start(rph_start))?;

                                let mut ph_t_start = BytesStart::new(t_tag);
                                if super::needs_space_preserve(phonetic) {
                                    ph_t_start.push_attribute(("xml:space", "preserve"));
                                }
                                writer.write_event(Event::Start(ph_t_start))?;
                                writer.write_event(Event::Text(BytesText::new(phonetic)))?;
                                writer.write_event(Event::End(BytesEnd::new(t_tag)))?;
                                writer.write_event(Event::End(BytesEnd::new(r_ph_tag)))?;
                            }
                            writer.write_event(Event::End(BytesEnd::new(is_tag)))?;
                        }
                    }
                    CellValueKind::Str => {
                        writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                        let v = super::raw_or_str(meta, s);
                        writer.write_event(Event::Text(BytesText::new(&v)))?;
                        writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
                    }
                    _ => {
                        let idx = super::shared_string_index(doc, meta, value, shared_lookup);
                        writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                        writer.write_event(Event::Text(BytesText::new(&idx.to_string())))?;
                        writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
                    }
                }
            }
            value @ CellValue::Image(image) => {
                if let Some(s) = image.alt_text.as_deref().filter(|s| !s.is_empty()) {
                    match &value_kind {
                        CellValueKind::SharedString { .. } => {
                            let idx = super::shared_string_index(doc, meta, value, shared_lookup);
                            writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                            writer.write_event(Event::Text(BytesText::new(&idx.to_string())))?;
                            writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
                        }
                        CellValueKind::InlineString => {
                            if let Some(is_events) = preserved_inline_is.take() {
                                for ev in is_events {
                                    writer.write_event(ev)?;
                                }
                            } else {
                                writer.write_event(Event::Start(BytesStart::new(is_tag)))?;
                                let mut t_start = BytesStart::new(t_tag);
                                if super::needs_space_preserve(s) {
                                    t_start.push_attribute(("xml:space", "preserve"));
                                }
                                writer.write_event(Event::Start(t_start))?;
                                writer.write_event(Event::Text(BytesText::new(s)))?;
                                writer.write_event(Event::End(BytesEnd::new(t_tag)))?;
                                if let Some(phonetic) = cell.phonetic.as_deref() {
                                    let base_len = s.chars().count().to_string();
                                    let mut rph_start = BytesStart::new(r_ph_tag);
                                    rph_start.push_attribute(("sb", "0"));
                                    rph_start.push_attribute(("eb", base_len.as_str()));
                                    writer.write_event(Event::Start(rph_start))?;

                                    let mut ph_t_start = BytesStart::new(t_tag);
                                    if super::needs_space_preserve(phonetic) {
                                        ph_t_start.push_attribute(("xml:space", "preserve"));
                                    }
                                    writer.write_event(Event::Start(ph_t_start))?;
                                    writer.write_event(Event::Text(BytesText::new(phonetic)))?;
                                    writer.write_event(Event::End(BytesEnd::new(t_tag)))?;
                                    writer.write_event(Event::End(BytesEnd::new(r_ph_tag)))?;
                                }
                                writer.write_event(Event::End(BytesEnd::new(is_tag)))?;
                            }
                        }
                        CellValueKind::Str => {
                            writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                            let v = super::raw_or_str(meta, s);
                            writer.write_event(Event::Text(BytesText::new(&v)))?;
                            writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
                        }
                        _ => {
                            let idx = super::shared_string_index(doc, meta, value, shared_lookup);
                            writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                            writer.write_event(Event::Text(BytesText::new(&idx.to_string())))?;
                            writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
                        }
                    }
                }
            }
            value @ CellValue::RichText(rich) => {
                if let Some(is_events) = preserved_inline_is.take() {
                    for ev in is_events {
                        writer.write_event(ev)?;
                    }
                } else {
                    let idx = super::shared_string_index(doc, meta, value, shared_lookup);
                    if idx != 0 || !rich.text.is_empty() {
                        writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
                        writer.write_event(Event::Text(BytesText::new(&idx.to_string())))?;
                        writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
                    }
                }
            }
            _ => {
                // Array/Spill not yet modeled for writing. Preserve as blank.
            }
        }
    }

    for ev in preserved_children {
        writer.write_event(ev)?;
    }

    writer.write_event(Event::End(BytesEnd::new(cell_tag.as_str())))?;
    Ok(())
}

fn formula_meta_for_cell(
    meta: Option<&crate::CellMeta>,
    cell: &formula_model::Cell,
) -> Option<crate::FormulaMeta> {
    let model_formula = cell.formula.as_deref();
    match (model_formula, meta.and_then(|m| m.formula.clone())) {
        (Some(_), Some(meta)) => Some(meta),
        (Some(formula), None) => Some(crate::FormulaMeta {
            file_text: crate::formula_text::add_xlfn_prefixes(super::strip_leading_equals(formula)),
            ..Default::default()
        }),
        (None, Some(meta)) => {
            // Keep follower shared formulas (represented in SpreadsheetML metadata but not modeled
            // in memory), but drop stale formula text when the model clears it.
            if meta.file_text.is_empty()
                && meta.t.is_none()
                && meta.reference.is_none()
                && meta.shared_index.is_none()
                && meta.always_calc.is_none()
            {
                None
            } else if meta.file_text.is_empty() {
                Some(meta)
            } else {
                None
            }
        }
        (None, None) => None,
    }
}

fn row_num_from_attrs(e: &BytesStart<'_>) -> Result<u32, WriteError> {
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"r" {
            return Ok(attr.unescape_value()?.trim().parse().unwrap_or(0));
        }
    }
    Ok(0)
}

fn row_has_unknown_attrs(e: &BytesStart<'_>) -> Result<bool, WriteError> {
    for attr in e.attributes() {
        let attr = attr?;
        // Attributes we manage from the worksheet model. These should not force retention of an
        // otherwise-empty row once we rewrite the `<row>` start tag to match the model.
        //
        // Anything else (including `spans` and `x14ac:*`) is treated as "unknown" and preserved
        // byte-for-byte, so we must keep the `<row>` element around.
        match attr.key.as_ref() {
            b"r" | b"ht" | b"customHeight" | b"hidden" | b"s" | b"customFormat" => {}
            _ => return Ok(true),
        }
    }
    Ok(false)
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
struct RowSemantics {
    height: Option<f32>,
    hidden: bool,
    /// Optional style xf index, gated by `customFormat="1"` semantics.
    style_xf: Option<u32>,
}

fn row_semantics_from_model(
    sheet: &Worksheet,
    style_to_xf: &HashMap<u32, u32>,
    row_num: u32,
) -> RowSemantics {
    let props = sheet.row_properties(row_num.saturating_sub(1));
    let height = props.and_then(|p| p.height);
    let hidden = props.map(|p| p.hidden).unwrap_or(false);
    let style_xf = props
        .and_then(|p| p.style_id)
        .map(|style_id| {
            if style_id == 0 {
                0
            } else {
                style_to_xf.get(&style_id).copied().unwrap_or(0)
            }
        });

    RowSemantics {
        height,
        hidden,
        style_xf,
    }
}

fn row_semantics_from_xml(e: &BytesStart<'_>) -> Result<RowSemantics, WriteError> {
    let mut ht: Option<f32> = None;
    let mut custom_height: Option<bool> = None;
    let mut hidden = false;
    let mut s: Option<u32> = None;
    let mut custom_format = false;

    for attr in e.attributes() {
        let attr = attr?;
        let val = attr.unescape_value()?.into_owned();
        match attr.key.as_ref() {
            b"ht" => ht = val.parse::<f32>().ok(),
            b"customHeight" => custom_height = Some(super::parse_xml_bool(&val)),
            b"hidden" => hidden = super::parse_xml_bool(&val),
            b"s" => s = val.parse::<u32>().ok(),
            b"customFormat" => custom_format = super::parse_xml_bool(&val),
            _ => {}
        }
    }

    let height = if custom_height == Some(false) { None } else { ht };
    let style_xf = if custom_format { Some(s.unwrap_or(0)) } else { None };

    Ok(RowSemantics {
        height,
        hidden,
        style_xf,
    })
}

fn is_xml_whitespace(b: u8) -> bool {
    matches!(b, b' ' | b'\n' | b'\r' | b'\t')
}

fn extract_unknown_row_attr_segments(e: &BytesStart<'_>) -> Vec<Vec<u8>> {
    // `BytesStart`'s underlying buffer contains the original start tag content:
    // `row r="1" spans="1:2" x14ac:dyDescent="0.25"`.
    //
    // We preserve *unknown* attribute segments byte-for-byte by slicing this raw buffer,
    // while we synthesize managed attributes from the model.
    //
    // If parsing fails for any reason, return an empty list (fallback is to rewrite the tag
    // without preserving unknown attrs).
    const MANAGED: [&[u8]; 6] = [b"r", b"ht", b"customHeight", b"hidden", b"s", b"customFormat"];

    let raw: &[u8] = e.as_ref();
    let name_len = e.name().as_ref().len();
    if raw.len() < name_len {
        return Vec::new();
    }

    let mut out: Vec<Vec<u8>> = Vec::new();
    let mut i = name_len;
    while i < raw.len() {
        let seg_start = i;
        while i < raw.len() && is_xml_whitespace(raw[i]) {
            i += 1;
        }
        if i >= raw.len() {
            break;
        }

        let key_start = i;
        while i < raw.len() && !is_xml_whitespace(raw[i]) && raw[i] != b'=' {
            i += 1;
        }
        let key_end = i;

        while i < raw.len() && is_xml_whitespace(raw[i]) {
            i += 1;
        }
        if i >= raw.len() || raw[i] != b'=' {
            break;
        }
        i += 1;
        while i < raw.len() && is_xml_whitespace(raw[i]) {
            i += 1;
        }
        if i >= raw.len() {
            break;
        }

        let quote = raw[i];
        if quote != b'"' && quote != b'\'' {
            break;
        }
        i += 1;
        while i < raw.len() {
            if raw[i] == quote {
                i += 1;
                break;
            }
            i += 1;
        }
        if i > raw.len() {
            break;
        }

        let key = &raw[key_start..key_end];
        if !MANAGED.iter().any(|managed| managed == &key) {
            out.push(raw[seg_start..i].to_vec());
        }
    }

    out
}

fn build_patched_row_tag_bytes(
    e: &BytesStart<'_>,
    row_num: u32,
    desired: RowSemantics,
    empty_element: bool,
) -> Vec<u8> {
    let name = e.name();
    let name = name.as_ref();

    let mut out = Vec::new();
    out.push(b'<');
    out.extend_from_slice(name);

    // Managed attrs.
    out.extend_from_slice(br#" r=""#);
    out.extend_from_slice(row_num.to_string().as_bytes());
    out.push(b'"');

    if let Some(height) = desired.height {
        out.extend_from_slice(br#" ht=""#);
        out.extend_from_slice(height.to_string().as_bytes());
        out.push(b'"');
        out.extend_from_slice(br#" customHeight="1""#);
    }

    if desired.hidden {
        out.extend_from_slice(br#" hidden="1""#);
    }

    if let Some(style_xf) = desired.style_xf {
        out.extend_from_slice(br#" s=""#);
        out.extend_from_slice(style_xf.to_string().as_bytes());
        out.push(b'"');
        out.extend_from_slice(br#" customFormat="1""#);
    }

    // Unknown attrs, preserved byte-for-byte.
    for seg in extract_unknown_row_attr_segments(e) {
        out.extend_from_slice(&seg);
    }

    if empty_element {
        out.extend_from_slice(b"/>");
    } else {
        out.push(b'>');
    }
    out
}

fn write_patched_row_tag<W: Write>(
    writer: &mut Writer<W>,
    e: BytesStart<'static>,
    sheet: &Worksheet,
    style_to_xf: &HashMap<u32, u32>,
    row_num: u32,
    empty_element: bool,
) -> Result<(), WriteError> {
    // If we can't address the row, preserve it unchanged.
    if row_num == 0 {
        if empty_element {
            writer.write_event(Event::Empty(e))?;
        } else {
            writer.write_event(Event::Start(e))?;
        }
        return Ok(());
    }

    let desired = row_semantics_from_model(sheet, style_to_xf, row_num);
    let original = row_semantics_from_xml(&e)?;

    // If no semantic change is required, preserve the original start tag verbatim.
    if desired == original {
        if empty_element {
            writer.write_event(Event::Empty(e))?;
        } else {
            writer.write_event(Event::Start(e))?;
        }
        return Ok(());
    }

    let bytes = build_patched_row_tag_bytes(&e, row_num, desired, empty_element);
    writer.get_mut().write_all(&bytes)?;
    Ok(())
}

fn peek_row(desired_cells: &[(CellRef, &formula_model::Cell)], idx: usize) -> Option<u32> {
    desired_cells.get(idx).map(|(r, _)| r.row + 1)
}

fn collect_full_element<R: std::io::BufRead>(
    start: Event<'static>,
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
) -> Result<Vec<Event<'static>>, WriteError> {
    let mut events: Vec<Event<'static>> = vec![start];
    let mut depth: usize = 1;
    loop {
        let ev = reader.read_event_into(buf)?.into_owned();
        match ev {
            Event::Start(_) => {
                depth += 1;
                events.push(ev);
            }
            Event::End(_) => {
                if depth == 0 {
                    events.push(ev);
                    break;
                }
                depth -= 1;
                events.push(ev);
                if depth == 0 {
                    break;
                }
            }
            Event::Eof => break,
            _ => events.push(ev),
        }
        buf.clear();
    }
    Ok(events)
}

fn write_events<W: Write>(
    writer: &mut Writer<W>,
    events: Vec<Event<'static>>,
) -> Result<(), WriteError> {
    for ev in events {
        writer.write_event(ev)?;
    }
    Ok(())
}

fn cell_ref_from_cell_events(
    events: &[Event<'static>],
) -> Result<Option<(CellRef, u32)>, WriteError> {
    let start = match events.first() {
        Some(Event::Start(e)) => e,
        Some(Event::Empty(e)) => e,
        _ => return Ok(None),
    };

    for attr in start.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"r" {
            let a1 = attr.unescape_value()?.into_owned();
            if let Ok(cell_ref) = CellRef::from_a1(a1.trim()) {
                return Ok(Some((cell_ref, cell_ref.col + 1)));
            }
            return Ok(None);
        }
    }
    Ok(None)
}

#[derive(Debug, Clone)]
struct CellSemantics {
    style_xf: u32,
    formula: Option<String>,
    value: CellValue,
    phonetic: Option<String>,
}

impl PartialEq for CellSemantics {
    fn eq(&self, other: &Self) -> bool {
        self.style_xf == other.style_xf
            && self.formula == other.formula
            && self.phonetic == other.phonetic
            && cell_value_semantics_eq(&self.value, &other.value)
    }
}

impl CellSemantics {
    fn from_model(cell: &formula_model::Cell, style_to_xf: &HashMap<u32, u32>) -> Self {
        let style_xf = if cell.style_id == 0 {
            0
        } else {
            style_to_xf.get(&cell.style_id).copied().unwrap_or(0)
        };
        // Compare entities/records by their scalar display representation so that simply *modeling*
        // a rich value does not force the writer to rewrite the cell XML on a no-op roundtrip.
        //
        // The `vm` attribute and richData payloads are preserved separately via `XlsxDocument.parts`.
        let value = match &cell.value {
            CellValue::Entity(entity) => CellValue::String(entity.display_value.clone()),
            CellValue::Record(record) => CellValue::String(record.to_string()),
            CellValue::Image(image) => match image.alt_text.as_deref().filter(|s| !s.is_empty()) {
                Some(alt) => CellValue::String(alt.to_string()),
                None => CellValue::Empty,
            },
            other => other.clone(),
        };
        let phonetic = match &cell.value {
            CellValue::String(_) => cell.phonetic.clone(),
            _ => None,
        };
        Self {
            style_xf,
            formula: cell
                .formula
                .as_deref()
                .map(crate::formula_text::normalize_display_formula)
                .filter(|f| !f.is_empty()),
            value,
            phonetic,
        }
    }

    fn is_truly_empty(&self) -> bool {
        self.style_xf == 0
            && self.formula.is_none()
            && matches!(self.value, CellValue::Empty)
            && self.phonetic.is_none()
    }
}

fn cell_value_semantics_eq(a: &CellValue, b: &CellValue) -> bool {
    match (a, b) {
        (CellValue::RichText(a), CellValue::RichText(b)) => rich_text_semantics_eq(a, b),
        (CellValue::String(s), CellValue::RichText(rich))
        | (CellValue::RichText(rich), CellValue::String(s)) => {
            &rich.text == s && rich_text_has_no_formatting(rich)
        }
        _ => a == b,
    }
}

fn rich_text_has_no_formatting(rich: &RichText) -> bool {
    rich_text_normalized_runs(rich).is_empty()
}

fn rich_text_semantics_eq(a: &RichText, b: &RichText) -> bool {
    a.text == b.text && rich_text_normalized_runs(a) == rich_text_normalized_runs(b)
}

fn rich_text_normalized_runs(rich: &RichText) -> Vec<RichTextRun> {
    // Normalize away representation differences that do not affect formatting semantics:
    // - Drop empty-style runs (they carry no overrides).
    // - Drop zero-length runs (they cannot affect any characters).
    // - Merge adjacent runs that have identical styles.
    //
    // This matches the semantic equality logic used by the in-memory cell patcher so that the
    // sheet-data patcher can avoid unnecessary rewrites (and preserve unknown OOXML tags).
    let mut runs: Vec<RichTextRun> = Vec::new();
    let _ = runs.try_reserve(rich.runs.len());
    for run in &rich.runs {
        if run.start < run.end && !run.style.is_empty() {
            runs.push(run.clone());
        }
    }
    runs.sort_by_key(|run| (run.start, run.end));

    let mut merged: Vec<RichTextRun> = Vec::new();
    let _ = merged.try_reserve(runs.len());
    for run in runs {
        match merged.last_mut() {
            Some(prev) if prev.style == run.style && prev.end == run.start => {
                prev.end = run.end;
            }
            _ => merged.push(run),
        }
    }

    merged
}

fn parse_cell_semantics(
    _doc: &XlsxDocument,
    _sheet_meta: &SheetMeta,
    events: &[Event<'static>],
    shared_strings: &[RichText],
) -> Result<CellSemantics, WriteError> {
    let start = match events.first() {
        Some(Event::Start(e)) => e,
        Some(Event::Empty(e)) => e,
        _ => {
            return Ok(CellSemantics {
                style_xf: 0,
                formula: None,
                value: CellValue::Empty,
                phonetic: None,
            })
        }
    };

    let mut style_xf: u32 = 0;
    let mut cell_type: Option<String> = None;
    for attr in start.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"s" => style_xf = attr.unescape_value()?.trim().parse().unwrap_or(0),
            b"t" => cell_type = Some(attr.unescape_value()?.into_owned()),
            _ => {}
        }
    }

    if matches!(events.first(), Some(Event::Empty(_))) {
        return Ok(CellSemantics {
            style_xf,
            formula: None,
            value: CellValue::Empty,
            phonetic: None,
        });
    }

    let mut v_text: Option<String> = None;
    let mut f_text: Option<String> = None;
    let mut inline_text: Option<String> = None;
    let mut phonetic_text: Option<String> = None;
    let mut in_v = false;
    let mut in_f = false;
    let mut in_inline_t = false;
    let mut in_phonetic_t = false;

    let is_inline_str = cell_type.as_deref() == Some("inlineStr");

    // Track the local-name path so we only treat `<t>` as visible text when:
    // - `<is><t>...` (direct child of `<is>`), or
    // - `<is><r>...<t>` (rich text runs).
    //
    // This intentionally ignores `<rPh><t>` phonetic runs and any other `<t>` elements in
    // non-visible subtrees.
    let mut tag_stack: Vec<Vec<u8>> = Vec::new();

    for ev in events {
        match ev {
            Event::Start(e) => {
                let name = e.name();
                let local = super::local_name(name.as_ref());
                match local {
                    b"v" => in_v = true,
                    b"f" => in_f = true,
                    b"rPh" if is_inline_str => {
                        // Presence of `<rPh>` implies phonetic metadata even if it contains no text.
                        phonetic_text.get_or_insert_with(String::new);
                    }
                    b"t" if is_inline_str => {
                        if tag_stack.iter().any(|n| n.as_slice() == b"rPh") {
                            in_phonetic_t = true;
                        } else {
                            let visible = match tag_stack.as_slice() {
                                // `<is><t>`
                                [.., parent] if parent.as_slice() == b"is" => true,
                                // `<is><r><t>`
                                [.., grandparent, parent]
                                    if grandparent.as_slice() == b"is" && parent.as_slice() == b"r" =>
                                {
                                    true
                                }
                                _ => false,
                            };
                            in_inline_t = visible;
                        }
                    }
                    _ => {}
                }
                tag_stack.push(local.to_vec());
            }
            Event::End(e) => {
                let name = e.name();
                let local = super::local_name(name.as_ref());
                match local {
                    b"v" => in_v = false,
                    b"f" => in_f = false,
                    b"t" if is_inline_str => {
                        in_inline_t = false;
                        in_phonetic_t = false;
                    }
                    _ => {}
                }
                tag_stack.pop();
            }
            Event::Empty(e) => {
                if super::local_name(e.name().as_ref()) == b"f" {
                    // Shared formulas may be represented as <f .../> with no text.
                    f_text.get_or_insert_with(String::new);
                } else if is_inline_str && super::local_name(e.name().as_ref()) == b"rPh" {
                    phonetic_text.get_or_insert_with(String::new);
                }
            }
            Event::Text(t) => {
                let text = t.unescape()?.into_owned();
                if in_v {
                    v_text = Some(text);
                } else if in_f {
                    match f_text.as_mut() {
                        Some(existing) => existing.push_str(&text),
                        None => f_text = Some(text),
                    }
                } else if in_inline_t {
                    match inline_text.as_mut() {
                        Some(existing) => existing.push_str(&text),
                        None => inline_text = Some(text),
                    }
                } else if in_phonetic_t {
                    match phonetic_text.as_mut() {
                        Some(existing) => existing.push_str(&text),
                        None => phonetic_text = Some(text),
                    }
                }
            }
            Event::CData(c) => {
                let text = String::from_utf8_lossy(c.as_ref()).into_owned();
                if in_v {
                    v_text = Some(text);
                } else if in_f {
                    match f_text.as_mut() {
                        Some(existing) => existing.push_str(&text),
                        None => f_text = Some(text),
                    }
                } else if in_inline_t {
                    match inline_text.as_mut() {
                        Some(existing) => existing.push_str(&text),
                        None => inline_text = Some(text),
                    }
                } else if in_phonetic_t {
                    match phonetic_text.as_mut() {
                        Some(existing) => existing.push_str(&text),
                        None => phonetic_text = Some(text),
                    }
                }
            }
            _ => {}
        }
    }

    let inline_value = if cell_type.as_deref() == Some("inlineStr") {
        parse_inline_is_cell_value(events)?
    } else {
        None
    };
    let value = interpret_cell_value(cell_type.as_deref(), &v_text, &inline_value, shared_strings);
    let formula = f_text
        .as_deref()
        .filter(|t| !t.is_empty())
        .map(crate::formula_text::normalize_display_formula);

    Ok(CellSemantics {
        style_xf,
        formula,
        value,
        phonetic: phonetic_text,
    })
}

fn cell_type_from_cell_events(events: &[Event<'static>]) -> Result<Option<String>, WriteError> {
    let start = match events.first() {
        Some(Event::Start(e)) => e,
        Some(Event::Empty(e)) => e,
        _ => return Ok(None),
    };

    for attr in start.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"t" {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }
    Ok(None)
}

fn extract_is_subtree(events: &[Event<'static>]) -> Option<Vec<Event<'static>>> {
    let mut depth: usize = 0;

    // Walk the event stream and find an `<is>` element that is a direct child of the cell.
    for (idx, ev) in events.iter().enumerate() {
        match ev {
            Event::Start(e) => {
                let name = e.name();
                let local = super::local_name(name.as_ref());
                if depth == 1 && local == b"is" {
                    let mut out = Vec::new();
                    let mut sub_depth: usize = 0;
                    for sub_ev in &events[idx..] {
                        out.push(sub_ev.clone());
                        match sub_ev {
                            Event::Start(_) => sub_depth += 1,
                            Event::End(_) => {
                                if sub_depth == 0 {
                                    break;
                                }
                                sub_depth -= 1;
                                if sub_depth == 0 {
                                    break;
                                }
                            }
                            _ => {}
                        }
                    }
                    return Some(out);
                }
                depth += 1;
            }
            Event::Empty(e) => {
                let name = e.name();
                let local = super::local_name(name.as_ref());
                if depth == 1 && local == b"is" {
                    return Some(vec![ev.clone()]);
                }
            }
            Event::End(_) => {
                if depth > 0 {
                    depth -= 1;
                }
            }
            _ => {}
        }
    }

    None
}

fn interpret_cell_value(
    t: Option<&str>,
    v_text: &Option<String>,
    inline_value: &Option<CellValue>,
    shared_strings: &[RichText],
) -> CellValue {
    match t {
        Some("s") => {
            let raw = v_text.clone().unwrap_or_default();
            let idx: u32 = raw.parse().unwrap_or(0);
            shared_strings
                .get(idx as usize)
                .map(|rich| {
                    if rich.runs.is_empty() {
                        CellValue::String(rich.text.clone())
                    } else {
                        CellValue::RichText(rich.clone())
                    }
                })
                .unwrap_or_else(|| CellValue::String(String::new()))
        }
        Some("b") => {
            let raw = v_text.clone().unwrap_or_default();
            CellValue::Boolean(raw == "1")
        }
        Some("e") => {
            let raw = v_text.clone().unwrap_or_default();
            let err = raw.parse::<ErrorValue>().unwrap_or(ErrorValue::Unknown);
            CellValue::Error(err)
        }
        Some("str") => {
            let raw = v_text.clone().unwrap_or_default();
            CellValue::String(raw)
        }
        Some("inlineStr") => {
            inline_value
                .clone()
                .unwrap_or_else(|| CellValue::String(String::new()))
        }
        Some(_) | None => {
            if let Some(raw) = v_text.clone() {
                raw.parse::<f64>()
                    .map(CellValue::Number)
                    .unwrap_or(CellValue::String(raw))
            } else {
                CellValue::Empty
            }
        }
    }
}

fn parse_inline_is_cell_value(events: &[Event<'static>]) -> Result<Option<CellValue>, WriteError> {
    let mut idx = 0usize;
    while idx < events.len() {
        match &events[idx] {
            Event::Empty(e) if super::local_name(e.name().as_ref()) == b"is" => {
                return Ok(Some(CellValue::String(String::new())));
            }
            Event::Start(e) if super::local_name(e.name().as_ref()) == b"is" => {
                let (value, _) = parse_inline_is_at(events, idx)?;
                return Ok(Some(value));
            }
            _ => {}
        }
        idx += 1;
    }
    Ok(None)
}

fn parse_inline_is_at(
    events: &[Event<'static>],
    start_idx: usize,
) -> Result<(CellValue, usize), WriteError> {
    let mut segments: Vec<(String, RichTextRunStyle)> = Vec::new();
    let mut idx = start_idx + 1;

    while idx < events.len() {
        match &events[idx] {
            Event::Start(e) if super::local_name(e.name().as_ref()) == b"t" => {
                let (text, next) = read_text_from_events(events, idx, b"t")?;
                segments.push((text, RichTextRunStyle::default()));
                idx = next;
            }
            Event::Start(e) if super::local_name(e.name().as_ref()) == b"r" => {
                let ((text, style), next) = parse_inline_r_at(events, idx)?;
                segments.push((text, style));
                idx = next;
            }
            Event::Start(_) => {
                // Only treat `<t>` as visible text when it is a direct child of `<is>` or
                // inside `<is><r>...</r></is>`. Other subtrees (phonetic/ruby runs, extensions)
                // may contain `<t>` elements that should not be concatenated into the display
                // string.
                idx = skip_element(events, idx);
            }
            Event::End(e) if super::local_name(e.name().as_ref()) == b"is" => {
                idx += 1;
                break;
            }
            _ => idx += 1,
        }
    }

    if segments.is_empty() {
        return Ok((CellValue::String(String::new()), idx));
    }

    if segments.iter().all(|(_, style)| style.is_empty()) {
        Ok((
            CellValue::String(segments.into_iter().map(|(text, _)| text).collect()),
            idx,
        ))
    } else {
        Ok((CellValue::RichText(RichText::from_segments(segments)), idx))
    }
}

fn parse_inline_r_at(
    events: &[Event<'static>],
    start_idx: usize,
) -> Result<((String, RichTextRunStyle), usize), WriteError> {
    let mut style = RichTextRunStyle::default();
    let mut text = String::new();
    let mut idx = start_idx + 1;

    while idx < events.len() {
        match &events[idx] {
            Event::Start(e) if super::local_name(e.name().as_ref()) == b"rPr" => {
                let (parsed, next) = parse_inline_rpr_at(events, idx)?;
                style = parsed;
                idx = next;
            }
            Event::Empty(e) if super::local_name(e.name().as_ref()) == b"rPr" => {
                // `<rPr/>` with no style tags.
                style = RichTextRunStyle::default();
                idx += 1;
            }
            Event::Start(e) if super::local_name(e.name().as_ref()) == b"t" => {
                let (t, next) = read_text_from_events(events, idx, b"t")?;
                text.push_str(&t);
                idx = next;
            }
            Event::Start(_) => {
                idx = skip_element(events, idx);
            }
            Event::End(e) if super::local_name(e.name().as_ref()) == b"r" => {
                idx += 1;
                break;
            }
            _ => idx += 1,
        }
    }

    Ok(((text, style), idx))
}

fn parse_inline_rpr_at(
    events: &[Event<'static>],
    start_idx: usize,
) -> Result<(RichTextRunStyle, usize), WriteError> {
    let mut style = RichTextRunStyle::default();
    let mut idx = start_idx + 1;

    while idx < events.len() {
        match &events[idx] {
            Event::Empty(e) => {
                parse_inline_rpr_tag(e, &mut style)?;
                idx += 1;
            }
            Event::Start(e) => {
                parse_inline_rpr_tag(e, &mut style)?;
                idx = skip_element(events, idx);
            }
            Event::End(e) if super::local_name(e.name().as_ref()) == b"rPr" => {
                idx += 1;
                break;
            }
            _ => idx += 1,
        }
    }

    Ok((style, idx))
}

fn parse_inline_rpr_tag(e: &BytesStart<'_>, style: &mut RichTextRunStyle) -> Result<(), WriteError> {
    match super::local_name(e.name().as_ref()) {
        b"b" => style.bold = Some(parse_inline_rpr_bool_val(e)?),
        b"i" => style.italic = Some(parse_inline_rpr_bool_val(e)?),
        b"u" => {
            let val = attr_value(e, b"val")?;
            if let Some(ul) = Underline::from_ooxml(val.as_deref()) {
                style.underline = Some(ul);
            }
        }
        b"color" => {
            if let Some(rgb) = attr_value(e, b"rgb")? {
                if rgb.len() == 8 {
                    if let Ok(argb) = u32::from_str_radix(&rgb, 16) {
                        style.color = Some(Color::new_argb(argb));
                    }
                }
            }
        }
        b"rFont" | b"name" => {
            if let Some(val) = attr_value(e, b"val")? {
                style.font = Some(val);
            }
        }
        b"sz" => {
            if let Some(val) = attr_value(e, b"val")? {
                if let Some(sz) = parse_size_100pt(&val) {
                    style.size_100pt = Some(sz);
                }
            }
        }
        _ => {}
    }
    Ok(())
}

fn parse_inline_rpr_bool_val(e: &BytesStart<'_>) -> Result<bool, WriteError> {
    let Some(val) = attr_value(e, b"val")? else {
        return Ok(true);
    };
    Ok(!(val == "0" || val.eq_ignore_ascii_case("false")))
}

fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Result<Option<String>, WriteError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        if attr.key.as_ref() == key {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }
    Ok(None)
}

fn parse_size_100pt(val: &str) -> Option<u16> {
    let val = val.trim();
    if val.is_empty() {
        return None;
    }

    if let Some((int_part, frac_part)) = val.split_once('.') {
        let int: u16 = int_part.parse().ok()?;
        let mut frac = frac_part.chars().take(2).collect::<String>();
        while frac.len() < 2 {
            frac.push('0');
        }
        let frac: u16 = frac.parse().ok()?;
        int.checked_mul(100)?.checked_add(frac)
    } else {
        let int: u16 = val.parse().ok()?;
        int.checked_mul(100)
    }
}

fn read_text_from_events(
    events: &[Event<'static>],
    start_idx: usize,
    end_local: &[u8],
) -> Result<(String, usize), WriteError> {
    let mut text = String::new();
    let mut idx = start_idx + 1;
    while idx < events.len() {
        match &events[idx] {
            Event::Text(t) => {
                text.push_str(&t.unescape()?.into_owned());
            }
            Event::CData(c) => {
                text.push_str(&String::from_utf8_lossy(c.as_ref()));
            }
            Event::End(e) if super::local_name(e.name().as_ref()) == end_local => {
                idx += 1;
                break;
            }
            _ => {}
        }
        idx += 1;
    }
    Ok((text, idx))
}

fn skip_element(events: &[Event<'static>], start_idx: usize) -> usize {
    if !matches!(events.get(start_idx), Some(Event::Start(_))) {
        return start_idx + 1;
    }

    let mut depth = 1usize;
    let mut idx = start_idx + 1;
    while idx < events.len() {
        match &events[idx] {
            Event::Start(_) => depth += 1,
            Event::End(_) => {
                depth -= 1;
                if depth == 0 {
                    return idx + 1;
                }
            }
            _ => {}
        }
        idx += 1;
    }
    events.len()
}

fn extract_preserved_cell_children(events: &[Event<'static>]) -> Vec<Event<'static>> {
    if events.len() < 2 {
        return Vec::new();
    }

    let mut preserved = Vec::new();
    let mut skipping: usize = 0;

    for ev in events.iter().skip(1).take(events.len() - 2) {
        if skipping > 0 {
            match ev {
                Event::Start(_) => skipping += 1,
                Event::End(_) => skipping -= 1,
                _ => {}
            }
            continue;
        }

        match ev {
            Event::Start(e)
                if matches!(super::local_name(e.name().as_ref()), b"f" | b"v" | b"is") =>
            {
                skipping = 1;
            }
            Event::Empty(e)
                if matches!(super::local_name(e.name().as_ref()), b"f" | b"v" | b"is") => {}
            other => preserved.push(other.clone()),
        }
    }

    preserved
}
