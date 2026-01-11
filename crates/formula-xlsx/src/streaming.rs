use std::collections::{BTreeMap, HashMap};
use std::io::{BufRead, BufReader, Read, Seek, Write};

use formula_model::{CellRef, CellValue};
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};
use thiserror::Error;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

#[derive(Debug, Error)]
pub enum StreamingPatchError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml attribute error: {0}")]
    XmlAttr(#[from] quick_xml::events::attributes::AttrError),
    #[error("unsupported cell value kind for streaming patch: {0:?}")]
    UnsupportedCellValue(CellValue),
    #[error("invalid cell reference in worksheet xml: {0}")]
    InvalidCellRef(String),
    #[error("worksheet part referenced by patch not found in input zip: {0}")]
    MissingWorksheetPart(String),
    #[error("worksheet xml is missing required <sheetData> section: {0}")]
    MissingSheetData(String),
}

/// Patch a single cell in a worksheet part (`xl/worksheets/sheetN.xml`).
///
/// The patch is applied by rewriting only the worksheet XML part. All other ZIP entries are
/// copied from the input XLSX/XLSM streamingly.
#[derive(Debug, Clone)]
pub struct WorksheetCellPatch {
    /// ZIP entry name for the worksheet XML (e.g. `xl/worksheets/sheet1.xml`).
    pub worksheet_part: String,
    /// Cell reference to patch.
    pub cell: CellRef,
    /// New cached value for the cell.
    pub value: CellValue,
    /// Optional formula to write into the `<f>` element. Leading `=` is permitted.
    pub formula: Option<String>,
}

impl WorksheetCellPatch {
    pub fn new(
        worksheet_part: impl Into<String>,
        cell: CellRef,
        value: CellValue,
        formula: Option<String>,
    ) -> Self {
        Self {
            worksheet_part: worksheet_part.into(),
            cell,
            value,
            formula,
        }
    }
}

/// Streaming XLSX/XLSM patcher.
///
/// This function reads the input ZIP once and writes a new ZIP by stream-copying all unchanged
/// parts. Only worksheet XML parts mentioned in `cell_patches` are rewritten.
pub fn patch_xlsx_streaming<R: Read + Seek, W: Write + Seek>(
    input: R,
    output: W,
    cell_patches: &[WorksheetCellPatch],
) -> Result<(), StreamingPatchError> {
    let mut patches_by_part: HashMap<String, Vec<WorksheetCellPatch>> = HashMap::new();
    for patch in cell_patches {
        patches_by_part
            .entry(patch.worksheet_part.clone())
            .or_default()
            .push(patch.clone());
    }
    // Deterministic patching within a worksheet.
    for patches in patches_by_part.values_mut() {
        patches.sort_by_key(|p| (p.cell.row, p.cell.col));
    }

    let mut missing_parts: BTreeMap<String, ()> = patches_by_part
        .keys()
        .map(|k| (k.clone(), ()))
        .collect();

    let mut archive = ZipArchive::new(input)?;
    let mut zip = ZipWriter::new(output);
    let options =
        FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        if file.is_dir() {
            continue;
        }

        let name = file.name().to_string();
        zip.start_file(name.clone(), options)?;

        if let Some(patches) = patches_by_part.get(&name) {
            missing_parts.remove(&name);
            patch_worksheet_xml_streaming(&mut file, &mut zip, &name, patches)?;
        } else {
            std::io::copy(&mut file, &mut zip)?;
        }
    }

    if let Some((missing, _)) = missing_parts.into_iter().next() {
        return Err(StreamingPatchError::MissingWorksheetPart(missing));
    }

    zip.finish()?;
    Ok(())
}

#[derive(Debug, Clone)]
struct CellPatch {
    row_1: u32,
    col_0: u32,
    value: CellValue,
    formula: Option<String>,
}

struct RowState {
    row_1: u32,
    pending: Vec<CellPatch>,
    next_idx: usize,
}

fn patch_worksheet_xml_streaming<R: Read, W: Write>(
    input: R,
    output: W,
    worksheet_part: &str,
    patches: &[WorksheetCellPatch],
) -> Result<(), StreamingPatchError> {
    let patch_bounds = bounds_for_patches(patches);

    let mut patches_by_row: BTreeMap<u32, Vec<CellPatch>> = BTreeMap::new();
    for patch in patches {
        let row_1 = patch.cell.row + 1;
        let col_0 = patch.cell.col;
        patches_by_row
            .entry(row_1)
            .or_default()
            .push(CellPatch {
                row_1,
                col_0,
                value: patch.value.clone(),
                formula: patch.formula.clone(),
            });
    }
    for row_patches in patches_by_row.values_mut() {
        row_patches.sort_by_key(|p| p.col_0);
    }

    let mut reader = Reader::from_reader(BufReader::new(input));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(output);

    let mut buf = Vec::new();
    let mut in_sheet_data = false;
    let mut saw_sheet_data = false;
    let mut patched_dimension = false;

    let mut row_state: Option<RowState> = None;
    let mut in_cell = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,

            Event::Start(ref e) if e.name().as_ref() == b"sheetData" => {
                saw_sheet_data = true;
                in_sheet_data = true;
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if e.name().as_ref() == b"sheetData" => {
                saw_sheet_data = true;
                if patches_by_row.is_empty() {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                } else {
                    in_sheet_data = false;
                    // Expand `<sheetData/>` into `<sheetData>...</sheetData>`.
                    writer.write_event(Event::Start(e.to_owned()))?;
                    write_pending_rows(&mut writer, &mut patches_by_row)?;
                    writer.write_event(Event::End(BytesEnd::new("sheetData")))?;
                }
            }
            Event::End(ref e) if e.name().as_ref() == b"sheetData" => {
                // Flush any remaining patch rows at the end of sheetData.
                write_pending_rows(&mut writer, &mut patches_by_row)?;
                in_sheet_data = false;
                writer.write_event(Event::End(e.to_owned()))?;
            }

            Event::Start(ref e) if e.name().as_ref() == b"dimension" => {
                if !patched_dimension {
                    patched_dimension = true;
                    if let Some(bounds) = patch_bounds {
                        let updated = updated_dimension_element(e, bounds)?;
                        writer.write_event(Event::Start(updated))?;
                    } else {
                        writer.write_event(Event::Start(e.to_owned()))?;
                    }
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
            }
            Event::Empty(ref e) if e.name().as_ref() == b"dimension" => {
                if !patched_dimension {
                    patched_dimension = true;
                    if let Some(bounds) = patch_bounds {
                        let updated = updated_dimension_element(e, bounds)?;
                        writer.write_event(Event::Empty(updated))?;
                    } else {
                        writer.write_event(Event::Empty(e.to_owned()))?;
                    }
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }

            Event::Start(ref e) if in_sheet_data && e.name().as_ref() == b"row" => {
                let row_1 = parse_row_number(e)?;
                in_cell = false;

                // Insert any patch rows that should appear before this row.
                while let Some((&next_row, _)) = patches_by_row.iter().next() {
                    if next_row < row_1 {
                        let pending = patches_by_row.remove(&next_row).unwrap_or_default();
                        write_inserted_row(&mut writer, next_row, &pending)?;
                    } else {
                        break;
                    }
                }

                let pending = patches_by_row.remove(&row_1);
                if let Some(mut pending) = pending {
                    pending.sort_by_key(|p| p.col_0);
                    row_state = Some(RowState {
                        row_1,
                        pending,
                        next_idx: 0,
                    });
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if in_sheet_data && e.name().as_ref() == b"row" => {
                let row_1 = parse_row_number(e)?;
                in_cell = false;

                // Insert patch rows that should appear before this row.
                while let Some((&next_row, _)) = patches_by_row.iter().next() {
                    if next_row < row_1 {
                        let pending = patches_by_row.remove(&next_row).unwrap_or_default();
                        write_inserted_row(&mut writer, next_row, &pending)?;
                    } else {
                        break;
                    }
                }

                if let Some(mut pending) = patches_by_row.remove(&row_1) {
                    pending.sort_by_key(|p| p.col_0);
                    // Expand `<row/>` into `<row>...</row>` and insert cells.
                    writer.write_event(Event::Start(e.to_owned()))?;
                    write_inserted_cells(&mut writer, &pending)?;
                    writer.write_event(Event::End(BytesEnd::new("row")))?;
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::End(ref e) if in_sheet_data && e.name().as_ref() == b"row" => {
                if let Some(state) = row_state.take() {
                    write_remaining_row_cells(&mut writer, &state.pending, state.next_idx)?;
                }
                in_cell = false;
                writer.write_event(Event::End(e.to_owned()))?;
            }

            // Inside a row that needs patching, intercept cell events.
            Event::Start(ref e) if in_sheet_data && row_state.is_some() && e.name().as_ref() == b"c" => {
                let state = row_state.as_mut().expect("row_state just checked");
                let (cell_ref, col_0) = parse_cell_ref_and_col(e)?;

                // Insert any pending patches that come before this cell.
                insert_pending_before_cell(&mut writer, state, col_0)?;

                if let Some(patch) = take_patch_for_col(state, col_0) {
                    patch_existing_cell(&mut reader, &mut writer, e, &cell_ref, &patch)?;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                    in_cell = true;
                }
            }
            Event::Empty(ref e) if in_sheet_data && row_state.is_some() && e.name().as_ref() == b"c" => {
                let state = row_state.as_mut().expect("row_state just checked");
                let (cell_ref, col_0) = parse_cell_ref_and_col(e)?;

                insert_pending_before_cell(&mut writer, state, col_0)?;

                if let Some(patch) = take_patch_for_col(state, col_0) {
                    write_patched_cell(&mut writer, Some(e), &cell_ref, &patch)?;
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::End(ref e) if in_sheet_data && row_state.is_some() && in_cell && e.name().as_ref() == b"c" => {
                in_cell = false;
                writer.write_event(Event::End(e.to_owned()))?;
            }
            // Ensure cells are emitted before any non-cell elements (e.g. extLst) in the row.
            Event::Start(ref e)
                if in_sheet_data && row_state.is_some() && !in_cell && e.name().as_ref() != b"c" =>
            {
                let state = row_state.as_mut().expect("row_state just checked");
                insert_pending_before_non_cell(&mut writer, state)?;
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e)
                if in_sheet_data && row_state.is_some() && !in_cell && e.name().as_ref() != b"c" =>
            {
                let state = row_state.as_mut().expect("row_state just checked");
                insert_pending_before_non_cell(&mut writer, state)?;
                writer.write_event(Event::Empty(e.to_owned()))?;
            }

            // Default passthrough.
            ev => writer.write_event(ev.into_owned())?,
        }

        buf.clear();
    }

    if !saw_sheet_data {
        return Err(StreamingPatchError::MissingSheetData(worksheet_part.to_string()));
    }

    Ok(())
}

fn parse_row_number(e: &BytesStart<'_>) -> Result<u32, StreamingPatchError> {
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"r" {
            let v = attr.unescape_value()?.into_owned();
            return Ok(v.parse::<u32>().unwrap_or(0));
        }
    }
    Ok(0)
}

fn parse_cell_ref_and_col(e: &BytesStart<'_>) -> Result<(CellRef, u32), StreamingPatchError> {
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"r" {
            let a1 = attr.unescape_value()?.into_owned();
            let cell_ref =
                CellRef::from_a1(&a1).map_err(|_| StreamingPatchError::InvalidCellRef(a1))?;
            return Ok((cell_ref, cell_ref.col));
        }
    }
    // Malformed cell - treat as A1 so it at least serializes.
    Ok((CellRef::new(0, 0), 0))
}

fn write_pending_rows<W: Write>(
    writer: &mut Writer<W>,
    patches_by_row: &mut BTreeMap<u32, Vec<CellPatch>>,
) -> Result<(), StreamingPatchError> {
    while let Some((&row_1, _)) = patches_by_row.iter().next() {
        let pending = patches_by_row.remove(&row_1).unwrap_or_default();
        write_inserted_row(writer, row_1, &pending)?;
    }
    Ok(())
}

fn write_inserted_row<W: Write>(
    writer: &mut Writer<W>,
    row_1: u32,
    patches: &[CellPatch],
) -> Result<(), StreamingPatchError> {
    let mut row = BytesStart::new("row");
    let row_num = row_1.to_string();
    row.push_attribute(("r", row_num.as_str()));
    writer.write_event(Event::Start(row))?;
    write_inserted_cells(writer, patches)?;
    writer.write_event(Event::End(BytesEnd::new("row")))?;
    Ok(())
}

fn write_inserted_cells<W: Write>(
    writer: &mut Writer<W>,
    patches: &[CellPatch],
) -> Result<(), StreamingPatchError> {
    for patch in patches {
        let cell_ref = CellRef::new(patch.row_1 - 1, patch.col_0);
        write_patched_cell::<W>(writer, None, &cell_ref, patch)?;
    }
    Ok(())
}

fn write_remaining_row_cells<W: Write>(
    writer: &mut Writer<W>,
    pending: &[CellPatch],
    next_idx: usize,
) -> Result<(), StreamingPatchError> {
    if next_idx >= pending.len() {
        return Ok(());
    }
    for patch in &pending[next_idx..] {
        let cell_ref = CellRef::new(patch.row_1 - 1, patch.col_0);
        write_patched_cell::<W>(writer, None, &cell_ref, patch)?;
    }
    Ok(())
}

fn insert_pending_before_cell<W: Write>(
    writer: &mut Writer<W>,
    state: &mut RowState,
    col_0: u32,
) -> Result<(), StreamingPatchError> {
    while let Some(patch) = state.pending.get(state.next_idx) {
        if patch.col_0 < col_0 {
            let cell_ref = CellRef::new(state.row_1 - 1, patch.col_0);
            write_patched_cell::<W>(writer, None, &cell_ref, patch)?;
            state.next_idx += 1;
        } else {
            break;
        }
    }
    Ok(())
}

fn insert_pending_before_non_cell<W: Write>(
    writer: &mut Writer<W>,
    state: &mut RowState,
) -> Result<(), StreamingPatchError> {
    if state.next_idx >= state.pending.len() {
        return Ok(());
    }
    for patch in &state.pending[state.next_idx..] {
        let cell_ref = CellRef::new(state.row_1 - 1, patch.col_0);
        write_patched_cell::<W>(writer, None, &cell_ref, patch)?;
    }
    state.next_idx = state.pending.len();
    Ok(())
}

fn take_patch_for_col(state: &mut RowState, col_0: u32) -> Option<CellPatch> {
    if state.next_idx >= state.pending.len() {
        return None;
    }
    let patch = state.pending.get(state.next_idx)?;
    if patch.col_0 == col_0 {
        let taken = patch.clone();
        state.next_idx += 1;
        Some(taken)
    } else {
        None
    }
}

fn patch_existing_cell<R: BufRead, W: Write>(
    reader: &mut Reader<R>,
    writer: &mut Writer<W>,
    cell_start: &BytesStart<'_>,
    cell_ref: &CellRef,
    patch: &CellPatch,
) -> Result<(), StreamingPatchError> {
    let (cell_t, body_kind) = cell_representation(&patch.value, patch.formula.as_deref())?;
    let patch_formula = patch.formula.as_deref();

    let mut c = BytesStart::new("c");
    let mut has_r = false;
    for attr in cell_start.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"t" {
            continue;
        }
        if attr.key.as_ref() == b"r" {
            has_r = true;
        }
        c.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
    }
    if !has_r {
        let a1 = cell_ref.to_a1();
        c.push_attribute(("r", a1.as_str()));
    }
    if let Some(t) = cell_t {
        c.push_attribute(("t", t));
    }

    writer.write_event(Event::Start(c))?;

    let mut inner_buf = Vec::new();
    let mut inner_events: Vec<Event<'static>> = Vec::new();
    loop {
        let ev = reader.read_event_into(&mut inner_buf)?;
        match ev {
            Event::End(ref e) if e.name().as_ref() == b"c" => break,
            Event::Eof => break,
            ev => inner_events.push(ev.into_owned()),
        }
        inner_buf.clear();
    }

    write_patched_cell_children(writer, &inner_events, patch_formula, &body_kind)?;
    writer.write_event(Event::End(BytesEnd::new("c")))?;
    Ok(())
}

fn write_patched_cell_children<W: Write>(
    writer: &mut Writer<W>,
    inner_events: &[Event<'static>],
    patch_formula: Option<&str>,
    body_kind: &CellBodyKind,
) -> Result<(), StreamingPatchError> {
    let mut formula_written = patch_formula.is_none();
    let mut value_written = matches!(body_kind, CellBodyKind::None);
    let mut saw_formula = false;
    let mut saw_value = false;

    let mut idx = 0usize;
    while idx < inner_events.len() {
        match &inner_events[idx] {
            Event::Start(e) if e.name().as_ref() == b"f" => {
                saw_formula = true;
                if !formula_written {
                    if let Some(formula) = patch_formula {
                        let detach_shared = should_detach_shared_formula(e, formula);
                        write_formula_element(writer, Some(e), formula, detach_shared)?;
                        formula_written = true;
                    }
                }
                idx = skip_owned_subtree(inner_events, idx);
                continue;
            }
            Event::Empty(e) if e.name().as_ref() == b"f" => {
                saw_formula = true;
                if !formula_written {
                    if let Some(formula) = patch_formula {
                        let detach_shared = should_detach_shared_formula(e, formula);
                        write_formula_element(writer, Some(e), formula, detach_shared)?;
                        formula_written = true;
                    }
                }
                idx += 1;
                continue;
            }
            Event::Start(e)
                if e.name().as_ref() == b"v" || e.name().as_ref() == b"is" =>
            {
                saw_value = true;

                if !formula_written {
                    if let Some(formula) = patch_formula {
                        // Original cell has no <f> before the value; insert one.
                        write_formula_element(writer, None, formula, false)?;
                        formula_written = true;
                    }
                }
                if !value_written {
                    write_value_element(writer, body_kind)?;
                    value_written = true;
                }

                idx = skip_owned_subtree(inner_events, idx);
                continue;
            }
            Event::Empty(e) if e.name().as_ref() == b"v" || e.name().as_ref() == b"is" => {
                saw_value = true;

                if !formula_written {
                    if let Some(formula) = patch_formula {
                        write_formula_element(writer, None, formula, false)?;
                        formula_written = true;
                    }
                }
                if !value_written {
                    write_value_element(writer, body_kind)?;
                    value_written = true;
                }

                idx += 1;
                continue;
            }
            ev => {
                if !formula_written && !saw_formula {
                    if let Some(formula) = patch_formula {
                        write_formula_element(writer, None, formula, false)?;
                        formula_written = true;
                    }
                }
                if !value_written && !saw_value {
                    write_value_element(writer, body_kind)?;
                    value_written = true;
                }
                writer.write_event(ev.clone())?;
            }
        }
        idx += 1;
    }

    if !formula_written {
        if let Some(formula) = patch_formula {
            write_formula_element(writer, None, formula, false)?;
        }
    }
    if !value_written {
        write_value_element(writer, body_kind)?;
    }

    Ok(())
}

fn skip_owned_subtree(events: &[Event<'static>], mut idx: usize) -> usize {
    match &events[idx] {
        Event::Start(_) => {
            let mut depth = 1usize;
            idx += 1;
            while idx < events.len() {
                match &events[idx] {
                    Event::Start(_) => depth += 1,
                    Event::End(_) => {
                        depth = depth.saturating_sub(1);
                        if depth == 0 {
                            idx += 1;
                            break;
                        }
                    }
                    Event::Empty(_) => {}
                    _ => {}
                }
                idx += 1;
            }
            idx
        }
        _ => idx + 1,
    }
}

fn write_formula_element<W: Write>(
    writer: &mut Writer<W>,
    original: Option<&BytesStart<'_>>,
    formula: &str,
    detach_shared: bool,
) -> Result<(), StreamingPatchError> {
    let formula_display = crate::formula_text::normalize_display_formula(formula);
    let file_formula = crate::formula_text::add_xlfn_prefixes(&formula_display);

    let mut f = BytesStart::new("f");
    if let Some(orig) = original {
        for attr in orig.attributes() {
            let attr = attr?;
            if detach_shared
                && matches!(attr.key.as_ref(), b"t" | b"ref" | b"si")
            {
                continue;
            }
            f.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
        }
    }

    if file_formula.is_empty() {
        writer.write_event(Event::Empty(f))?;
    } else {
        writer.write_event(Event::Start(f))?;
        writer.write_event(Event::Text(BytesText::new(&file_formula)))?;
        writer.write_event(Event::End(BytesEnd::new("f")))?;
    }
    Ok(())
}

fn write_value_element<W: Write>(
    writer: &mut Writer<W>,
    body_kind: &CellBodyKind,
) -> Result<(), StreamingPatchError> {
    match body_kind {
        CellBodyKind::V(text) => {
            writer.write_event(Event::Start(BytesStart::new("v")))?;
            writer.write_event(Event::Text(BytesText::new(text)))?;
            writer.write_event(Event::End(BytesEnd::new("v")))?;
        }
        CellBodyKind::InlineStr(text) => {
            writer.write_event(Event::Start(BytesStart::new("is")))?;
            let mut t = BytesStart::new("t");
            if needs_space_preserve(text) {
                t.push_attribute(("xml:space", "preserve"));
            }
            writer.write_event(Event::Start(t))?;
            writer.write_event(Event::Text(BytesText::new(text)))?;
            writer.write_event(Event::End(BytesEnd::new("t")))?;
            writer.write_event(Event::End(BytesEnd::new("is")))?;
        }
        CellBodyKind::None => {}
    }

    Ok(())
}

fn should_detach_shared_formula(f: &BytesStart<'_>, patch_formula: &str) -> bool {
    let patch_formula = patch_formula.strip_prefix('=').unwrap_or(patch_formula);
    if patch_formula.is_empty() {
        return false;
    }

    let mut is_shared = false;
    let mut has_ref = false;
    for attr in f.attributes().flatten() {
        match attr.key.as_ref() {
            b"t" if attr.value.as_ref() == b"shared" => is_shared = true,
            b"ref" => has_ref = true,
            _ => {}
        }
    }

    is_shared && !has_ref
}

fn write_patched_cell<W: Write>(
    writer: &mut Writer<W>,
    original: Option<&BytesStart<'_>>,
    cell_ref: &CellRef,
    patch: &CellPatch,
) -> Result<(), StreamingPatchError> {
    let (cell_t, body_kind) = cell_representation(&patch.value, patch.formula.as_deref())?;

    let mut c = BytesStart::new("c");
    let inserted_a1 = original.is_none().then(|| cell_ref.to_a1());

    if let Some(orig) = original {
        for attr in orig.attributes() {
            let attr = attr?;
            if attr.key.as_ref() == b"t" {
                continue;
            }
            c.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
        }
    } else {
        let a1 = inserted_a1.as_ref().expect("just set");
        c.push_attribute(("r", a1.as_str()));
    }

    if let Some(t) = cell_t {
        c.push_attribute(("t", t));
    }

    writer.write_event(Event::Start(c))?;

    if let Some(formula) = patch.formula.as_deref() {
        write_formula_element(writer, None, formula, false)?;
    }

    match body_kind {
        CellBodyKind::V(text) => {
            writer.write_event(Event::Start(BytesStart::new("v")))?;
            writer.write_event(Event::Text(BytesText::new(&text)))?;
            writer.write_event(Event::End(BytesEnd::new("v")))?;
        }
        CellBodyKind::InlineStr(text) => {
            writer.write_event(Event::Start(BytesStart::new("is")))?;
            let mut t = BytesStart::new("t");
            if needs_space_preserve(&text) {
                t.push_attribute(("xml:space", "preserve"));
            }
            writer.write_event(Event::Start(t))?;
            writer.write_event(Event::Text(BytesText::new(&text)))?;
            writer.write_event(Event::End(BytesEnd::new("t")))?;
            writer.write_event(Event::End(BytesEnd::new("is")))?;
        }
        CellBodyKind::None => {}
    }

    writer.write_event(Event::End(BytesEnd::new("c")))?;
    Ok(())
}

#[derive(Debug, Clone)]
enum CellBodyKind {
    None,
    V(String),
    InlineStr(String),
}

fn cell_representation(
    value: &CellValue,
    formula: Option<&str>,
) -> Result<(Option<&'static str>, CellBodyKind), StreamingPatchError> {
    match value {
        CellValue::Empty => Ok((None, CellBodyKind::None)),
        CellValue::Number(n) => Ok((None, CellBodyKind::V(n.to_string()))),
        CellValue::Boolean(b) => Ok((Some("b"), CellBodyKind::V(if *b { "1" } else { "0" }.to_string()))),
        CellValue::Error(err) => Ok((Some("e"), CellBodyKind::V(err.as_str().to_string()))),
        CellValue::String(s) => {
            if formula.is_some() {
                Ok((Some("str"), CellBodyKind::V(s.clone())))
            } else {
                Ok((Some("inlineStr"), CellBodyKind::InlineStr(s.clone())))
            }
        }
        other => Err(StreamingPatchError::UnsupportedCellValue(other.clone())),
    }
}

fn needs_space_preserve(s: &str) -> bool {
    s.starts_with(char::is_whitespace) || s.ends_with(char::is_whitespace)
}

#[derive(Debug, Clone, Copy)]
struct PatchBounds {
    min_row_0: u32,
    max_row_0: u32,
    min_col_0: u32,
    max_col_0: u32,
}

fn bounds_for_patches(patches: &[WorksheetCellPatch]) -> Option<PatchBounds> {
    let mut iter = patches.iter();
    let first = iter.next()?;
    let mut min_row_0 = first.cell.row;
    let mut max_row_0 = first.cell.row;
    let mut min_col_0 = first.cell.col;
    let mut max_col_0 = first.cell.col;

    for patch in iter {
        min_row_0 = min_row_0.min(patch.cell.row);
        max_row_0 = max_row_0.max(patch.cell.row);
        min_col_0 = min_col_0.min(patch.cell.col);
        max_col_0 = max_col_0.max(patch.cell.col);
    }

    Some(PatchBounds {
        min_row_0,
        max_row_0,
        min_col_0,
        max_col_0,
    })
}

fn updated_dimension_element(
    original: &BytesStart<'_>,
    bounds: PatchBounds,
) -> Result<BytesStart<'static>, StreamingPatchError> {
    let original_ref = original
        .attributes()
        .flatten()
        .find(|a| a.key.as_ref() == b"ref")
        .and_then(|a| a.unescape_value().ok())
        .map(|v| v.into_owned());

    let mut min_row_0 = bounds.min_row_0;
    let mut max_row_0 = bounds.max_row_0;
    let mut min_col_0 = bounds.min_col_0;
    let mut max_col_0 = bounds.max_col_0;

    if let Some(existing) = original_ref.as_deref() {
        if let Some((start, end)) = parse_dimension_ref(existing) {
            min_row_0 = min_row_0.min(start.row);
            max_row_0 = max_row_0.max(end.row);
            min_col_0 = min_col_0.min(start.col);
            max_col_0 = max_col_0.max(end.col);
        }
    }

    let start = CellRef::new(min_row_0, min_col_0);
    let end = CellRef::new(max_row_0, max_col_0);
    let updated_ref = if start == end {
        start.to_a1()
    } else {
        format!("{}:{}", start.to_a1(), end.to_a1())
    };

    // Preserve attribute ordering where possible by rewriting `ref` in-place.
    let mut out = BytesStart::new("dimension");
    for attr in original.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"ref" {
            out.push_attribute(("ref", updated_ref.as_str()));
        } else {
            out.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
        }
    }

    if original
        .attributes()
        .flatten()
        .all(|a| a.key.as_ref() != b"ref")
    {
        out.push_attribute(("ref", updated_ref.as_str()));
    }

    Ok(out.into_owned())
}

fn parse_dimension_ref(s: &str) -> Option<(CellRef, CellRef)> {
    let mut parts = s.split(':');
    let start = parts.next()?;
    let start = CellRef::from_a1(start).ok()?;
    let end = parts.next().and_then(|p| CellRef::from_a1(p).ok()).unwrap_or(start);
    Some((start, end))
}
