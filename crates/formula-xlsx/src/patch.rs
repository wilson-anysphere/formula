//! Part-preserving cell edit model + patch application.
//!
//! This module provides a small edit DSL (`WorkbookCellPatches`) that can be
//! applied to an existing [`crate::XlsxPackage`] without regenerating the whole
//! workbook. The implementation focuses on preserving every unrelated part
//! (charts, pivots, customXml, VBA, etc.) while rewriting only the affected
//! worksheet XML parts (plus `sharedStrings.xml` / `workbook.xml` when needed).

use std::collections::{BTreeMap, HashMap};

use formula_model::rich_text::RichText;
use formula_model::{CellRef, CellValue};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer};

use crate::openxml::{parse_relationships, rels_part_name, resolve_relationship_target};
use crate::path::resolve_target;
use crate::shared_strings::{parse_shared_strings_xml, write_shared_strings_xml, SharedStrings};
use crate::{WorkbookSheetInfo, XlsxError, XlsxPackage};

const WORKBOOK_PART: &str = "xl/workbook.xml";
const REL_TYPE_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

/// An owned set of cell edits to apply to an existing workbook package.
///
/// Patches are keyed by **worksheet (tab) name**, then by cell address.
#[derive(Debug, Clone, Default)]
pub struct WorkbookCellPatches {
    sheets: BTreeMap<String, WorksheetCellPatches>,
}

impl WorkbookCellPatches {
    /// Returns `true` if there are no pending edits.
    pub fn is_empty(&self) -> bool {
        self.sheets.values().all(WorksheetCellPatches::is_empty)
    }

    /// Get (or create) the patch set for a worksheet by name.
    pub fn sheet_mut(&mut self, sheet_name: impl Into<String>) -> &mut WorksheetCellPatches {
        self.sheets.entry(sheet_name.into()).or_default()
    }

    /// Insert/replace a patch for a single cell.
    pub fn set_cell(&mut self, sheet_name: impl Into<String>, cell: CellRef, patch: CellPatch) {
        self.sheet_mut(sheet_name).set_cell(cell, patch);
    }

    pub(crate) fn sheets(&self) -> impl Iterator<Item = (&str, &WorksheetCellPatches)> {
        self.sheets
            .iter()
            .map(|(name, patches)| (name.as_str(), patches))
    }
}

/// A set of cell edits within a single worksheet.
#[derive(Debug, Clone, Default)]
pub struct WorksheetCellPatches {
    // Deterministic ordering (row-major) makes patch application deterministic.
    cells: BTreeMap<(u32, u32), CellPatch>,
}

impl WorksheetCellPatches {
    /// Returns `true` if there are no pending edits.
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty()
    }

    /// Insert/replace a patch for a single cell.
    pub fn set_cell(&mut self, cell: CellRef, patch: CellPatch) {
        self.cells.insert((cell.row, cell.col), patch);
    }

    pub(crate) fn iter(&self) -> impl Iterator<Item = (CellRef, &CellPatch)> {
        self.cells
            .iter()
            .map(|((row, col), patch)| (CellRef::new(*row, *col), patch))
    }

    fn by_row(&self) -> BTreeMap<u32, Vec<(u32, &CellPatch)>> {
        let mut out: BTreeMap<u32, Vec<(u32, &CellPatch)>> = BTreeMap::new();
        for (&(row0, col0), patch) in &self.cells {
            out.entry(row0 + 1).or_default().push((col0, patch));
        }
        for cells in out.values_mut() {
            cells.sort_by_key(|(col, _)| *col);
        }
        out
    }
}

/// A single cell edit.
#[derive(Debug, Clone, PartialEq)]
pub enum CellPatch {
    /// Clear cell contents (formula + value). Formatting is preserved unless
    /// `style_index` overrides it.
    Clear {
        /// Optional `s` attribute override (cell XF index).
        style_index: Option<u32>,
    },
    /// Set a cell value (and optionally a formula).
    Set {
        value: CellValue,
        /// If provided, writes an `<f>` element (leading `=` is accepted).
        formula: Option<String>,
        /// Optional `s` attribute override (cell XF index).
        style_index: Option<u32>,
    },
}

impl CellPatch {
    pub fn clear() -> Self {
        Self::Clear { style_index: None }
    }

    pub fn clear_with_style(style_index: u32) -> Self {
        Self::Clear {
            style_index: Some(style_index),
        }
    }

    pub fn set_value(value: CellValue) -> Self {
        Self::Set {
            value,
            formula: None,
            style_index: None,
        }
    }

    pub fn set_value_with_formula(value: CellValue, formula: impl Into<String>) -> Self {
        Self::Set {
            value,
            formula: Some(formula.into()),
            style_index: None,
        }
    }

    pub fn style_index(&self) -> Option<u32> {
        match self {
            CellPatch::Clear { style_index } | CellPatch::Set { style_index, .. } => *style_index,
        }
    }
}

#[derive(Debug, Default)]
struct SharedStringsState {
    items: Vec<RichText>,
    plain_index: HashMap<String, u32>,
    dirty: bool,
}

impl SharedStringsState {
    fn from_part(bytes: &[u8]) -> Result<Self, XlsxError> {
        let xml = String::from_utf8(bytes.to_vec())?;
        let parsed = parse_shared_strings_xml(&xml)
            .map_err(|e| XlsxError::Invalid(format!("sharedStrings.xml parse error: {e}")))?;
        let mut plain_index = HashMap::new();
        for (idx, item) in parsed.items.iter().enumerate() {
            if item.runs.is_empty() {
                plain_index.insert(item.text.clone(), idx as u32);
            }
        }
        Ok(Self {
            items: parsed.items,
            plain_index,
            dirty: false,
        })
    }

    fn get_or_insert_plain(&mut self, text: &str) -> u32 {
        if let Some(idx) = self.plain_index.get(text).copied() {
            return idx;
        }
        let idx = self.items.len() as u32;
        self.items.push(RichText::new(text.to_string()));
        self.plain_index.insert(text.to_string(), idx);
        self.dirty = true;
        idx
    }

    fn get_or_insert_rich(&mut self, rich: &RichText) -> u32 {
        if let Some((idx, _)) = self
            .items
            .iter()
            .enumerate()
            .find(|(_, item)| *item == rich)
        {
            return idx as u32;
        }
        let idx = self.items.len() as u32;
        self.items.push(rich.clone());
        self.dirty = true;
        idx
    }

    fn write_if_dirty(&self) -> Result<Option<Vec<u8>>, XlsxError> {
        if !self.dirty {
            return Ok(None);
        }
        let xml = write_shared_strings_xml(&SharedStrings {
            items: self.items.clone(),
        })
        .map_err(|e| XlsxError::Invalid(format!("sharedStrings.xml write error: {e}")))?;
        Ok(Some(xml.into_bytes()))
    }
}

pub(crate) fn apply_cell_patches_to_package(
    pkg: &mut XlsxPackage,
    patches: &WorkbookCellPatches,
) -> Result<(), XlsxError> {
    if patches.is_empty() {
        return Ok(());
    }

    let workbook_sheets = pkg.workbook_sheets()?;

    let shared_strings_part_name = resolve_shared_strings_part_name(pkg)?
        .or_else(|| pkg.part("xl/sharedStrings.xml").map(|_| "xl/sharedStrings.xml".to_string()));
    let mut shared_strings = shared_strings_part_name
        .as_deref()
        .and_then(|part_name| pkg.part(part_name))
        .map(SharedStringsState::from_part)
        .transpose()?;

    let mut any_formula_changed = false;

    for (sheet_name, sheet_patches) in patches.sheets() {
        if sheet_patches.is_empty() {
            continue;
        }

        // Excel treats sheet names as case-insensitive; accept patches keyed by any casing.
        let sheet = workbook_sheets
            .iter()
            .find(|s| s.name.eq_ignore_ascii_case(sheet_name))
            .ok_or_else(|| XlsxError::Invalid(format!("unknown sheet name: {sheet_name}")))?;

        let worksheet_part = resolve_worksheet_part(pkg, sheet)?;
        let original = pkg
            .part(&worksheet_part)
            .ok_or_else(|| XlsxError::MissingPart(worksheet_part.clone()))?;

        let (updated, formula_changed) =
            patch_worksheet_xml(original, sheet_patches, shared_strings.as_mut())?;
        any_formula_changed |= formula_changed;

        pkg.set_part(worksheet_part, updated);
    }

    if let Some(ss) = shared_strings.as_ref() {
        if let Some(updated) = ss.write_if_dirty()? {
            let Some(part_name) = shared_strings_part_name.as_deref() else {
                return Err(XlsxError::Invalid(
                    "shared strings table was modified but part name could not be resolved"
                        .to_string(),
                ));
            };
            pkg.set_part(part_name, updated);
        }
    }

    if any_formula_changed {
        // Excel can crash / show "repaired records" dialogs if calcChain.xml gets out of sync
        // with edited formulas. We choose the conservative option: remove it and force Excel to
        // rebuild.
        pkg.parts_map_mut().remove("xl/calcChain.xml");
        ensure_workbook_full_calc_on_load(pkg)?;
    }

    Ok(())
}

fn resolve_shared_strings_part_name(pkg: &XlsxPackage) -> Result<Option<String>, XlsxError> {
    let rels_name = rels_part_name(WORKBOOK_PART);
    let rels_bytes = match pkg.part(&rels_name) {
        Some(bytes) => bytes,
        None => return Ok(None),
    };
    let rels = parse_relationships(rels_bytes)?;
    Ok(rels
        .into_iter()
        .find(|rel| rel.type_uri == REL_TYPE_SHARED_STRINGS)
        .map(|rel| resolve_target(WORKBOOK_PART, &rel.target)))
}

fn resolve_worksheet_part(
    pkg: &XlsxPackage,
    sheet: &WorkbookSheetInfo,
) -> Result<String, XlsxError> {
    resolve_relationship_target(pkg, WORKBOOK_PART, &sheet.rel_id)?.ok_or_else(|| {
        XlsxError::Invalid(format!("missing worksheet relationship for {}", sheet.name))
    })
}

fn patch_worksheet_xml(
    original: &[u8],
    patches: &WorksheetCellPatches,
    mut shared_strings: Option<&mut SharedStringsState>,
) -> Result<(Vec<u8>, bool), XlsxError> {
    // Adding a formula where none existed is always a "formula change" for the workbook.
    // (Removing formulas is detected while patching existing cells.)
    let mut formula_changed = patches.iter().any(|(_, patch)| {
        matches!(
            patch,
            CellPatch::Set {
                formula: Some(_),
                ..
            }
        )
    });

    // Track the bounds of "non-empty" patches (cells that will contain a formula or value) so we
    // can expand the worksheet `<dimension ref="..."/>` if needed.
    //
    // We don't shrink dimensions (clears), mirroring Excel's typical behavior.
    let patch_bounds = patch_bounds(patches);

    let row_patches = patches.by_row();
    let mut remaining_patch_rows: Vec<u32> = row_patches.keys().copied().collect();
    let mut patch_row_idx = 0usize;

    let mut reader = Reader::from_reader(original);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(
        original.len() + patches.cells.len() * 64,
    ));

    let mut buf = Vec::new();
    let mut saw_sheet_data = false;
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Empty(e) if local_name(e.name().as_ref()) == b"dimension" => {
                if let Some(bounds) = patch_bounds {
                    writer.write_event(Event::Empty(rewrite_dimension(&e, bounds)?))?;
                } else {
                    writer.write_event(Event::Empty(e.into_owned()))?;
                }
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"dimension" => {
                if let Some(bounds) = patch_bounds {
                    writer.write_event(Event::Start(rewrite_dimension(&e, bounds)?))?;
                } else {
                    writer.write_event(Event::Start(e.into_owned()))?;
                }
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"sheetData" => {
                saw_sheet_data = true;
                writer.write_event(Event::Start(e.into_owned()))?;
                let changed = patch_sheet_data(
                    &mut reader,
                    &mut writer,
                    &row_patches,
                    &mut remaining_patch_rows,
                    &mut patch_row_idx,
                    &mut shared_strings,
                )?;
                formula_changed |= changed;
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"sheetData" => {
                saw_sheet_data = true;
                if row_patches.is_empty() {
                    writer.write_event(Event::Empty(e.into_owned()))?;
                } else {
                    // Convert `<sheetData/>` into `<sheetData>...</sheetData>`.
                    writer.write_event(Event::Start(e.into_owned()))?;
                    for row in remaining_patch_rows.iter().skip(patch_row_idx).copied() {
                        let cells = row_patches.get(&row).map(Vec::as_slice).unwrap_or_default();
                        write_new_row(&mut writer, row, cells, &mut shared_strings)?;
                    }
                    patch_row_idx = remaining_patch_rows.len();
                    writer.write_event(Event::End(BytesEnd::new("sheetData")))?;
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"worksheet" => {
                if !saw_sheet_data && !row_patches.is_empty() {
                    // Insert missing <sheetData> just before </worksheet>.
                    writer.write_event(Event::Start(BytesStart::new("sheetData")))?;
                    for row in remaining_patch_rows.iter().skip(patch_row_idx).copied() {
                        let cells = row_patches.get(&row).map(Vec::as_slice).unwrap_or_default();
                        write_new_row(&mut writer, row, cells, &mut shared_strings)?;
                    }
                    patch_row_idx = remaining_patch_rows.len();
                    writer.write_event(Event::End(BytesEnd::new("sheetData")))?;
                }
                writer.write_event(Event::End(e.into_owned()))?;
            }
            Event::Eof => break,
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok((writer.into_inner(), formula_changed))
}

fn patch_sheet_data<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    writer: &mut Writer<Vec<u8>>,
    row_patches: &BTreeMap<u32, Vec<(u32, &CellPatch)>>,
    remaining_patch_rows: &mut [u32],
    patch_row_idx: &mut usize,
    shared_strings: &mut Option<&mut SharedStringsState>,
) -> Result<bool, XlsxError> {
    let mut buf = Vec::new();
    let mut formula_changed = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"row" => {
                let row_start = e.into_owned();
                let Some(row_num) = parse_row_r(&row_start)? else {
                    writer.write_event(Event::Start(row_start))?;
                    continue;
                };

                while *patch_row_idx < remaining_patch_rows.len()
                    && remaining_patch_rows[*patch_row_idx] < row_num
                {
                    let row = remaining_patch_rows[*patch_row_idx];
                    let cells = row_patches.get(&row).map(Vec::as_slice).unwrap_or_default();
                    write_new_row(writer, row, cells, shared_strings)?;
                    *patch_row_idx += 1;
                }

                if let Some(cells) = row_patches.get(&row_num) {
                    // Consume this patch row if it matches.
                    if *patch_row_idx < remaining_patch_rows.len()
                        && remaining_patch_rows[*patch_row_idx] == row_num
                    {
                        *patch_row_idx += 1;
                    }

                    writer.write_event(Event::Start(row_start.clone()))?;
                    let changed = patch_row(reader, writer, row_num, cells, shared_strings)?;
                    formula_changed |= changed;
                    // patch_row writes the row end.
                } else {
                    writer.write_event(Event::Start(row_start))?;
                }
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"row" => {
                let row_empty = e.into_owned();
                let Some(row_num) = parse_row_r(&row_empty)? else {
                    writer.write_event(Event::Empty(row_empty))?;
                    continue;
                };

                while *patch_row_idx < remaining_patch_rows.len()
                    && remaining_patch_rows[*patch_row_idx] < row_num
                {
                    let row = remaining_patch_rows[*patch_row_idx];
                    let cells = row_patches.get(&row).map(Vec::as_slice).unwrap_or_default();
                    write_new_row(writer, row, cells, shared_strings)?;
                    *patch_row_idx += 1;
                }

                if let Some(cells) = row_patches.get(&row_num) {
                    if *patch_row_idx < remaining_patch_rows.len()
                        && remaining_patch_rows[*patch_row_idx] == row_num
                    {
                        *patch_row_idx += 1;
                    }

                    // Convert `<row/>` into `<row>...</row>`.
                    writer.write_event(Event::Start(row_empty.clone()))?;
                    for (col, patch) in cells {
                        write_cell_patch(writer, row_num, *col, patch, None, None, shared_strings)?;
                    }
                    writer.write_event(Event::End(BytesEnd::new("row")))?;
                } else {
                    writer.write_event(Event::Empty(row_empty))?;
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"sheetData" => {
                // Insert remaining patch rows before closing </sheetData>.
                while *patch_row_idx < remaining_patch_rows.len() {
                    let row = remaining_patch_rows[*patch_row_idx];
                    let cells = row_patches.get(&row).map(Vec::as_slice).unwrap_or_default();
                    write_new_row(writer, row, cells, shared_strings)?;
                    *patch_row_idx += 1;
                }
                writer.write_event(Event::End(e.into_owned()))?;
                break;
            }
            Event::Eof => {
                return Err(XlsxError::Invalid(
                    "unexpected EOF while patching sheetData".to_string(),
                ))
            }
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(formula_changed)
}

fn patch_row<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    writer: &mut Writer<Vec<u8>>,
    row_num: u32,
    patches: &[(u32, &CellPatch)],
    shared_strings: &mut Option<&mut SharedStringsState>,
) -> Result<bool, XlsxError> {
    let mut buf = Vec::new();
    let mut patch_idx = 0usize;
    let mut formula_changed = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"c" => {
                let cell_start = e.into_owned();
                let Some((cell_ref, existing_t, existing_s)) =
                    parse_cell_addr_and_attrs(&cell_start)?
                else {
                    writer.write_event(Event::Start(cell_start))?;
                    continue;
                };

                if cell_ref.row + 1 != row_num {
                    // Defensive: mismatched cell refs are preserved unchanged.
                    writer.write_event(Event::Start(cell_start))?;
                    continue;
                }

                let col = cell_ref.col;
                while patch_idx < patches.len() && patches[patch_idx].0 < col {
                    let (patch_col, patch) = patches[patch_idx];
                    write_cell_patch(
                        writer,
                        row_num,
                        patch_col,
                        patch,
                        None,
                        None,
                        shared_strings,
                    )?;
                    patch_idx += 1;
                }

                if patch_idx < patches.len() && patches[patch_idx].0 == col {
                    let patch = patches[patch_idx].1;
                    patch_idx += 1;

                    let mut existing_formula = false;
                    let mut depth = 1usize;
                    loop {
                        match reader.read_event_into(&mut buf)? {
                            Event::Start(inner) => {
                                if depth == 1 && local_name(inner.name().as_ref()) == b"f" {
                                    existing_formula = true;
                                }
                                depth += 1;
                            }
                            Event::Empty(inner) => {
                                if depth == 1 && local_name(inner.name().as_ref()) == b"f" {
                                    existing_formula = true;
                                }
                            }
                            Event::End(inner) => {
                                depth = depth.saturating_sub(1);
                                if depth == 0 && local_name(inner.name().as_ref()) == b"c" {
                                    break;
                                }
                            }
                            Event::Eof => {
                                return Err(XlsxError::Invalid(
                                    "unexpected EOF while skipping patched cell".to_string(),
                                ))
                            }
                            _ => {}
                        }
                        buf.clear();
                    }

                    let _changed = write_cell_patch(
                        writer,
                        row_num,
                        col,
                        patch,
                        existing_t.as_deref(),
                        existing_s.as_deref(),
                        shared_strings,
                    )?;

                    // Any formula removal counts as a formula change.
                    let patch_formula = matches!(
                        patch,
                        CellPatch::Set {
                            formula: Some(_),
                            ..
                        }
                    );
                    if patch_formula || (existing_formula && !patch_formula) {
                        formula_changed = true;
                    }

                    // `_changed` indicates we wrote a cell patch (always true when called).
                } else {
                    writer.write_event(Event::Start(cell_start))?;
                }
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"c" => {
                let cell_empty = e.into_owned();
                let Some((cell_ref, existing_t, existing_s)) =
                    parse_cell_addr_and_attrs(&cell_empty)?
                else {
                    writer.write_event(Event::Empty(cell_empty))?;
                    continue;
                };

                if cell_ref.row + 1 != row_num {
                    writer.write_event(Event::Empty(cell_empty))?;
                    continue;
                }

                let col = cell_ref.col;
                while patch_idx < patches.len() && patches[patch_idx].0 < col {
                    let (patch_col, patch) = patches[patch_idx];
                    write_cell_patch(
                        writer,
                        row_num,
                        patch_col,
                        patch,
                        None,
                        None,
                        shared_strings,
                    )?;
                    patch_idx += 1;
                }

                if patch_idx < patches.len() && patches[patch_idx].0 == col {
                    let patch = patches[patch_idx].1;
                    patch_idx += 1;
                    let patch_formula = matches!(
                        patch,
                        CellPatch::Set {
                            formula: Some(_),
                            ..
                        }
                    );
                    if patch_formula {
                        formula_changed = true;
                    }
                    write_cell_patch(
                        writer,
                        row_num,
                        col,
                        patch,
                        existing_t.as_deref(),
                        existing_s.as_deref(),
                        shared_strings,
                    )?;
                } else {
                    writer.write_event(Event::Empty(cell_empty))?;
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"row" => {
                while patch_idx < patches.len() {
                    let (col, patch) = patches[patch_idx];
                    write_cell_patch(writer, row_num, col, patch, None, None, shared_strings)?;
                    patch_idx += 1;
                }
                writer.write_event(Event::End(e.into_owned()))?;
                break;
            }
            Event::Eof => {
                return Err(XlsxError::Invalid(
                    "unexpected EOF while patching row".to_string(),
                ))
            }
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(formula_changed)
}

fn write_new_row(
    writer: &mut Writer<Vec<u8>>,
    row_num: u32,
    patches: &[(u32, &CellPatch)],
    shared_strings: &mut Option<&mut SharedStringsState>,
) -> Result<(), XlsxError> {
    let mut row = BytesStart::new("row");
    row.push_attribute(("r", row_num.to_string().as_str()));
    writer.write_event(Event::Start(row))?;
    for (col, patch) in patches {
        write_cell_patch(writer, row_num, *col, patch, None, None, shared_strings)?;
    }
    writer.write_event(Event::End(BytesEnd::new("row")))?;
    Ok(())
}

fn write_cell_patch(
    writer: &mut Writer<Vec<u8>>,
    row_num: u32,
    col: u32,
    patch: &CellPatch,
    existing_t: Option<&str>,
    existing_s: Option<&str>,
    shared_strings: &mut Option<&mut SharedStringsState>,
) -> Result<bool, XlsxError> {
    let cell_ref = CellRef::new(row_num - 1, col);
    let a1 = cell_ref.to_a1();

    // Style: explicit override wins, otherwise preserve existing s=... if present.
    let style_index = patch
        .style_index()
        .or_else(|| existing_s.and_then(|s| s.parse::<u32>().ok()));

    let mut cell = String::new();
    cell.push_str(r#"<c r=""#);
    cell.push_str(&a1);
    cell.push('"');

    if let Some(s) = style_index.filter(|s| *s != 0) {
        cell.push_str(&format!(r#" s="{s}""#));
    }

    let (value, formula) = match patch {
        CellPatch::Clear { .. } => (None, None),
        CellPatch::Set { value, formula, .. } => (Some(value), formula.as_deref()),
    };

    let mut ty: Option<&str> = None;
    let mut value_xml = String::new();

    if let Some(formula) = formula {
        let formula = formula.strip_prefix('=').unwrap_or(formula);
        value_xml.push_str("<f>");
        value_xml.push_str(&escape_text(formula));
        value_xml.push_str("</f>");
    }

    if let Some(value) = value {
        match value {
            CellValue::Empty => {}
            CellValue::Number(n) => {
                value_xml.push_str("<v>");
                value_xml.push_str(&escape_text(&n.to_string()));
                value_xml.push_str("</v>");
            }
            CellValue::Boolean(b) => {
                ty = Some("b");
                value_xml.push_str("<v>");
                value_xml.push_str(if *b { "1" } else { "0" });
                value_xml.push_str("</v>");
            }
            CellValue::Error(e) => {
                ty = Some("e");
                value_xml.push_str("<v>");
                value_xml.push_str(&escape_text(e.as_str()));
                value_xml.push_str("</v>");
            }
            CellValue::String(s) => {
                // Preserve existing string storage form when possible; otherwise default to shared
                // strings when the package already has `sharedStrings.xml`.
                // If the existing cell uses an unknown/less-common type (e.g. `t="d"`), keep the
                // type and write the raw value text into `<v>`. This avoids corrupting the sheet
                // by rewriting it as a shared string / inline string when we don't understand the
                // original semantics.
                if let Some(existing_t) = existing_t {
                    if !matches!(existing_t, "s" | "b" | "e" | "n" | "str" | "inlineStr") {
                        ty = Some(existing_t);
                        value_xml.push_str("<v>");
                        value_xml.push_str(&escape_text(s));
                        value_xml.push_str("</v>");
                    } else {
                        let prefer_shared =
                            shared_strings.is_some() && existing_t != "inlineStr";

                        match (existing_t, prefer_shared) {
                            ("inlineStr", _) => {
                                ty = Some("inlineStr");
                                value_xml.push_str("<is><t");
                                if needs_space_preserve(s) {
                                    value_xml.push_str(r#" xml:space="preserve""#);
                                }
                                value_xml.push('>');
                                value_xml.push_str(&escape_text(s));
                                value_xml.push_str("</t></is>");
                            }
                            ("str", _) => {
                                ty = Some("str");
                                value_xml.push_str("<v>");
                                value_xml.push_str(&escape_text(s));
                                value_xml.push_str("</v>");
                            }
                            (_, true) => {
                                let idx = shared_strings
                                    .as_deref_mut()
                                    .map(|ss| ss.get_or_insert_plain(s))
                                    .unwrap_or(0);
                                ty = Some("s");
                                value_xml.push_str("<v>");
                                value_xml.push_str(&idx.to_string());
                                value_xml.push_str("</v>");
                            }
                            _ => {
                                ty = Some("inlineStr");
                                value_xml.push_str("<is><t");
                                if needs_space_preserve(s) {
                                    value_xml.push_str(r#" xml:space="preserve""#);
                                }
                                value_xml.push('>');
                                value_xml.push_str(&escape_text(s));
                                value_xml.push_str("</t></is>");
                            }
                        }
                    }
                } else {
                    let prefer_shared = shared_strings.is_some();

                    if prefer_shared {
                        let idx = shared_strings
                            .as_deref_mut()
                            .map(|ss| ss.get_or_insert_plain(s))
                            .unwrap_or(0);
                        ty = Some("s");
                        value_xml.push_str("<v>");
                        value_xml.push_str(&idx.to_string());
                        value_xml.push_str("</v>");
                    } else {
                        ty = Some("inlineStr");
                        value_xml.push_str("<is><t");
                        if needs_space_preserve(s) {
                            value_xml.push_str(r#" xml:space="preserve""#);
                        }
                        value_xml.push('>');
                        value_xml.push_str(&escape_text(s));
                        value_xml.push_str("</t></is>");
                    }
                }
            }
            CellValue::RichText(rich) => {
                let prefer_shared = shared_strings.is_some() && existing_t != Some("inlineStr");
                if prefer_shared {
                    let idx = shared_strings
                        .as_deref_mut()
                        .map(|ss| ss.get_or_insert_rich(rich))
                        .unwrap_or(0);
                    ty = Some("s");
                    value_xml.push_str("<v>");
                    value_xml.push_str(&idx.to_string());
                    value_xml.push_str("</v>");
                } else {
                    // Inline rich text support would require writing `<is><r>...`; for now we
                    // preserve the plain text as an inline string when a shared strings table is
                    // unavailable.
                    ty = Some("inlineStr");
                    value_xml.push_str("<is><t");
                    if needs_space_preserve(&rich.text) {
                        value_xml.push_str(r#" xml:space="preserve""#);
                    }
                    value_xml.push('>');
                    value_xml.push_str(&escape_text(&rich.text));
                    value_xml.push_str("</t></is>");
                }
            }
            CellValue::Array(_) | CellValue::Spill(_) => {
                return Err(XlsxError::Invalid(format!(
                    "unsupported cell value type for patch: {value:?}"
                )));
            }
        }
    }

    if let Some(t) = ty {
        cell.push_str(&format!(r#" t="{t}""#));
    }

    if value_xml.is_empty() {
        cell.push_str("/>");
    } else {
        cell.push('>');
        cell.push_str(&value_xml);
        cell.push_str("</c>");
    }

    writer.get_mut().extend_from_slice(cell.as_bytes());
    Ok(true)
}

fn parse_row_r(row: &BytesStart<'_>) -> Result<Option<u32>, XlsxError> {
    for attr in row.attributes() {
        let attr = attr?;
        if local_name(attr.key.as_ref()) == b"r" {
            let value = attr.unescape_value()?.into_owned();
            return Ok(value.parse::<u32>().ok());
        }
    }
    Ok(None)
}

fn parse_cell_addr_and_attrs(
    cell: &BytesStart<'_>,
) -> Result<Option<(CellRef, Option<String>, Option<String>)>, XlsxError> {
    let mut r = None;
    let mut t = None;
    let mut s = None;
    for attr in cell.attributes() {
        let attr = attr?;
        let key = local_name(attr.key.as_ref());
        let value = attr.unescape_value()?.into_owned();
        match key {
            b"r" => r = Some(value),
            b"t" => t = Some(value),
            b"s" => s = Some(value),
            _ => {}
        }
    }
    let Some(r) = r else { return Ok(None) };
    let cell_ref = CellRef::from_a1(&r).ok();
    Ok(cell_ref.map(|cr| (cr, t, s)))
}

fn ensure_workbook_full_calc_on_load(pkg: &mut XlsxPackage) -> Result<(), XlsxError> {
    let part = "xl/workbook.xml";
    let Some(bytes) = pkg.part(part) else {
        return Ok(());
    };

    if workbook_has_full_calc_on_load(bytes)? {
        return Ok(());
    }

    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(bytes.len() + 32));
    let mut buf = Vec::new();

    let mut saw_calc_pr = false;
    let mut skipping_calc_pr = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"calcPr" => {
                saw_calc_pr = true;
                skipping_calc_pr = true;
                writer
                    .get_mut()
                    .extend_from_slice(&render_calc_pr_with_full_calc_on_load(&e)?);
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"calcPr" => {
                saw_calc_pr = true;
                writer
                    .get_mut()
                    .extend_from_slice(&render_calc_pr_with_full_calc_on_load(&e)?);
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"calcPr" => {
                if skipping_calc_pr {
                    skipping_calc_pr = false;
                } else {
                    writer.write_event(Event::End(e.into_owned()))?;
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"workbook" => {
                if !saw_calc_pr {
                    // Insert calcPr before closing workbook.
                    writer
                        .get_mut()
                        .extend_from_slice(br#"<calcPr fullCalcOnLoad="1"/>"#);
                }
                writer.write_event(Event::End(e.into_owned()))?;
            }
            Event::Eof => break,
            ev if skipping_calc_pr => drop(ev),
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    pkg.set_part(part, writer.into_inner());
    Ok(())
}

fn workbook_has_full_calc_on_load(bytes: &[u8]) -> Result<bool, XlsxError> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"calcPr" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    if local_name(attr.key.as_ref()) == b"fullCalcOnLoad" {
                        let v = attr.unescape_value()?.into_owned();
                        return Ok(v == "1" || v.eq_ignore_ascii_case("true"));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(false)
}

fn render_calc_pr_with_full_calc_on_load(start: &BytesStart<'_>) -> Result<Vec<u8>, XlsxError> {
    let mut out = String::new();
    out.push_str("<calcPr");
    let mut has_full = false;
    for attr in start.attributes() {
        let attr = attr?;
        let key_bytes = attr.key.as_ref();
        let key = std::str::from_utf8(key_bytes).unwrap_or("attr");
        let local = local_name(key_bytes);
        if local == b"fullCalcOnLoad" {
            has_full = true;
            out.push_str(r#" fullCalcOnLoad="1""#);
            continue;
        }
        let value = attr.unescape_value()?.into_owned();
        out.push(' ');
        out.push_str(key);
        out.push_str(r#"=""#);
        out.push_str(&escape_text(&value).replace('\"', "&quot;"));
        out.push('"');
    }
    if !has_full {
        out.push_str(r#" fullCalcOnLoad="1""#);
    }
    out.push_str("/>");
    Ok(out.into_bytes())
}

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(idx) => &name[idx + 1..],
        None => name,
    }
}

fn needs_space_preserve(text: &str) -> bool {
    text.starts_with(char::is_whitespace) || text.ends_with(char::is_whitespace)
}

fn escape_text(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn patch_bounds(patches: &WorksheetCellPatches) -> Option<(u32, u32, u32, u32)> {
    let mut min_row = u32::MAX;
    let mut min_col = u32::MAX;
    let mut max_row = 0u32;
    let mut max_col = 0u32;

    for (cell_ref, patch) in patches.iter() {
        let is_non_empty = match patch {
            CellPatch::Clear { .. } => false,
            CellPatch::Set { value, formula, .. } => {
                formula.as_ref().is_some() || !matches!(value, CellValue::Empty)
            }
        };

        if !is_non_empty {
            continue;
        }

        // Convert to 1-based coordinates for A1 formatting.
        let row_1 = cell_ref.row.saturating_add(1);
        let col_1 = cell_ref.col.saturating_add(1);

        min_row = min_row.min(row_1);
        min_col = min_col.min(col_1);
        max_row = max_row.max(row_1);
        max_col = max_col.max(col_1);
    }

    if min_row == u32::MAX {
        None
    } else {
        Some((min_row, min_col, max_row, max_col))
    }
}

fn rewrite_dimension(
    e: &BytesStart<'_>,
    patch_bounds: (u32, u32, u32, u32),
) -> Result<BytesStart<'static>, XlsxError> {
    let (p_min_r, p_min_c, p_max_r, p_max_c) = patch_bounds;

    // `<dimension>` is typically unprefixed in real-world worksheet XML.
    let mut start = BytesStart::new("dimension");

    let mut existing_ref: Option<String> = None;
    let mut other_attrs: Vec<(Vec<u8>, String)> = Vec::new();
    for attr in e.attributes() {
        let attr = attr?;
        if local_name(attr.key.as_ref()) == b"ref" {
            existing_ref = Some(attr.unescape_value()?.into_owned());
            continue;
        }
        other_attrs.push((attr.key.as_ref().to_vec(), attr.unescape_value()?.into_owned()));
    }

    for (k, v) in other_attrs {
        start.push_attribute((k.as_slice(), v.as_bytes()));
    }

    let new_ref = existing_ref
        .as_deref()
        .and_then(parse_dimension_ref)
        .map(|(min_r, min_c, max_r, max_c)| {
            let min_r = min_r.min(p_min_r);
            let min_c = min_c.min(p_min_c);
            let max_r = max_r.max(p_max_r);
            let max_c = max_c.max(p_max_c);
            format_dimension(min_r, min_c, max_r, max_c)
        })
        .unwrap_or_else(|| format_dimension(p_min_r, p_min_c, p_max_r, p_max_c));

    start.push_attribute(("ref", new_ref.as_str()));
    Ok(start.into_owned())
}

fn parse_dimension_ref(ref_str: &str) -> Option<(u32, u32, u32, u32)> {
    let s = ref_str.trim();
    let (a, b) = s.split_once(':').unwrap_or((s, s));
    let start = CellRef::from_a1(a).ok()?;
    let end = CellRef::from_a1(b).ok()?;
    Some((
        start.row + 1,
        start.col + 1,
        end.row + 1,
        end.col + 1,
    ))
}

fn format_dimension(min_r: u32, min_c: u32, max_r: u32, max_c: u32) -> String {
    let start = CellRef::new(min_r.saturating_sub(1), min_c.saturating_sub(1)).to_a1();
    let end = CellRef::new(max_r.saturating_sub(1), max_c.saturating_sub(1)).to_a1();
    if start == end {
        start
    } else {
        format!("{start}:{end}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::XlsxPackage;
    use std::io::{Cursor, Write};

    fn build_dimension_fixture() -> Vec<u8> {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn expands_dimension_when_patching_outside_existing_bounds() {
        let bytes = build_dimension_fixture();
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        let mut patches = WorkbookCellPatches::default();
        patches.set_cell(
            "Sheet1",
            CellRef::new(2, 2), // C3
            CellPatch::set_value(CellValue::Number(42.0)),
        );

        pkg.apply_cell_patches(&patches).expect("apply patches");

        let xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
        assert!(
            xml.contains(r#"<dimension ref="A1:C3""#) || xml.contains(r#"ref="A1:C3""#),
            "expected dimension to expand, got: {xml}"
        );
    }
}
