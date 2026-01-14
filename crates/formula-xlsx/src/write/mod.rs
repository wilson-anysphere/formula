use std::borrow::Cow;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::io::{Cursor, Write};

use formula_columnar::{ColumnType as ColumnarType, Value as ColumnarValue};
use formula_engine::{parse_formula, CellAddr, ParseOptions, SerializeOptions};
use formula_model::drawings::DrawingObjectKind;
use formula_model::rich_text::{RichText, Underline};
use formula_model::{
    CellIsOperator, CellRef, CellValue, CfRule, CfRuleKind, Comment, CommentKind,
    DataValidationAssignment, DataValidationErrorStyle, DataValidationKind, DataValidationOperator,
    ErrorValue, Hyperlink, HyperlinkTarget, Outline, OutlineEntry, Range, SheetProtection,
    SheetVisibility, WorkbookProtection, WorkbookWindowState, Worksheet, WorksheetId,
};
use quick_xml::events::attributes::AttrError;
use quick_xml::events::Event;
use quick_xml::Reader;
use quick_xml::Writer;
use thiserror::Error;
use zip::write::FileOptions;
use zip::ZipWriter;

use crate::autofilter::AutoFilterParseError;
use crate::path::{rels_for_part, resolve_target};
use crate::recalc_policy::{apply_recalc_policy_to_parts, RecalcPolicyError};
use crate::shared_strings::preserve::SharedStringsEditor;
use crate::sheet_metadata::{parse_sheet_tab_color, write_sheet_tab_color};
use crate::styles::XlsxStylesEditor;
use crate::ConditionalFormattingDxfAggregation;
use crate::{
    CellValueKind, DateSystem, RecalcPolicy, SheetMeta, WorkbookKind, XlsxDocument, XlsxError,
};

const WORKBOOK_PART: &str = "xl/workbook.xml";
const WORKBOOK_RELS_PART: &str = "xl/_rels/workbook.xml.rels";
const REL_TYPE_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_TYPE_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

const SPREADSHEETML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

mod dimension;
mod data_validations;
mod sheetdata_patch;

#[derive(Debug, Error)]
pub enum WriteError {
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    #[error("xml attribute error: {0}")]
    XmlAttr(#[from] AttrError),
    #[error(transparent)]
    Styles(#[from] crate::styles::StylesPartError),
    #[error(transparent)]
    Xlsx(#[from] XlsxError),
}

impl From<RecalcPolicyError> for WriteError {
    fn from(err: RecalcPolicyError) -> Self {
        match err {
            RecalcPolicyError::Io(err) => WriteError::Io(err),
            RecalcPolicyError::Xml(err) => WriteError::Xml(err),
            RecalcPolicyError::XmlAttr(err) => WriteError::XmlAttr(err),
        }
    }
}

const WORKSHEET_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const WORKSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const DRAWING_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
const DRAWING_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
const CHART_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CHART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

#[derive(Debug)]
struct SheetStructurePlan {
    sheets: Vec<SheetMeta>,
    cell_meta_sheet_ids: HashMap<WorksheetId, WorksheetId>,
}

fn sheet_state_from_visibility(visibility: SheetVisibility) -> Option<String> {
    match visibility {
        SheetVisibility::Visible => None,
        SheetVisibility::Hidden => Some("hidden".to_string()),
        SheetVisibility::VeryHidden => Some("veryHidden".to_string()),
    }
}

fn sheet_part_number(path: &str) -> Option<u32> {
    let file = path.rsplit('/').next()?;
    if !file.starts_with("sheet") || !file.ends_with(".xml") {
        return None;
    }
    let digits = file.strip_prefix("sheet")?.strip_suffix(".xml")?;
    digits.parse::<u32>().ok()
}

fn next_sheet_part_number<'a>(paths: impl Iterator<Item = &'a str>) -> u32 {
    paths.filter_map(sheet_part_number).max().unwrap_or(0) + 1
}

fn drawing_part_number(path: &str) -> Option<u32> {
    let file = path.rsplit('/').next()?;
    if !file.starts_with("drawing") || !file.ends_with(".xml") {
        return None;
    }
    let digits = file.strip_prefix("drawing")?.strip_suffix(".xml")?;
    digits.parse::<u32>().ok()
}

fn next_drawing_part_number<'a>(paths: impl Iterator<Item = &'a str>) -> u32 {
    paths.filter_map(drawing_part_number).max().unwrap_or(0) + 1
}

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(idx) => &name[idx + 1..],
        None => name,
    }
}

fn element_prefix(name: &[u8]) -> Option<&[u8]> {
    name.iter()
        .rposition(|b| *b == b':')
        .map(|idx| &name[..idx])
}

fn prefixed_tag(prefix: Option<&str>, local: &str) -> String {
    match prefix {
        Some(prefix) => format!("{prefix}:{local}"),
        None => local.to_string(),
    }
}

fn office_relationships_prefix_from_xmlns(
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<Option<String>, WriteError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        let key = attr.key.as_ref();
        let Some(prefix) = key.strip_prefix(b"xmlns:") else {
            continue;
        };
        if attr.value.as_ref() == crate::xml::OFFICE_RELATIONSHIPS_NS.as_bytes() {
            return Ok(Some(String::from_utf8_lossy(prefix).into_owned()));
        }
    }
    Ok(None)
}

fn worksheet_has_default_spreadsheetml_ns(
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<bool, WriteError> {
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"xmlns" && attr.value.as_ref() == SPREADSHEETML_NS.as_bytes() {
            return Ok(true);
        }
    }
    Ok(false)
}

fn content_types_has_default_ns(e: &quick_xml::events::BytesStart<'_>) -> Result<bool, WriteError> {
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"xmlns" && attr.value.as_ref() == CONTENT_TYPES_NS.as_bytes() {
            return Ok(true);
        }
    }
    Ok(false)
}

fn plan_sheet_structure(
    doc: &XlsxDocument,
    parts: &mut BTreeMap<String, Vec<u8>>,
    is_new: bool,
) -> Result<SheetStructurePlan, WriteError> {
    // Match model sheets to preserved sheet metadata using the XLSX identity fields when
    // available. This makes sheet structure edits robust even if the workbook model
    // was reconstructed (e.g. loaded from persisted state) with different internal
    // `WorksheetId`s.
    let mut meta_by_ws_id: HashMap<WorksheetId, usize> = HashMap::new();
    let mut meta_by_rel_id: HashMap<&str, usize> = HashMap::new();
    let mut meta_by_sheet_id: HashMap<u32, usize> = HashMap::new();
    for (idx, meta) in doc.meta.sheets.iter().enumerate() {
        meta_by_ws_id.insert(meta.worksheet_id, idx);
        meta_by_rel_id.insert(meta.relationship_id.as_str(), idx);
        meta_by_sheet_id.insert(meta.sheet_id, idx);
    }

    let mut matched_meta_idxs: HashSet<usize> = HashSet::new();
    let mut matched_meta_by_ws_id: HashMap<WorksheetId, usize> = HashMap::new();
    for sheet in &doc.workbook.sheets {
        let idx = sheet
            .xlsx_rel_id
            .as_deref()
            .and_then(|rid| meta_by_rel_id.get(rid).copied())
            .or_else(|| {
                sheet
                    .xlsx_sheet_id
                    .and_then(|sid| meta_by_sheet_id.get(&sid).copied())
            })
            .or_else(|| meta_by_ws_id.get(&sheet.id).copied());

        if let Some(idx) = idx {
            matched_meta_idxs.insert(idx);
            matched_meta_by_ws_id.insert(sheet.id, idx);
        }
    }

    let mut cell_meta_sheet_ids: HashMap<WorksheetId, WorksheetId> = HashMap::new();
    for (worksheet_id, idx) in &matched_meta_by_ws_id {
        let meta_sheet_id = doc
            .meta
            .sheets
            .get(*idx)
            .map(|meta| meta.worksheet_id)
            .unwrap_or(*worksheet_id);
        cell_meta_sheet_ids.insert(*worksheet_id, meta_sheet_id);
    }

    let removed: Vec<SheetMeta> = doc
        .meta
        .sheets
        .iter()
        .enumerate()
        .filter(|(idx, _)| !matched_meta_idxs.contains(idx))
        .map(|(_, meta)| meta.clone())
        .collect();

    // Excel allocates new `sheetId` values as `max(existing)+1`. We intentionally consider the
    // entire original workbook sheet list (including sheets that may be deleted in this edit)
    // to avoid reusing sheetIds that other preserved/orphaned parts might still reference.
    let mut next_sheet_id = doc
        .meta
        .sheets
        .iter()
        .map(|meta| meta.sheet_id)
        .max()
        .unwrap_or(0)
        + 1;

    let mut next_rel_id_num = if is_new {
        doc.meta
            .sheets
            .iter()
            .enumerate()
            .filter(|(idx, _)| matched_meta_idxs.contains(idx))
            .filter_map(|(_, meta)| {
                meta.relationship_id
                    .strip_prefix("rId")?
                    .parse::<u32>()
                    .ok()
            })
            .max()
            .unwrap_or(0)
            + 1
    } else {
        parts
            .get("xl/_rels/workbook.xml.rels")
            .and_then(|b| std::str::from_utf8(b).ok())
            .map(next_relationship_id_in_xml)
            .unwrap_or(1)
    };

    let existing_paths = doc.meta.sheets.iter().map(|meta| meta.path.as_str());
    let part_paths = parts.keys().map(|p| p.as_str());
    let mut next_sheet_part = next_sheet_part_number(existing_paths.chain(part_paths));
    let mut used_paths: HashSet<String> = doc.meta.sheets.iter().map(|m| m.path.clone()).collect();

    let mut sheets: Vec<SheetMeta> = Vec::with_capacity(doc.workbook.sheets.len());
    let mut added: Vec<SheetMeta> = Vec::new();

    for sheet in &doc.workbook.sheets {
        if let Some(idx) = matched_meta_by_ws_id.get(&sheet.id).copied() {
            let mut meta = doc.meta.sheets[idx].clone();
            meta.worksheet_id = sheet.id;
            meta.state = sheet_state_from_visibility(sheet.visibility);
            sheets.push(meta);
            continue;
        }

        let relationship_id = format!("rId{next_rel_id_num}");
        next_rel_id_num += 1;

        let mut path;
        loop {
            path = format!("xl/worksheets/sheet{next_sheet_part}.xml");
            next_sheet_part += 1;
            if !used_paths.contains(&path) && !parts.contains_key(&path) {
                break;
            }
        }
        used_paths.insert(path.clone());

        let meta = SheetMeta {
            worksheet_id: sheet.id,
            sheet_id: next_sheet_id,
            relationship_id,
            state: sheet_state_from_visibility(sheet.visibility),
            path,
        };
        next_sheet_id += 1;

        added.push(meta.clone());
        sheets.push(meta);
    }

    if !is_new && (!added.is_empty() || !removed.is_empty()) {
        for meta in &removed {
            parts.remove(&meta.path);
            parts.remove(&crate::openxml::rels_part_name(&meta.path));
        }

        patch_workbook_rels_for_sheet_edits(parts, &removed, &added)?;
        patch_content_types_for_sheet_edits(parts, &removed, &added)?;
    }

    Ok(SheetStructurePlan {
        sheets,
        cell_meta_sheet_ids,
    })
}

pub fn write_to_vec(doc: &XlsxDocument) -> Result<Vec<u8>, WriteError> {
    write_to_vec_with_recalc_policy(doc, RecalcPolicy::default())
}

pub fn write_to_vec_with_recalc_policy(
    doc: &XlsxDocument,
    recalc_policy: RecalcPolicy,
) -> Result<Vec<u8>, WriteError> {
    let formula_changed = formulas_changed(doc);
    let changed_formula_cells = if recalc_policy.clear_cached_values_on_formula_change {
        formula_changed_cells(doc)
    } else {
        HashSet::new()
    };

    let mut parts = build_parts(doc, &changed_formula_cells)?;
    if formula_changed {
        apply_recalc_policy_to_parts(&mut parts, recalc_policy)?;
    }

    if parts
        .keys()
        .any(|name| crate::zip_util::zip_part_names_equivalent(name.as_str(), "xl/vbaProject.bin"))
    {
        crate::macro_repair::ensure_xlsm_content_types(&mut parts)?;
        crate::macro_repair::ensure_workbook_rels_has_vba(&mut parts)?;
        crate::macro_repair::ensure_vba_project_rels_has_signature(&mut parts)?;
    }

    // Deterministic ordering helps debugging and makes fixtures stable.
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in parts.iter_mut() {
        zip.start_file(name, options)?;
        zip.write_all(bytes)?;
    }

    let cursor = zip.finish()?;
    Ok(cursor.into_inner())
}

fn formulas_changed(doc: &XlsxDocument) -> bool {
    let mut seen: HashSet<(WorksheetId, CellRef)> = HashSet::new();

    // WorksheetId values can differ between the in-memory workbook and the preserved metadata
    // (e.g. if the model is reconstructed from persisted state). Use the workbook's stable XLSX
    // identity fields when available to map back to the metadata worksheet IDs.
    let mut meta_by_ws_id: HashMap<WorksheetId, usize> = HashMap::new();
    let mut meta_by_rel_id: HashMap<&str, usize> = HashMap::new();
    let mut meta_by_sheet_id: HashMap<u32, usize> = HashMap::new();
    for (idx, meta) in doc.meta.sheets.iter().enumerate() {
        meta_by_ws_id.insert(meta.worksheet_id, idx);
        meta_by_rel_id.insert(meta.relationship_id.as_str(), idx);
        meta_by_sheet_id.insert(meta.sheet_id, idx);
    }

    let mut workbook_to_meta_sheet_id: HashMap<WorksheetId, WorksheetId> = HashMap::new();
    let mut meta_to_workbook_sheet_id: HashMap<WorksheetId, WorksheetId> = HashMap::new();
    for sheet in &doc.workbook.sheets {
        let idx = sheet
            .xlsx_rel_id
            .as_deref()
            .and_then(|rid| meta_by_rel_id.get(rid).copied())
            .or_else(|| {
                sheet
                    .xlsx_sheet_id
                    .and_then(|sid| meta_by_sheet_id.get(&sid).copied())
            })
            .or_else(|| meta_by_ws_id.get(&sheet.id).copied());
        let Some(idx) = idx else {
            continue;
        };
        let meta_sheet_id = doc
            .meta
            .sheets
            .get(idx)
            .map(|meta| meta.worksheet_id)
            .unwrap_or(sheet.id);
        workbook_to_meta_sheet_id.insert(sheet.id, meta_sheet_id);
        meta_to_workbook_sheet_id
            .entry(meta_sheet_id)
            .or_insert(sheet.id);
    }

    let mut shared_formulas_by_sheet: HashMap<WorksheetId, HashMap<u32, SharedFormulaGroup>> =
        HashMap::new();

    for sheet in &doc.workbook.sheets {
        let sheet_id = sheet.id;
        let meta_sheet_id = workbook_to_meta_sheet_id
            .get(&sheet_id)
            .copied()
            .unwrap_or(sheet_id);
        let shared_formulas = shared_formulas_by_sheet
            .entry(meta_sheet_id)
            .or_insert_with(|| shared_formula_groups(doc, meta_sheet_id));
        for (cell_ref, cell) in sheet.iter_cells() {
            let Some(formula) = cell.formula.as_deref() else {
                continue;
            };
            if strip_leading_equals(formula).is_empty() {
                continue;
            }

            seen.insert((meta_sheet_id, cell_ref));
            let meta_formula = doc
                .meta
                .cell_meta
                .get(&(meta_sheet_id, cell_ref))
                .and_then(|m| m.formula.as_ref());

            // `read` expands textless shared-formula follower cells into explicit formulas in the
            // in-memory model. Those synthesized formulas should not count as edits when deciding
            // whether we need to drop `xl/calcChain.xml`.
            if let Some(meta_formula) = meta_formula {
                let is_textless_shared_follower = meta_formula.t.as_deref() == Some("shared")
                    && meta_formula.reference.is_none()
                    && meta_formula.file_text.is_empty()
                    && meta_formula.shared_index.is_some();
                if is_textless_shared_follower {
                    if let Some(shared_index) = meta_formula.shared_index {
                        if let Some(expected) =
                            shared_formula_expected(shared_formulas, shared_index, cell_ref)
                        {
                            if !formula_text_differs(Some(expected.as_str()), Some(formula)) {
                                continue;
                            }
                        }
                    }

                    // If we can't validate equivalence with the shared-formula master, be
                    // conservative and treat this as a formula change.
                    return true;
                }
            }

            let baseline = meta_formula.map(|f| f.file_text.as_str());
            if formula_text_differs(baseline, Some(formula)) {
                return true;
            }
        }
    }

    // Detect removed formulas (cells that had a stored formula, but now have none).
    for ((sheet_id, cell_ref), meta) in &doc.meta.cell_meta {
        let Some(formula_meta) = meta.formula.as_ref() else {
            continue;
        };
        if formula_meta.file_text.is_empty() {
            // Textless shared formula followers are represented in the model as expanded formulas,
            // but do not correspond to explicit `f` text in the file. They are handled in the
            // loop above, which checks whether the expanded formula still matches the shared group.
            continue;
        }
        if seen.contains(&(*sheet_id, *cell_ref)) {
            continue;
        }
        let model_formula = doc
            .workbook
            .sheet(
                meta_to_workbook_sheet_id
                    .get(sheet_id)
                    .copied()
                    .unwrap_or(*sheet_id),
            )
            .and_then(|s| s.cell(*cell_ref))
            .and_then(|c| c.formula.as_deref());
        if formula_text_differs(Some(formula_meta.file_text.as_str()), model_formula) {
            return true;
        }
    }

    false
}

/// Returns the set of formula cells whose *material* formula text differs from the baseline
/// metadata, excluding formula removals.
///
/// This mirrors the comparison logic in [`formulas_changed`], but instead of returning early it
/// records the addresses of changed/added formulas so callers can selectively clear cached `<v>`
/// values for those edited cells.
fn formula_changed_cells(doc: &XlsxDocument) -> HashSet<(WorksheetId, CellRef)> {
    let mut changed: HashSet<(WorksheetId, CellRef)> = HashSet::new();

    // WorksheetId values can differ between the in-memory workbook and the preserved metadata
    // (e.g. if the model is reconstructed from persisted state). Use the workbook's stable XLSX
    // identity fields when available to map back to the metadata worksheet IDs.
    let mut meta_by_ws_id: HashMap<WorksheetId, usize> = HashMap::new();
    let mut meta_by_rel_id: HashMap<&str, usize> = HashMap::new();
    let mut meta_by_sheet_id: HashMap<u32, usize> = HashMap::new();
    for (idx, meta) in doc.meta.sheets.iter().enumerate() {
        meta_by_ws_id.insert(meta.worksheet_id, idx);
        meta_by_rel_id.insert(meta.relationship_id.as_str(), idx);
        meta_by_sheet_id.insert(meta.sheet_id, idx);
    }

    let mut workbook_to_meta_sheet_id: HashMap<WorksheetId, WorksheetId> = HashMap::new();
    for sheet in &doc.workbook.sheets {
        let idx = sheet
            .xlsx_rel_id
            .as_deref()
            .and_then(|rid| meta_by_rel_id.get(rid).copied())
            .or_else(|| {
                sheet
                    .xlsx_sheet_id
                    .and_then(|sid| meta_by_sheet_id.get(&sid).copied())
            })
            .or_else(|| meta_by_ws_id.get(&sheet.id).copied());
        let Some(idx) = idx else {
            continue;
        };
        let meta_sheet_id = doc
            .meta
            .sheets
            .get(idx)
            .map(|meta| meta.worksheet_id)
            .unwrap_or(sheet.id);
        workbook_to_meta_sheet_id.insert(sheet.id, meta_sheet_id);
    }

    let mut shared_formulas_by_sheet: HashMap<WorksheetId, HashMap<u32, SharedFormulaGroup>> =
        HashMap::new();

    for sheet in &doc.workbook.sheets {
        let sheet_id = sheet.id;
        let meta_sheet_id = workbook_to_meta_sheet_id
            .get(&sheet_id)
            .copied()
            .unwrap_or(sheet_id);
        let shared_formulas = shared_formulas_by_sheet
            .entry(meta_sheet_id)
            .or_insert_with(|| shared_formula_groups(doc, meta_sheet_id));
        for (cell_ref, cell) in sheet.iter_cells() {
            let Some(formula) = cell.formula.as_deref() else {
                continue;
            };
            if strip_leading_equals(formula).is_empty() {
                continue;
            }

            let meta_formula = doc
                .meta
                .cell_meta
                .get(&(meta_sheet_id, cell_ref))
                .and_then(|m| m.formula.as_ref());

            // `read` expands textless shared-formula follower cells into explicit formulas in the
            // in-memory model. Those synthesized formulas should not count as edits when deciding
            // whether we need to drop `xl/calcChain.xml`, and similarly should not trigger cached
            // value clearing.
            if let Some(meta_formula) = meta_formula {
                let is_textless_shared_follower = meta_formula.t.as_deref() == Some("shared")
                    && meta_formula.reference.is_none()
                    && meta_formula.file_text.is_empty()
                    && meta_formula.shared_index.is_some();
                if is_textless_shared_follower {
                    if let Some(shared_index) = meta_formula.shared_index {
                        if let Some(expected) =
                            shared_formula_expected(shared_formulas, shared_index, cell_ref)
                        {
                            if !formula_text_differs(Some(expected.as_str()), Some(formula)) {
                                continue;
                            }
                        }
                    }

                    // If we can't validate equivalence with the shared-formula master, be
                    // conservative and treat this as a formula change.
                    changed.insert((meta_sheet_id, cell_ref));
                    continue;
                }
            }

            let baseline = meta_formula.map(|f| f.file_text.as_str());
            if formula_text_differs(baseline, Some(formula)) {
                changed.insert((meta_sheet_id, cell_ref));
            }
        }
    }

    changed
}

fn formula_text_differs(baseline_file_text: Option<&str>, model_formula: Option<&str>) -> bool {
    let baseline = normalize_formula_for_compare(baseline_file_text);
    let model = normalize_formula_for_compare(model_formula);
    baseline != model
}

fn normalize_formula_for_compare(formula: Option<&str>) -> Option<String> {
    let formula = formula?;
    // SpreadsheetML `<f>` text is typically stored without surrounding whitespace, but
    // fixtures (and some generators) may pretty-print formulas with indentation/newlines.
    // Treat those as semantically equivalent so we don't trigger recalc safety on a no-op save.
    let trimmed = formula.trim();
    let stripped = strip_leading_equals(trimmed).trim();
    if stripped.is_empty() {
        return None;
    }
    Some(crate::formula_text::strip_xlfn_prefixes(stripped))
}

fn build_parts(
    doc: &XlsxDocument,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
) -> Result<BTreeMap<String, Vec<u8>>, WriteError> {
    let mut parts = doc.parts.clone();
    let is_new = parts.is_empty();

    let sheet_plan = plan_sheet_structure(doc, &mut parts, is_new)?;
    if is_new {
        parts = generate_minimal_package(&sheet_plan.sheets, doc.workbook_kind)?;
    }

    let (mut styles_part_name, mut shared_strings_part_name) = (
        "xl/styles.xml".to_string(),
        "xl/sharedStrings.xml".to_string(),
    );
    let mut synthesize_styles_for_missing_relationship = false;
    if let Some(rels) = parts.get(WORKBOOK_RELS_PART).map(|b| b.as_slice()) {
        if let Some(target) = relationship_target_by_type(rels, REL_TYPE_STYLES)? {
            styles_part_name = resolve_target(WORKBOOK_PART, &target);
        } else if relationships_root_is_prefix_only(rels)? {
            // Some workbooks use a prefix-only relationships namespace
            // (`<rel:Relationships xmlns:rel="...">`). When we need to synthesize a missing styles
            // relationship for those files, we also synthesize `xl/styles.xml` so we can insert a
            // prefixed `<rel:Relationship ...>` element deterministically.
            synthesize_styles_for_missing_relationship = true;
        }
        if let Some(target) = relationship_target_by_type(rels, REL_TYPE_SHARED_STRINGS)? {
            shared_strings_part_name = resolve_target(WORKBOOK_PART, &target);
        }
    }

    let original_shared_strings = parts
        .get(&shared_strings_part_name)
        .map(|bytes| bytes.as_slice());
    let (shared_strings_xml, shared_string_lookup) = build_shared_strings_xml(
        doc,
        &sheet_plan.sheets,
        &sheet_plan.cell_meta_sheet_ids,
        original_shared_strings,
    )?;
    if is_new || !shared_string_lookup.is_empty() || parts.contains_key(&shared_strings_part_name) {
        parts.insert(shared_strings_part_name.clone(), shared_strings_xml);
    }

    // Parse/update styles.xml (cellXfs) so cell `s` attributes refer to real xf indices.
    let mut style_table = doc.workbook.styles.clone();
    let mut styles_editor = XlsxStylesEditor::parse_or_default(
        parts.get(&styles_part_name).map(|b| b.as_slice()),
        &mut style_table,
    )?;

    // Conditional formatting uses a single workbook-global `<dxfs>` table inside `styles.xml`, but
    // the in-memory model stores differential formats per-sheet.
    //
    // For round-trip writers, treat the existing `styles.xml` `<dxfs>` table as the canonical base
    // (so existing `cfRule/@dxfId` indices remain valid), and only append additional dxfs when new
    // sheets/rules introduce them.
    let existing_cf_dxfs = styles_editor.styles_part().conditional_formatting_dxfs();
    let cf_dxfs = ConditionalFormattingDxfAggregation::from_worksheets_with_base_global_dxfs(
        &doc.workbook.sheets,
        &existing_cf_dxfs,
    );

    // Collect all style ids referenced by the workbook so `styles.xml` includes corresponding
    // `<xf>` entries and we can map model `style_id` -> SpreadsheetML `cellXf` index.
    //
    // This includes:
    // - per-cell styles (`c/@s`)
    // - per-row default styles (`row/@s`)
    // - per-column default styles (`col/@style`)
    let style_ids = doc.workbook.sheets.iter().flat_map(|sheet| {
        sheet
            .iter_cells()
            .map(|(_, cell)| cell.style_id)
            .chain(sheet.row_properties.values().filter_map(|props| props.style_id))
            .chain(sheet.col_properties.values().filter_map(|props| props.style_id))
    });
    let style_ids = style_ids.filter(|style_id| *style_id != 0);
    let style_to_xf = styles_editor.ensure_styles_for_style_ids(style_ids, &style_table)?;
    // Preserve workbooks that omit a `styles.xml` part: if the source package didn't have one and
    // the model doesn't reference any non-default style IDs, keep the part absent on round-trip.
    let has_existing_styles_part = parts.contains_key(&styles_part_name);
    let should_write_styles_part = is_new
        || !style_to_xf.is_empty()
        || has_existing_styles_part
        || synthesize_styles_for_missing_relationship
        || !cf_dxfs.global_dxfs.is_empty();
    if should_write_styles_part {
        if !cf_dxfs.global_dxfs.is_empty() {
            if is_new || !has_existing_styles_part {
                // For new/synthesized styles.xml, we control the full payload and can write the
                // global `<dxfs>` table from scratch.
                styles_editor
                    .styles_part_mut()
                    .set_conditional_formatting_dxfs(&cf_dxfs.global_dxfs);
            } else if cf_dxfs.global_dxfs.len() > existing_cf_dxfs.len() {
                // For existing workbooks, preserve existing `<dxf>` entries (which may contain
                // unknown/unmodeled XML) and only append newly introduced dxfs.
                let new_dxfs = &cf_dxfs.global_dxfs[existing_cf_dxfs.len()..];
                styles_editor
                    .styles_part_mut()
                    .append_conditional_formatting_dxfs(new_dxfs);
            }
        }
        parts.insert(styles_part_name.clone(), styles_editor.to_styles_xml_bytes());
    }

    // Ensure core relationship/content types metadata exists when we synthesize new
    // parts for existing packages. For existing relationships we preserve IDs by
    // only adding missing entries with a new `rIdN`.
    if parts.contains_key(&shared_strings_part_name) {
        ensure_content_types_override(
            &mut parts,
            &format!("/{shared_strings_part_name}"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        )?;
        ensure_workbook_rels_has_relationship(
            &mut parts,
            REL_TYPE_SHARED_STRINGS,
            &relationship_target_from_workbook(&shared_strings_part_name),
        )?;
    }
    if parts.contains_key(&styles_part_name) {
        ensure_content_types_override(
            &mut parts,
            &format!("/{styles_part_name}"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
        )?;
        ensure_workbook_rels_has_relationship(
            &mut parts,
            REL_TYPE_STYLES,
            &relationship_target_from_workbook(&styles_part_name),
        )?;
    }

    let workbook_orig = parts.get("xl/workbook.xml").map(|b| b.as_slice());
    parts.insert(
        "xl/workbook.xml".to_string(),
        write_workbook_xml(doc, workbook_orig, &sheet_plan.sheets)?,
    );

    let mut next_drawing_part = next_drawing_part_number(parts.keys().map(|p| p.as_str()));

    for (sheet_index, sheet_meta) in sheet_plan.sheets.iter().enumerate() {
        // Chartsheets (`xl/chartsheets/*.xml`) are not modeled semantically today. They must be
        // preserved byte-for-byte to keep their DrawingML relationship graph intact (chartsheet ->
        // drawing -> chart).
        //
        // The worksheet patching pipeline re-serializes XML via quick-xml which changes formatting
        // (whitespace/indentation), so running it on chartsheets would break no-op round-trips even
        // when the sheet model is untouched.
        if sheet_meta.path.starts_with("xl/chartsheets/") {
            continue;
        }
        let sheet = doc.workbook.sheet(sheet_meta.worksheet_id).ok_or_else(|| {
            WriteError::Io(std::io::Error::new(
                std::io::ErrorKind::NotFound,
                "worksheet not found",
            ))
        })?;
        let orig = parts.get(&sheet_meta.path).map(|b| b.as_slice());
        let is_new_sheet = orig.is_none();
        let rels_part = rels_for_part(&sheet_meta.path);
        let rels_xml: Option<String> = match parts.get(&rels_part) {
            Some(bytes) => Some(
                std::str::from_utf8(bytes).map_err(|e| {
                    WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e))
                })?
                .to_string(),
            ),
            None => None,
        };

        let (
            orig_tab_color,
            orig_merges,
            orig_hyperlinks,
            orig_drawing_rel_id,
            orig_data_validations,
            orig_views,
            orig_sheet_format,
            orig_cols,
            orig_autofilter,
            orig_sheet_protection,
            orig_has_data_validations,
            orig_has_conditional_formatting,
        ) = if let Some(orig) = orig {
            let orig_xml = std::str::from_utf8(orig).map_err(|e| {
                WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e))
            })?;
            let orig_tab_color = parse_sheet_tab_color(orig_xml)?;

            let orig_views = parse_sheet_view_settings(orig_xml)?;
            let orig_sheet_format = parse_sheet_format_settings(orig_xml)?;
            let orig_cols = parse_col_properties(orig_xml, &styles_editor)?;
            let orig_sheet_protection = parse_sheet_protection(orig_xml)?;
            let orig_has_data_validations = worksheet_has_data_validations(orig_xml)?;

            let orig_merges = crate::merge_cells::read_merge_cells_from_worksheet_xml(orig_xml)
                .map_err(|err| match err {
                    crate::merge_cells::MergeCellsError::Xml(e) => WriteError::Xml(e),
                    crate::merge_cells::MergeCellsError::Attr(e) => WriteError::XmlAttr(e),
                    crate::merge_cells::MergeCellsError::Utf8(e) => {
                        WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e))
                    }
                    crate::merge_cells::MergeCellsError::InvalidRef(r) => {
                        WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, r))
                    }
                    crate::merge_cells::MergeCellsError::Zip(e) => WriteError::Zip(e),
                    crate::merge_cells::MergeCellsError::Io(e) => WriteError::Io(e),
                })?;

            let orig_hyperlinks =
                crate::parse_worksheet_hyperlinks(orig_xml, rels_xml.as_deref())?;

            let orig_drawing_rel_id = worksheet_drawing_rel_id(orig_xml)?;

            let orig_data_validations = if orig_has_data_validations {
                crate::data_validations::read_data_validations_from_worksheet_xml(orig_xml)
                    .unwrap_or_default()
            } else {
                Vec::new()
            };

            let orig_autofilter = crate::autofilter::parse_worksheet_autofilter(orig_xml).map_err(
                |err| match err {
                    AutoFilterParseError::Xml(e) => WriteError::Xml(e),
                    AutoFilterParseError::Attr(e) => WriteError::XmlAttr(e),
                    AutoFilterParseError::MissingRef => WriteError::Io(std::io::Error::new(
                        std::io::ErrorKind::InvalidData,
                        "missing worksheet autoFilter ref attribute",
                    )),
                    AutoFilterParseError::InvalidRef(e) => WriteError::Io(std::io::Error::new(
                        std::io::ErrorKind::InvalidData,
                        e.to_string(),
                    )),
                },
            )?;

            // Conditional formatting blocks are only meaningful when they contain `<cfRule>`
            // children. Some producers (or corrupted files) may contain an empty
            // `<conditionalFormatting/>` element. Treat those as "no conditional formatting" so
            // callers can still insert conditional formatting rules via the streaming patcher.
            let orig_has_conditional_formatting =
                orig_xml.contains("<cfRule") || orig_xml.contains(":cfRule");

            (
                orig_tab_color,
                orig_merges,
                orig_hyperlinks,
                orig_drawing_rel_id,
                orig_data_validations,
                orig_views,
                orig_sheet_format,
                orig_cols,
                orig_autofilter,
                orig_sheet_protection,
                orig_has_data_validations,
                orig_has_conditional_formatting,
            )
        } else {
            (
                None,
                Vec::new(),
                Vec::new(),
                None,
                Vec::new(),
                SheetViewSettings::default(),
                SheetFormatSettings::default(),
                BTreeMap::new(),
                None,
                None,
                false,
                false,
            )
        };

        let current_merges = normalize_merge_ranges(sheet.merged_regions.iter().map(|r| r.range));
        let orig_merges = normalize_merge_ranges(orig_merges.iter().copied());
        let merges_changed = current_merges != orig_merges;
        let tab_color_changed = sheet.tab_color != orig_tab_color;

        let current_hyperlinks = normalize_hyperlinks(&assign_hyperlink_rel_ids(
            &sheet.hyperlinks,
            rels_xml.as_deref(),
        ));
        let orig_hyperlinks = normalize_hyperlinks(&orig_hyperlinks);
        let hyperlinks_changed = current_hyperlinks != orig_hyperlinks;

        let desired_views = SheetViewSettings::from_sheet(sheet);
        let views_changed = desired_views != orig_views;

        let desired_sheet_format = SheetFormatSettings::from_sheet(sheet);
        let sheet_format_changed = sheet_format_needs_patch(desired_sheet_format, orig_sheet_format);

        let cols_changed = &sheet.col_properties != &orig_cols;

        // `parse_col_properties` ignores outline-related attributes (`outlineLevel`, `collapsed`).
        // If callers mutate only `sheet.outline.cols` (including clearing outline levels), we
        // still need to rewrite `<cols>` so the outline metadata is preserved/emitted.
        let outline_cols_changed = if let Some(orig) = orig {
            let desired = cols_xml_props_from_sheet(sheet, &style_to_xf);
            let orig_xml = std::str::from_utf8(orig).map_err(|e| {
                WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e))
            })?;
            let original_cols = parse_cols_xml_props(orig_xml, &styles_editor)?;
            desired != original_cols
        } else {
            false
        };

        let autofilter_changed = sheet.auto_filter.as_ref() != orig_autofilter.as_ref();

        // New sheets rendered from scratch already include conditional formatting blocks (with
        // workbook-global `dxfId` remapping), so avoid re-inserting them via the streaming patcher
        // which currently operates on the model's per-sheet `dxf_id` values.
        let conditional_formatting_changed = !is_new_sheet
            && !sheet.conditional_formatting_rules.is_empty()
            && !orig_has_conditional_formatting;

        let sheet_protection_changed = if sheet.sheet_protection.enabled {
            match orig_sheet_protection.as_ref() {
                Some(orig) => orig != &sheet.sheet_protection,
                None => true,
            }
        } else {
            // When protection is disabled, we canonicalize by removing any `<sheetProtection>`
            // element that may still be present in the source document (including cases where it
            // was stored as `<sheetProtection sheet="0" .../>`).
            orig_sheet_protection.is_some()
        };

        let data_validations_changed = if sheet.data_validations.is_empty() && !orig_has_data_validations
        {
            false
        } else {
            let model_norm = normalize_data_validations(&sheet.data_validations);
            let orig_norm = normalize_parsed_data_validations(&orig_data_validations);
            if model_norm.is_empty() && orig_norm.is_empty() && orig_has_data_validations {
                // We couldn't parse any modeled validations from the existing worksheet, but we
                // can see a `<dataValidations>` block. Preserve it for no-op round trips unless the
                // caller supplies new validations.
                false
            } else {
                model_norm != orig_norm
            }
        };

        let local_to_global_dxf = cf_dxfs
            .local_to_global_by_sheet
            .get(&sheet.id)
            .map(|v| v.as_slice());
        let sheet_xml_bytes = write_worksheet_xml(
            doc,
            sheet_meta,
            sheet,
            orig,
            &shared_string_lookup,
            &style_to_xf,
            &sheet_plan.cell_meta_sheet_ids,
            local_to_global_dxf,
            changed_formula_cells,
        )?;
        let has_drawings = !sheet.drawings.is_empty();
        // Best-effort: only rewrite drawing-related parts when the worksheet model references new
        // media that does not already exist in the package, or when the in-memory drawing object
        // list diverges from the source drawing part.
        //
        // This is important for chart-heavy fixtures where a no-op round-trip is expected to
        // preserve `xl/drawings/*` and `.rels` parts byte-for-byte.
        let existing_sheet_drawing_rel = match rels_xml.as_deref() {
            Some(xml) => {
                if let Some(rid) = orig_drawing_rel_id.as_deref() {
                    match relationship_target_by_id(xml.as_bytes(), rid)? {
                        Some(target) => Some((rid.to_string(), target)),
                        None => relationship_id_and_target_by_type(xml.as_bytes(), DRAWING_REL_TYPE)?,
                    }
                } else {
                    relationship_id_and_target_by_type(xml.as_bytes(), DRAWING_REL_TYPE)?
                }
            }
            None => None,
        };
        let existing_drawing_part_path = existing_sheet_drawing_rel
            .as_ref()
            .map(|(_, target)| resolve_target(&sheet_meta.path, target));
        let has_existing_drawing_part = existing_drawing_part_path
            .as_deref()
            .is_some_and(|path| parts.contains_key(path));
        let missing_drawing_media = has_drawings
            && sheet.drawings.iter().any(|object| match &object.kind {
                DrawingObjectKind::Image { image_id } => {
                    !parts.contains_key(&format!("xl/media/{}", image_id.as_str()))
                }
                _ => false,
            });
        let original_drawings = doc.meta.drawings_snapshot.get(&sheet_meta.worksheet_id);
        let drawings_changed = if !has_drawings {
            false
        } else if let Some(orig) = original_drawings {
            orig != &sheet.drawings
        } else if has_existing_drawing_part && !missing_drawing_media {
            // Snapshot missing (e.g. workbook not loaded via `load_from_bytes`, or drawing parsing
            // failed on load). Fall back to parsing the existing drawing part to determine whether
            // we need to rewrite it.
            //
            // Best-effort: if parsing fails, assume unchanged so we preserve the original bytes.
            existing_drawing_part_path
                .as_deref()
                .and_then(|drawing_path| {
                    let mut tmp_workbook = formula_model::Workbook::new();
                    crate::drawings::DrawingPart::parse_from_parts(
                        sheet_index,
                        drawing_path,
                        &parts,
                        &mut tmp_workbook,
                    )
                    .ok()
                    .map(|part| part.objects != sheet.drawings)
                })
                .unwrap_or(false)
        } else {
            // No baseline; treat as changed so we emit a drawing part when needed.
            true
        };
        let drawings_need_emit =
            has_drawings && (!has_existing_drawing_part || missing_drawing_media || drawings_changed);
        let drawings_need_remove = if has_drawings {
            false
        } else if let Some(orig) = original_drawings {
            !orig.is_empty()
        } else if has_existing_drawing_part {
            // Snapshot missing; best-effort determine if the source drawing part was non-empty.
            existing_drawing_part_path
                .as_deref()
                .and_then(|drawing_path| {
                    let mut tmp_workbook = formula_model::Workbook::new();
                    crate::drawings::DrawingPart::parse_from_parts(
                        sheet_index,
                        drawing_path,
                        &parts,
                        &mut tmp_workbook,
                    )
                    .ok()
                    .map(|part| !part.objects.is_empty())
                })
                .unwrap_or(false)
        } else {
            false
        };

        if has_drawings && !drawings_need_emit {
            if let Some(drawing_part_path) = existing_drawing_part_path.as_deref() {
                ensure_drawing_part_content_types(&mut parts, drawing_part_path)?;
            }
        }
        if !is_new_sheet
            && !tab_color_changed
            && !merges_changed
            && !hyperlinks_changed
            && !views_changed
            && !sheet_format_changed
            && !cols_changed
            && !outline_cols_changed
            && !autofilter_changed
            && !data_validations_changed
            && !sheet_protection_changed
            && !drawings_need_emit
            && !drawings_need_remove
            && !conditional_formatting_changed
        {
            parts.insert(sheet_meta.path.clone(), sheet_xml_bytes);
            continue;
        }
        let mut sheet_xml = std::str::from_utf8(&sheet_xml_bytes)
            .map_err(|e| WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))?
            .to_string();

        // Apply sheet-level metadata updates.
        if is_new_sheet || tab_color_changed {
            sheet_xml = write_sheet_tab_color(&sheet_xml, sheet.tab_color.as_ref())?;
        }
        if is_new_sheet || views_changed {
            sheet_xml = update_sheet_views_xml(&sheet_xml, desired_views)?;
        }
        if is_new_sheet || sheet_format_changed {
            sheet_xml = update_sheet_format_pr_xml(&sheet_xml, desired_sheet_format)?;
        }
        if is_new_sheet || cols_changed || outline_cols_changed {
            let worksheet_prefix = crate::xml::worksheet_spreadsheetml_prefix(&sheet_xml)?;
            let cols_xml = render_cols(sheet, worksheet_prefix.as_deref(), &style_to_xf);
            sheet_xml = update_cols_xml(&sheet_xml, &cols_xml)?;
        }
        // Insert conditional formatting rules into preserved worksheet XML when the model contains
        // conditional formatting rules but the source XML didn't have any `<cfRule>` nodes (for
        // example, the sheet had no conditional formatting at all, or it contained an empty
        // `<conditionalFormatting/>` placeholder).
        if conditional_formatting_changed {
            let rules: Cow<'_, [CfRule]> = if sheet
                .conditional_formatting_rules
                .iter()
                .any(|rule| rule.dxf_id.is_some())
            {
                // `CfRule.dxf_id` indexes into the per-worksheet `conditional_formatting_dxfs`
                // vector. In SpreadsheetML, `cfRule/@dxfId` indexes into the *workbook-global*
                // `<dxfs>` table in `xl/styles.xml`, so remap local indices to the aggregated
                // workbook table.
                let mut owned = sheet.conditional_formatting_rules.clone();
                for rule in &mut owned {
                    rule.dxf_id = rule.dxf_id.and_then(|local| {
                        local_to_global_dxf.and_then(|map| map.get(local as usize).copied())
                    });
                }
                Cow::Owned(owned)
            } else {
                Cow::Borrowed(&sheet.conditional_formatting_rules)
            };

            sheet_xml = crate::conditional_formatting::update_worksheet_conditional_formatting_xml_with_seed(
                &sheet_xml,
                rules.as_ref(),
                sheet_meta.sheet_id as u128,
            )?;
        }
        if is_new_sheet || merges_changed {
            sheet_xml = crate::merge_cells::update_worksheet_xml(&sheet_xml, &current_merges)?;
        }
        if (is_new_sheet && !sheet.data_validations.is_empty()) || data_validations_changed {
            sheet_xml = data_validations::update_worksheet_data_validations_xml(
                &sheet_xml,
                &sheet.data_validations,
            )?;
        }
        if is_new_sheet || sheet_protection_changed {
            sheet_xml = update_sheet_protection_xml(&sheet_xml, &sheet.sheet_protection)?;
        }
        if (is_new_sheet && sheet.auto_filter.is_some()) || autofilter_changed {
            sheet_xml = crate::autofilter::write_worksheet_autofilter(
                &sheet_xml,
                sheet.auto_filter.as_ref(),
            )?;
        }
        if is_new_sheet || hyperlinks_changed {
            sheet_xml = crate::update_worksheet_xml(&sheet_xml, &current_hyperlinks)?;

            let updated_rels =
                crate::update_worksheet_relationships(rels_xml.as_deref(), &current_hyperlinks)?;
            match updated_rels {
                Some(xml) => {
                    parts.insert(rels_part.clone(), xml.into_bytes());
                }
                None => {
                    parts.remove(&rels_part);
                }
            }
        }

        if drawings_need_emit {
            // When possible, preserve an existing sheet->drawing relationship (stable `rId*` and
            // target path) by only rewriting the drawing part itself. Only synthesize a new
            // relationship when the sheet didn't have one.
            let existing_sheet_drawing_rel = match parts.get(&rels_part) {
                Some(bytes) => {
                    if let Some(rid) = orig_drawing_rel_id.as_deref() {
                        match relationship_target_by_id(bytes, rid)? {
                            Some(target) => Some((rid.to_string(), target)),
                            None => relationship_id_and_target_by_type(bytes, DRAWING_REL_TYPE)?,
                        }
                    } else {
                        relationship_id_and_target_by_type(bytes, DRAWING_REL_TYPE)?
                    }
                }
                None => None,
            };

            let (drawing_rel_id, drawing_part_path) = match existing_sheet_drawing_rel {
                Some((rid, target)) => (Some(rid), resolve_target(&sheet_meta.path, &target)),
                None => {
                    // If the expected per-sheet drawing part already exists (e.g. a workbook with
                    // a missing/corrupt worksheet `.rels`), reuse it instead of synthesizing a new
                    // `drawing{n}.xml` part and leaving the existing part orphaned.
                    let fallback = format!("xl/drawings/drawing{}.xml", sheet_index.saturating_add(1));
                    if parts.contains_key(&fallback) {
                        (None, fallback)
                    } else {
                        let n = next_drawing_part;
                        next_drawing_part += 1;
                        (None, format!("xl/drawings/drawing{n}.xml"))
                    }
                }
            };

            let _drawing_rel_id = match drawing_rel_id {
                Some(id) => {
                    // Ensure the worksheet XML has a `<drawing r:id="..."/>` pointer.
                    if worksheet_drawing_rel_id(&sheet_xml)?.as_deref() != Some(id.as_str()) {
                        sheet_xml = update_worksheet_drawing_xml(&sheet_xml, &id)?;
                    }
                    id
                }
                None => {
                    // Create a new worksheet relationship entry for the drawing part.
                    let rels_xml = parts
                        .get(&rels_part)
                        .and_then(|bytes| std::str::from_utf8(bytes).ok());
                    let mut rels = rels_xml
                        .map(crate::relationships::Relationships::from_xml)
                        .transpose()?
                        .unwrap_or_default();

                    let sheet_dir = sheet_meta
                        .path
                        .rsplit_once('/')
                        .map(|(dir, _)| dir)
                        .unwrap_or("");
                    let drawing_target = relative_target(sheet_dir, &drawing_part_path);

                    let drawing_rel_id = rels.next_r_id();
                    rels.push(crate::relationships::Relationship {
                        id: drawing_rel_id.clone(),
                        type_: DRAWING_REL_TYPE.to_string(),
                        target: drawing_target,
                        target_mode: None,
                    });
                    parts.insert(rels_part.clone(), rels.to_xml());

                    sheet_xml = update_worksheet_drawing_xml(&sheet_xml, &drawing_rel_id)?;
                    drawing_rel_id
                }
            };

            // Emit/update the drawing part + its relationships.
            let drawing_rels_path = crate::drawings::DrawingPart::rels_path_for(&drawing_part_path);
            let existing_drawing_rels_xml = parts
                .get(&drawing_rels_path)
                .and_then(|bytes| std::str::from_utf8(bytes).ok());
            let existing_drawing_xml = parts
                .get(&drawing_part_path)
                .and_then(|bytes| std::str::from_utf8(bytes).ok());
            let mut drawing_part = crate::drawings::DrawingPart::from_objects_with_existing_drawing_xml(
                sheet_index,
                drawing_part_path.clone(),
                sheet.drawings.clone(),
                existing_drawing_xml,
                existing_drawing_rels_xml,
            )?;
            drawing_part.write_into_parts(&mut parts, &doc.workbook)?;
            ensure_drawing_part_content_types(&mut parts, &drawing_part_path)?;
        } else if drawings_need_remove {
            // Remove the worksheet-level `<drawing>` pointer and its corresponding relationship
            // entry. We intentionally do not delete the underlying drawing parts/media because
            // they may contain other content that Formula doesn't model yet.
            let drawing_rid =
                worksheet_drawing_rel_id(&sheet_xml)?.or_else(|| orig_drawing_rel_id.clone());
            if worksheet_drawing_rel_id(&sheet_xml)?.is_some() {
                sheet_xml = remove_worksheet_drawing_xml(&sheet_xml)?;
            }

            if let Some(existing_rels) = parts
                .get(&rels_part)
                .and_then(|bytes| std::str::from_utf8(bytes).ok())
            {
                let rels = crate::relationships::Relationships::from_xml(existing_rels)?;
                let filtered: Vec<crate::relationships::Relationship> = rels
                    .iter()
                    .filter(|rel| {
                        let Some(drawing_rid) = drawing_rid.as_deref() else {
                            return true;
                        };
                        !(rel.id == drawing_rid && rel.type_ == DRAWING_REL_TYPE)
                    })
                    .cloned()
                    .collect();
                if filtered.is_empty() {
                    parts.remove(&rels_part);
                } else {
                    let rels = crate::relationships::Relationships::new(filtered);
                    parts.insert(rels_part.clone(), rels.to_xml());
                }
            }
        } else if has_drawings && existing_sheet_drawing_rel.is_some() {
            // We have an existing drawing relationship/part in the source package. Avoid touching
            // any `.rels` / `xl/drawings/*` parts unless we need to add new media. Still ensure the
            // `<drawing>` pointer is present in case other worksheet edits dropped it.
            if worksheet_drawing_rel_id(&sheet_xml)?.is_none() {
                if let Some((rid, _)) = existing_sheet_drawing_rel.as_ref() {
                    sheet_xml = update_worksheet_drawing_xml(&sheet_xml, rid)?;
                }
            }
        }

        parts.insert(sheet_meta.path.clone(), sheet_xml.into_bytes());
    }

    write_back_modified_comment_parts(
        doc,
        &sheet_plan.sheets,
        &sheet_plan.cell_meta_sheet_ids,
        &mut parts,
    );
    apply_print_settings_patches(doc, &sheet_plan.sheets, &mut parts)?;

    Ok(parts)
}

fn apply_print_settings_patches(
    doc: &XlsxDocument,
    sheets: &[SheetMeta],
    parts: &mut BTreeMap<String, Vec<u8>>,
) -> Result<(), WriteError> {
    if doc.workbook.print_settings == doc.meta.print_settings_snapshot {
        return Ok(());
    }

    fn to_write_error(err: crate::print::PrintError) -> WriteError {
        match err {
            crate::print::PrintError::Io(e) => WriteError::Io(e),
            crate::print::PrintError::Zip(e) => WriteError::Zip(e),
            crate::print::PrintError::Xml(e) => WriteError::Xml(e),
            crate::print::PrintError::XmlAttr(e) => WriteError::XmlAttr(e),
            crate::print::PrintError::Utf8(e) => {
                WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e))
            }
            crate::print::PrintError::InvalidA1(e) => {
                WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e))
            }
            crate::print::PrintError::MissingPart(part) => WriteError::Io(std::io::Error::new(
                std::io::ErrorKind::NotFound,
                format!("missing required xlsx part: {part}"),
            )),
            crate::print::PrintError::PartTooLarge { part, size, max } => {
                WriteError::Io(std::io::Error::new(
                    std::io::ErrorKind::InvalidData,
                    format!("xlsx part '{part}' is too large ({size} bytes, max {max})"),
                ))
            }
        }
    }

    fn model_range_to_cell_range(range: formula_model::Range) -> crate::print::CellRange {
        crate::print::CellRange {
            start_row: range.start.row.saturating_add(1),
            end_row: range.end.row.saturating_add(1),
            start_col: range.start.col.saturating_add(1),
            end_col: range.end.col.saturating_add(1),
        }
    }

    fn model_titles_to_print_titles(
        titles: formula_model::PrintTitles,
    ) -> crate::print::PrintTitles {
        crate::print::PrintTitles {
            repeat_rows: titles.repeat_rows.map(|r| crate::print::RowRange {
                start: r.start.saturating_add(1),
                end: r.end.saturating_add(1),
            }),
            repeat_cols: titles.repeat_cols.map(|c| crate::print::ColRange {
                start: c.start.saturating_add(1),
                end: c.end.saturating_add(1),
            }),
        }
    }

    fn model_page_setup_to_print(setup: formula_model::PageSetup) -> crate::print::PageSetup {
        crate::print::PageSetup {
            orientation: match setup.orientation {
                formula_model::Orientation::Portrait => crate::print::Orientation::Portrait,
                formula_model::Orientation::Landscape => crate::print::Orientation::Landscape,
            },
            paper_size: crate::print::PaperSize {
                code: setup.paper_size.code,
            },
            margins: crate::print::PageMargins {
                left: setup.margins.left,
                right: setup.margins.right,
                top: setup.margins.top,
                bottom: setup.margins.bottom,
                header: setup.margins.header,
                footer: setup.margins.footer,
            },
            scaling: match setup.scaling {
                formula_model::Scaling::Percent(pct) => crate::print::Scaling::Percent(pct),
                formula_model::Scaling::FitTo { width, height } => {
                    crate::print::Scaling::FitTo { width, height }
                }
            },
        }
    }

    fn model_breaks_to_print(
        breaks: formula_model::ManualPageBreaks,
    ) -> crate::print::ManualPageBreaks {
        crate::print::ManualPageBreaks {
            row_breaks_after: breaks
                .row_breaks_after
                .into_iter()
                .map(|v| v.saturating_add(1))
                .collect(),
            col_breaks_after: breaks
                .col_breaks_after
                .into_iter()
                .map(|v| v.saturating_add(1))
                .collect(),
        }
    }

    let mut sheet_name_to_meta: HashMap<String, (usize, String, String)> = HashMap::new();
    for (idx, meta) in sheets.iter().enumerate() {
        let Some(sheet) = doc.workbook.sheet(meta.worksheet_id) else {
            continue;
        };
        sheet_name_to_meta.insert(
            formula_model::sheet_name_casefold(&sheet.name),
            (idx, meta.path.clone(), sheet.name.clone()),
        );
    }

    // Track sheets with non-default print settings in either the original snapshot or the current model.
    let mut affected_sheets: HashSet<String> = HashSet::new();
    for sheet in &doc.workbook.print_settings.sheets {
        affected_sheets.insert(formula_model::sheet_name_casefold(&sheet.sheet_name));
    }
    for sheet in &doc.meta.print_settings_snapshot.sheets {
        affected_sheets.insert(formula_model::sheet_name_casefold(&sheet.sheet_name));
    }

    // Build per-sheet maps for diffing.
    let mut current_by_sheet: HashMap<String, &formula_model::SheetPrintSettings> = HashMap::new();
    for sheet in &doc.workbook.print_settings.sheets {
        current_by_sheet.insert(formula_model::sheet_name_casefold(&sheet.sheet_name), sheet);
    }
    let mut snapshot_by_sheet: HashMap<String, &formula_model::SheetPrintSettings> = HashMap::new();
    for sheet in &doc.meta.print_settings_snapshot.sheets {
        snapshot_by_sheet.insert(formula_model::sheet_name_casefold(&sheet.sheet_name), sheet);
    }

    let mut defined_name_edits: HashMap<(String, usize), crate::print::xlsx::DefinedNameEdit> =
        HashMap::new();

    for sheet_key in &affected_sheets {
        let Some((local_sheet_id, _path, sheet_name)) = sheet_name_to_meta.get(sheet_key) else {
            continue;
        };

        let current = current_by_sheet.get(sheet_key).copied();
        let snapshot = snapshot_by_sheet.get(sheet_key).copied();

        // Defined names: print area.
        let current_area = current.and_then(|s| s.print_area.as_deref());
        let snapshot_area = snapshot.and_then(|s| s.print_area.as_deref());
        if current_area != snapshot_area {
            let edit = match current_area {
                Some(ranges) => {
                    let ranges_1 = ranges
                        .iter()
                        .copied()
                        .map(model_range_to_cell_range)
                        .collect::<Vec<_>>();
                    let value = crate::print::format_print_area_defined_name(sheet_name, &ranges_1);
                    crate::print::xlsx::DefinedNameEdit::Set(value)
                }
                None => crate::print::xlsx::DefinedNameEdit::Remove,
            };
            defined_name_edits.insert(
                (formula_model::XLNM_PRINT_AREA.to_string(), *local_sheet_id),
                edit,
            );
        }

        // Defined names: print titles.
        let current_titles = current.and_then(|s| s.print_titles);
        let snapshot_titles = snapshot.and_then(|s| s.print_titles);
        if current_titles != snapshot_titles {
            let edit = match current_titles {
                Some(titles) => {
                    let titles_1 = model_titles_to_print_titles(titles);
                    let value =
                        crate::print::format_print_titles_defined_name(sheet_name, &titles_1);
                    crate::print::xlsx::DefinedNameEdit::Set(value)
                }
                None => crate::print::xlsx::DefinedNameEdit::Remove,
            };
            defined_name_edits.insert(
                (
                    formula_model::XLNM_PRINT_TITLES.to_string(),
                    *local_sheet_id,
                ),
                edit,
            );
        }
    }

    if !defined_name_edits.is_empty() {
        let workbook_xml = parts.get("xl/workbook.xml").ok_or_else(|| {
            WriteError::Io(std::io::Error::new(
                std::io::ErrorKind::NotFound,
                "missing xl/workbook.xml",
            ))
        })?;
        let updated = crate::print::xlsx::update_workbook_xml(workbook_xml, &defined_name_edits)
            .map_err(to_write_error)?;
        parts.insert("xl/workbook.xml".to_string(), updated);
    }

    // Worksheet-level print settings (page setup/margins/scaling + manual breaks).
    for sheet_key in affected_sheets {
        let Some((_local_sheet_id, path, sheet_name)) = sheet_name_to_meta.get(&sheet_key) else {
            continue;
        };

        let current = current_by_sheet.get(&sheet_key).copied();
        let snapshot = snapshot_by_sheet.get(&sheet_key).copied();

        let current_page_setup = current.map(|s| s.page_setup.clone()).unwrap_or_default();
        let snapshot_page_setup = snapshot.map(|s| s.page_setup.clone()).unwrap_or_default();
        let current_breaks = current
            .map(|s| s.manual_page_breaks.clone())
            .unwrap_or_default();
        let snapshot_breaks = snapshot
            .map(|s| s.manual_page_breaks.clone())
            .unwrap_or_default();

        if current_page_setup == snapshot_page_setup && current_breaks == snapshot_breaks {
            continue;
        }

        let Some(sheet_xml) = parts.get(path).map(|b| b.as_slice()) else {
            continue;
        };

        let settings = crate::print::SheetPrintSettings {
            sheet_name: sheet_name.clone(),
            print_area: current.and_then(|s| s.print_area.as_ref()).map(|ranges| {
                ranges
                    .iter()
                    .copied()
                    .map(model_range_to_cell_range)
                    .collect()
            }),
            print_titles: current
                .and_then(|s| s.print_titles)
                .map(model_titles_to_print_titles),
            page_setup: model_page_setup_to_print(current_page_setup),
            manual_page_breaks: model_breaks_to_print(current_breaks),
        };

        let updated = crate::print::xlsx::update_worksheet_xml(sheet_xml, &settings)
            .map_err(to_write_error)?;
        parts.insert(path.clone(), updated);
    }

    Ok(())
}

fn normalize_merge_ranges(ranges: impl Iterator<Item = Range>) -> Vec<Range> {
    let mut merges: Vec<Range> = ranges.filter(|r| !r.is_single_cell()).collect();
    merges.sort_by_key(|r| (r.start.row, r.start.col, r.end.row, r.end.col));
    merges
}

fn normalize_hyperlinks(links: &[Hyperlink]) -> Vec<Hyperlink> {
    let mut out = links.to_vec();
    out.sort_by(cmp_hyperlink);
    out
}

fn write_back_modified_comment_parts(
    doc: &XlsxDocument,
    sheets: &[SheetMeta],
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    parts: &mut BTreeMap<String, Vec<u8>>,
) {
    for sheet_meta in sheets {
        let Some(sheet) = doc.workbook.sheet(sheet_meta.worksheet_id) else {
            continue;
        };

        // WorksheetId values can differ between the in-memory workbook and the preserved metadata
        // (e.g. when the model is reconstructed from persisted state). Use the same workbook ->
        // meta sheet mapping as `cell_meta` and other round-trip metadata.
        let meta_sheet_id = cell_meta_sheet_ids
            .get(&sheet_meta.worksheet_id)
            .copied()
            .unwrap_or(sheet_meta.worksheet_id);

        let Some(part_names) = doc.meta.comment_part_names.get(&meta_sheet_id) else {
            continue;
        };
        let Some(snapshot) = doc.meta.comment_snapshot.get(&meta_sheet_id) else {
            continue;
        };

        let current = normalize_worksheet_comments(sheet);
        if &current == snapshot {
            continue;
        }

        // Rewrite only the comment XML parts that exist in the original workbook.
        let current_notes = filter_comments_by_kind(&current, CommentKind::Note);
        let snapshot_notes = filter_comments_by_kind(snapshot, CommentKind::Note);
        if current_notes != snapshot_notes {
            if let Some(path) = &part_names.legacy_comments {
                if parts.contains_key(path) {
                    parts.insert(
                        path.clone(),
                        crate::comments::legacy::write_comments_xml(&current_notes),
                    );
                }
            }
        }

        let current_threaded = filter_comments_by_kind(&current, CommentKind::Threaded);
        let snapshot_threaded = filter_comments_by_kind(snapshot, CommentKind::Threaded);
        if current_threaded != snapshot_threaded {
            if let Some(path) = &part_names.threaded_comments {
                if parts.contains_key(path) {
                    parts.insert(
                        path.clone(),
                        crate::comments::threaded::write_threaded_comments_xml(&current_threaded),
                    );
                }
            }
        }
    }
}

fn normalize_worksheet_comments(worksheet: &Worksheet) -> Vec<Comment> {
    let mut out: Vec<Comment> = worksheet
        .iter_comments()
        .map(|(_, comment)| comment.clone())
        .collect();
    out.sort_by(|a, b| {
        (
            a.cell_ref.row,
            a.cell_ref.col,
            comment_kind_rank(a.kind),
            &a.id,
        )
            .cmp(&(
                b.cell_ref.row,
                b.cell_ref.col,
                comment_kind_rank(b.kind),
                &b.id,
            ))
    });
    out
}

fn filter_comments_by_kind(comments: &[Comment], kind: CommentKind) -> Vec<Comment> {
    let mut out: Vec<Comment> = comments
        .iter()
        .filter(|comment| comment.kind == kind)
        .cloned()
        .collect();
    out.sort_by(|a, b| {
        (a.cell_ref.row, a.cell_ref.col, &a.id).cmp(&(b.cell_ref.row, b.cell_ref.col, &b.id))
    });
    out
}

fn comment_kind_rank(kind: CommentKind) -> u8 {
    match kind {
        CommentKind::Note => 0,
        CommentKind::Threaded => 1,
    }
}

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
struct DataValidationKey {
    ranges: Vec<(u32, u32, u32, u32)>,
    kind: u8,
    operator: Option<u8>,
    formula1: String,
    formula2: Option<String>,
    allow_blank: bool,
    show_input_message: bool,
    show_error_message: bool,
    show_drop_down: bool,
    prompt_title: Option<String>,
    prompt: Option<String>,
    error_style: Option<u8>,
    error_title: Option<String>,
    error: Option<String>,
}

fn normalize_dv_formula(formula: &str) -> String {
    let trimmed = formula.trim();
    // Match `read` normalization for data validation formulas:
    // - strip a single leading '='
    // - strip `_xlfn.` prefixes at function-call boundaries
    crate::formula_text::strip_xlfn_prefixes(trimmed.strip_prefix('=').unwrap_or(trimmed))
}

fn normalize_dv_ranges(ranges: &[Range]) -> Vec<(u32, u32, u32, u32)> {
    let mut out: Vec<(u32, u32, u32, u32)> = ranges
        .iter()
        .map(|r| (r.start.row, r.start.col, r.end.row, r.end.col))
        .collect();
    out.sort();
    out
}

fn dv_kind_rank(kind: DataValidationKind) -> u8 {
    match kind {
        DataValidationKind::Whole => 0,
        DataValidationKind::Decimal => 1,
        DataValidationKind::List => 2,
        DataValidationKind::Date => 3,
        DataValidationKind::Time => 4,
        DataValidationKind::TextLength => 5,
        DataValidationKind::Custom => 6,
    }
}

fn dv_operator_rank(op: DataValidationOperator) -> u8 {
    match op {
        DataValidationOperator::Between => 0,
        DataValidationOperator::NotBetween => 1,
        DataValidationOperator::Equal => 2,
        DataValidationOperator::NotEqual => 3,
        DataValidationOperator::GreaterThan => 4,
        DataValidationOperator::GreaterThanOrEqual => 5,
        DataValidationOperator::LessThan => 6,
        DataValidationOperator::LessThanOrEqual => 7,
    }
}

fn dv_error_style_rank(style: DataValidationErrorStyle) -> u8 {
    match style {
        DataValidationErrorStyle::Stop => 0,
        DataValidationErrorStyle::Warning => 1,
        DataValidationErrorStyle::Information => 2,
    }
}

fn normalize_data_validation(
    validation: &formula_model::DataValidation,
    ranges: &[Range],
) -> DataValidationKey {
    let ranges = normalize_dv_ranges(ranges);
    let formula1 = normalize_dv_formula(&validation.formula1);
    let formula2 = validation
        .formula2
        .as_deref()
        .map(normalize_dv_formula)
        .filter(|f| !f.is_empty());
    let show_drop_down = if validation.kind == DataValidationKind::List {
        validation.show_drop_down
    } else {
        // `showDropDown` is only meaningful for list validations. Our reader canonicalizes
        // non-list rules to `show_drop_down=false`; mirror that here so no-op saves do not
        // spuriously rewrite `<dataValidations>` when round-tripping.
        false
    };

    let (prompt_title, prompt) = match &validation.input_message {
        Some(msg) => {
            let title = msg
                .title
                .as_ref()
                .map(|s| s.trim().to_string())
                .filter(|s| !s.is_empty());
            let body = msg
                .body
                .as_ref()
                .map(|s| s.trim().to_string())
                .filter(|s| !s.is_empty());
            if title.is_none() && body.is_none() {
                (None, None)
            } else {
                (title, body)
            }
        }
        None => (None, None),
    };

    let (error_style, error_title, error) = match &validation.error_alert {
        Some(alert) => {
            let title = alert
                .title
                .as_ref()
                .map(|s| s.trim().to_string())
                .filter(|s| !s.is_empty());
            let body = alert
                .body
                .as_ref()
                .map(|s| s.trim().to_string())
                .filter(|s| !s.is_empty());
            if title.is_none() && body.is_none() && alert.style == DataValidationErrorStyle::Stop {
                (None, None, None)
            } else {
                (Some(dv_error_style_rank(alert.style)), title, body)
            }
        }
        None => (None, None, None),
    };

    DataValidationKey {
        ranges,
        kind: dv_kind_rank(validation.kind),
        operator: validation.operator.map(dv_operator_rank),
        formula1,
        formula2,
        allow_blank: validation.allow_blank,
        show_input_message: validation.show_input_message,
        show_error_message: validation.show_error_message,
        show_drop_down,
        prompt_title,
        prompt,
        error_style,
        error_title,
        error,
    }
}

fn normalize_data_validations(validations: &[DataValidationAssignment]) -> Vec<DataValidationKey> {
    let mut out: Vec<DataValidationKey> = validations
        .iter()
        .filter(|dv| !dv.ranges.is_empty())
        .map(|dv| normalize_data_validation(&dv.validation, &dv.ranges))
        .collect();
    out.sort();
    out
}

fn normalize_parsed_data_validations(
    validations: &[crate::data_validations::ParsedDataValidation],
) -> Vec<DataValidationKey> {
    let mut out: Vec<DataValidationKey> = validations
        .iter()
        .filter(|dv| !dv.ranges.is_empty())
        .map(|dv| normalize_data_validation(&dv.validation, &dv.ranges))
        .collect();
    out.sort();
    out
}

#[derive(Clone, Debug, PartialEq)]
struct SheetViewSelectionSettings {
    active_cell: CellRef,
    sqref: String,
}

#[derive(Clone, Debug, PartialEq)]
struct SheetViewSettings {
    /// Zoom scale as an integer percentage (100 = 100%).
    zoom_scale: u32,

    /// Frozen pane row count (top).
    frozen_rows: u32,
    /// Frozen pane column count (left).
    frozen_cols: u32,

    /// Horizontal split position (non-freeze panes).
    x_split: Option<f32>,
    /// Vertical split position (non-freeze panes).
    y_split: Option<f32>,

    /// Top-left visible cell for the bottom-right pane (`pane/@topLeftCell`).
    top_left_cell: Option<CellRef>,

    show_grid_lines: bool,
    show_headings: bool,
    show_zeros: bool,

    selection: Option<SheetViewSelectionSettings>,
}

impl Default for SheetViewSettings {
    fn default() -> Self {
        Self {
            zoom_scale: 100,
            frozen_rows: 0,
            frozen_cols: 0,
            x_split: None,
            y_split: None,
            top_left_cell: None,
            show_grid_lines: true,
            show_headings: true,
            show_zeros: true,
            selection: None,
        }
    }
}

impl SheetViewSettings {
    fn from_sheet(sheet: &Worksheet) -> Self {
        let view_is_default = sheet.view == formula_model::SheetView::default();

        let zoom = if view_is_default {
            sheet.zoom
        } else {
            sheet.view.zoom
        };

        // Excel stores this as an integer percentage (`zoomScale="120"`).
        let mut zoom_scale = (zoom * 100.0).round() as i64;
        zoom_scale = zoom_scale.max(10).min(400);

        let (frozen_rows, frozen_cols) = if view_is_default {
            (sheet.frozen_rows, sheet.frozen_cols)
        } else {
            (sheet.view.pane.frozen_rows, sheet.view.pane.frozen_cols)
        };

        // We only serialize split offsets when there is no frozen pane state.
        let (x_split, y_split) = if view_is_default || frozen_rows > 0 || frozen_cols > 0 {
            (None, None)
        } else {
            (sheet.view.pane.x_split, sheet.view.pane.y_split)
        };

        let mut top_left_cell = if view_is_default {
            None
        } else {
            sheet.view.pane.top_left_cell
        };
        if top_left_cell.is_none() && (frozen_rows > 0 || frozen_cols > 0) {
            top_left_cell = Some(CellRef::new(frozen_rows, frozen_cols));
        }

        let (show_grid_lines, show_headings, show_zeros, selection) = if view_is_default {
            (true, true, true, None)
        } else {
            let selection = sheet.view.selection.as_ref().map(|sel| SheetViewSelectionSettings {
                active_cell: sel.active_cell,
                sqref: sel.sqref(),
            });
            (
                sheet.view.show_grid_lines,
                sheet.view.show_headings,
                sheet.view.show_zeros,
                selection,
            )
        };

        Self {
            zoom_scale: zoom_scale as u32,
            frozen_rows,
            frozen_cols,
            x_split,
            y_split,
            top_left_cell,
            show_grid_lines,
            show_headings,
            show_zeros,
            selection,
        }
    }

    fn is_default(&self) -> bool {
        self.zoom_scale == 100
            && self.frozen_rows == 0
            && self.frozen_cols == 0
            && self.x_split.is_none()
            && self.y_split.is_none()
            && self.top_left_cell.is_none()
            && self.show_grid_lines
            && self.show_headings
            && self.show_zeros
            && self.selection.is_none()
    }
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
struct SheetFormatSettings {
    default_col_width: Option<f32>,
    default_row_height: Option<f32>,
    base_col_width: Option<u16>,
}

impl SheetFormatSettings {
    fn from_sheet(sheet: &Worksheet) -> Self {
        Self {
            default_col_width: sheet.default_col_width,
            default_row_height: sheet.default_row_height,
            base_col_width: sheet.base_col_width,
        }
    }

    fn is_default(self) -> bool {
        self.default_col_width.is_none()
            && self.default_row_height.is_none()
            && self.base_col_width.is_none()
    }
}

fn parse_sheet_view_settings(xml: &str) -> Result<SheetViewSettings, WriteError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut settings = SheetViewSettings::default();
    let mut in_sheet_view = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) if e.local_name().as_ref() == b"sheetView" => {
                in_sheet_view = true;
                for attr in e.attributes() {
                    let attr = attr?;
                    let val = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"zoomScale" => {
                            if let Ok(scale) = val.parse::<u32>() {
                                settings.zoom_scale = scale;
                            }
                        }
                        b"showGridLines" => settings.show_grid_lines = parse_xml_bool(&val),
                        b"showHeadings" | b"showRowColHeaders" => {
                            settings.show_headings = parse_xml_bool(&val)
                        }
                        b"showZeros" => settings.show_zeros = parse_xml_bool(&val),
                        _ => {}
                    }
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == b"sheetView" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    let val = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"zoomScale" => {
                            if let Ok(scale) = val.parse::<u32>() {
                                settings.zoom_scale = scale;
                            }
                        }
                        b"showGridLines" => settings.show_grid_lines = parse_xml_bool(&val),
                        b"showHeadings" | b"showRowColHeaders" => {
                            settings.show_headings = parse_xml_bool(&val)
                        }
                        b"showZeros" => settings.show_zeros = parse_xml_bool(&val),
                        _ => {}
                    }
                }
            }
            Event::End(e) if e.local_name().as_ref() == b"sheetView" => {
                in_sheet_view = false;
                drop(e);
            }
            Event::Start(e) | Event::Empty(e) if in_sheet_view && e.local_name().as_ref() == b"pane" => {
                let mut state: Option<String> = None;
                let mut x_split: Option<String> = None;
                let mut y_split: Option<String> = None;
                let mut top_left_cell: Option<CellRef> = None;

                for attr in e.attributes() {
                    let attr = attr?;
                    let val = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"state" => state = Some(val),
                        b"xSplit" => x_split = Some(val),
                        b"ySplit" => y_split = Some(val),
                        b"topLeftCell" => {
                            top_left_cell = CellRef::from_a1(&val).ok();
                        }
                        _ => {}
                    }
                }

                match state.as_deref() {
                    Some("frozen") | Some("frozenSplit") => {
                        settings.frozen_cols = x_split.as_deref().and_then(|v| v.parse().ok()).unwrap_or(0);
                        settings.frozen_rows = y_split.as_deref().and_then(|v| v.parse().ok()).unwrap_or(0);
                        settings.x_split = None;
                        settings.y_split = None;
                        settings.top_left_cell = top_left_cell.or_else(|| {
                            if settings.frozen_rows > 0 || settings.frozen_cols > 0 {
                                Some(CellRef::new(settings.frozen_rows, settings.frozen_cols))
                            } else {
                                None
                            }
                        });
                    }
                    Some("split") => {
                        settings.x_split = x_split.as_deref().and_then(|v| v.parse().ok());
                        settings.y_split = y_split.as_deref().and_then(|v| v.parse().ok());
                        settings.top_left_cell = top_left_cell;
                    }
                    _ => {
                        // Best-effort: if no explicit state is set, treat the presence of split
                        // offsets as a split pane.
                        settings.x_split = x_split.as_deref().and_then(|v| v.parse().ok());
                        settings.y_split = y_split.as_deref().and_then(|v| v.parse().ok());
                        settings.top_left_cell = top_left_cell;
                    }
                }
            }
            Event::Start(e) | Event::Empty(e)
                if in_sheet_view && e.local_name().as_ref() == b"selection" =>
            {
                let mut active_cell: Option<CellRef> = None;
                let mut sqref: Option<String> = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    let val = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"activeCell" => active_cell = CellRef::from_a1(&val).ok(),
                        b"sqref" => sqref = Some(val),
                        _ => {}
                    }
                }
                if let Some(active_cell) = active_cell {
                    let sqref = sqref.unwrap_or_else(|| active_cell.to_a1());
                    settings.selection = Some(SheetViewSelectionSettings { active_cell, sqref });
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(settings)
}

fn parse_sheet_format_settings(xml: &str) -> Result<SheetFormatSettings, WriteError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut settings = SheetFormatSettings::default();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"sheetFormatPr" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    let val = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"defaultColWidth" => settings.default_col_width = val.parse::<f32>().ok(),
                        b"defaultRowHeight" => {
                            settings.default_row_height = val.parse::<f32>().ok()
                        }
                        b"baseColWidth" => settings.base_col_width = val.parse::<u16>().ok(),
                        _ => {}
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(settings)
}

fn f32_eq(a: f32, b: f32) -> bool {
    (a - b).abs() <= 0.0001
}

fn sheet_format_needs_patch(desired: SheetFormatSettings, original: SheetFormatSettings) -> bool {
    // When callers don't set any sheet format defaults in the model, preserve any existing
    // `<sheetFormatPr>` element byte-for-byte for round-trip fidelity.
    if desired.is_default() {
        return false;
    }

    if let Some(want) = desired.default_col_width {
        match original.default_col_width {
            Some(orig) if f32_eq(orig, want) => {}
            _ => return true,
        }
    }
    if let Some(want) = desired.default_row_height {
        match original.default_row_height {
            Some(orig) if f32_eq(orig, want) => {}
            _ => return true,
        }
    }
    if let Some(want) = desired.base_col_width {
        match original.base_col_width {
            Some(orig) if orig == want => {}
            _ => return true,
        }
    }

    false
}

fn render_sheet_views_section(views: SheetViewSettings, prefix: Option<&str>) -> String {
    if views.is_default() {
        return String::new();
    }

    let sheet_views_tag = crate::xml::prefixed_tag(prefix, "sheetViews");
    let sheet_view_tag = crate::xml::prefixed_tag(prefix, "sheetView");
    let pane_tag = crate::xml::prefixed_tag(prefix, "pane");
    let selection_tag = crate::xml::prefixed_tag(prefix, "selection");

    let mut out = String::new();
    out.push('<');
    out.push_str(&sheet_views_tag);
    out.push('>');
    out.push('<');
    out.push_str(&sheet_view_tag);
    out.push_str(" workbookViewId=\"0\"");
    if !views.show_grid_lines {
        out.push_str(r#" showGridLines="0""#);
    }
    if !views.show_headings {
        out.push_str(r#" showHeadings="0""#);
    }
    if !views.show_zeros {
        out.push_str(r#" showZeros="0""#);
    }
    if views.zoom_scale != 100 {
        out.push_str(&format!(r#" zoomScale="{}""#, views.zoom_scale));
    }

    let has_pane = views.frozen_rows > 0
        || views.frozen_cols > 0
        || views.x_split.is_some()
        || views.y_split.is_some();
    let has_selection = views.selection.is_some();

    if !has_pane && !has_selection {
        out.push_str("/></");
        out.push_str(&sheet_views_tag);
        out.push('>');
        return out;
    }

    out.push('>');

    if has_pane {
        out.push('<');
        out.push_str(&pane_tag);

        let frozen = views.frozen_rows > 0 || views.frozen_cols > 0;
        let x_present = if frozen {
            views.frozen_cols > 0
        } else {
            views.x_split.is_some()
        };
        let y_present = if frozen {
            views.frozen_rows > 0
        } else {
            views.y_split.is_some()
        };

        let state = if frozen { "frozen" } else { "split" };
        out.push_str(&format!(r#" state="{state}""#));

        if frozen {
            if views.frozen_cols > 0 {
                out.push_str(&format!(r#" xSplit="{}""#, views.frozen_cols));
            }
            if views.frozen_rows > 0 {
                out.push_str(&format!(r#" ySplit="{}""#, views.frozen_rows));
            }
        } else {
            if let Some(x_split) = views.x_split {
                out.push_str(&format!(r#" xSplit="{}""#, format_sheet_view_split(x_split)));
            }
            if let Some(y_split) = views.y_split {
                out.push_str(&format!(r#" ySplit="{}""#, format_sheet_view_split(y_split)));
            }
        }

        if let Some(top_left_cell) = views.top_left_cell.or_else(|| {
            if frozen {
                Some(CellRef::new(views.frozen_rows, views.frozen_cols))
            } else {
                None
            }
        }) {
            let top_left = top_left_cell.to_a1();
            out.push_str(&format!(r#" topLeftCell="{}""#, escape_attr(&top_left)));
        }

        let active_pane = if x_present && y_present {
            "bottomRight"
        } else if y_present {
            "bottomLeft"
        } else {
            "topRight"
        };
        out.push_str(&format!(r#" activePane="{active_pane}"/>"#));
    }

    if let Some(selection) = &views.selection {
        out.push('<');
        out.push_str(&selection_tag);
        out.push_str(&format!(
            r#" activeCell="{}" sqref="{}"/>"#,
            escape_attr(&selection.active_cell.to_a1()),
            escape_attr(&selection.sqref)
        ));
    }

    out.push_str("</");
    out.push_str(&sheet_view_tag);
    out.push_str("></");
    out.push_str(&sheet_views_tag);
    out.push('>');
    out
}

fn format_sheet_view_split(val: f32) -> String {
    if !val.is_finite() {
        return "0".to_string();
    }
    // For deterministic fixtures, trim trailing zeros (e.g. `1.0` -> `1`).
    let mut s = format!("{val}");
    if s.contains('.') {
        while s.ends_with('0') {
            s.pop();
        }
        if s.ends_with('.') {
            s.pop();
        }
    }
    if s.is_empty() {
        "0".to_string()
    } else {
        s
    }
}

fn update_sheet_views_xml(sheet_xml: &str, views: SheetViewSettings) -> Result<String, WriteError> {
    let worksheet_prefix = crate::xml::worksheet_spreadsheetml_prefix(sheet_xml)?;
    let new_section = render_sheet_views_section(views, worksheet_prefix.as_deref());

    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(false);

    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut skip_depth: usize = 0;
    let mut replaced = false;
    let mut inserted = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            _ if skip_depth > 0 => match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth = skip_depth.saturating_sub(1),
                Event::Empty(_) => {}
                _ => {}
            },
            Event::Start(ref e) if e.local_name().as_ref() == b"sheetViews" => {
                replaced = true;
                if !new_section.is_empty() {
                    writer.get_mut().extend_from_slice(new_section.as_bytes());
                }
                skip_depth = 1;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"sheetViews" => {
                replaced = true;
                if !new_section.is_empty() {
                    writer.get_mut().extend_from_slice(new_section.as_bytes());
                }
            }
            Event::Start(ref e)
                if e.local_name().as_ref() == b"sheetFormatPr"
                    || e.local_name().as_ref() == b"cols"
                    || e.local_name().as_ref() == b"sheetData" =>
            {
                if !replaced && !inserted && !new_section.is_empty() {
                    writer.get_mut().extend_from_slice(new_section.as_bytes());
                    inserted = true;
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e)
                if e.local_name().as_ref() == b"sheetFormatPr"
                    || e.local_name().as_ref() == b"cols"
                    || e.local_name().as_ref() == b"sheetData" =>
            {
                if !replaced && !inserted && !new_section.is_empty() {
                    writer.get_mut().extend_from_slice(new_section.as_bytes());
                    inserted = true;
                }
                writer.write_event(Event::Empty(e.to_owned()))?;
            }
            _ => {
                writer.write_event(event.to_owned())?;
            }
        }
        buf.clear();
    }

    String::from_utf8(writer.into_inner())
        .map_err(|e| WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))
}

fn render_sheet_format_pr(settings: SheetFormatSettings, prefix: Option<&str>) -> String {
    if settings.is_default() {
        return String::new();
    }
    let tag = crate::xml::prefixed_tag(prefix, "sheetFormatPr");
    let mut out = String::new();
    out.push('<');
    out.push_str(&tag);

    if let Some(base) = settings.base_col_width {
        out.push_str(&format!(r#" baseColWidth="{base}""#));
    }
    if let Some(width) = settings.default_col_width {
        // `f32::to_string()` prints `-0.0` as `-0`; normalize for XML stability.
        let width = if width == 0.0 { 0.0 } else { width };
        out.push_str(&format!(r#" defaultColWidth="{width}""#));
    }
    if let Some(height) = settings.default_row_height {
        // `f32::to_string()` prints `-0.0` as `-0`; normalize for XML stability.
        let height = if height == 0.0 { 0.0 } else { height };
        out.push_str(&format!(r#" defaultRowHeight="{height}""#));
    }

    out.push_str("/>");
    out
}

fn update_sheet_format_pr_xml(sheet_xml: &str, settings: SheetFormatSettings) -> Result<String, WriteError> {
    // If the model doesn't explicitly set any defaults, preserve any existing `<sheetFormatPr>`
    // element as-is for round-trip fidelity.
    let worksheet_prefix = crate::xml::worksheet_spreadsheetml_prefix(sheet_xml)?;
    let insert_section = render_sheet_format_pr(settings, worksheet_prefix.as_deref());
    if insert_section.is_empty() {
        return Ok(sheet_xml.to_string());
    }

    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(false);

    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut replaced = false;
    let mut inserted = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e) if e.local_name().as_ref() == b"sheetFormatPr" => {
                replaced = true;
                write_sheet_format_pr_element(&mut writer, e, settings, false)?;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"sheetFormatPr" => {
                replaced = true;
                write_sheet_format_pr_element(&mut writer, e, settings, true)?;
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if e.local_name().as_ref() == b"cols"
                    || e.local_name().as_ref() == b"sheetData"
                    || e.local_name().as_ref() == b"sheetProtection"
                    || e.local_name().as_ref() == b"autoFilter"
                    || e.local_name().as_ref() == b"mergeCells"
                    || e.local_name().as_ref() == b"hyperlinks"
                    || e.local_name().as_ref() == b"tableParts"
                    || e.local_name().as_ref() == b"drawing"
                    || e.local_name().as_ref() == b"extLst" =>
            {
                if !replaced && !inserted {
                    writer.get_mut().extend_from_slice(insert_section.as_bytes());
                    inserted = true;
                }
                writer.write_event(event.to_owned())?;
            }
            Event::End(ref e) if e.local_name().as_ref() == b"worksheet" => {
                if !replaced && !inserted {
                    writer.get_mut().extend_from_slice(insert_section.as_bytes());
                    inserted = true;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            _ => {
                writer.write_event(event.to_owned())?;
            }
        }
        buf.clear();
    }

    String::from_utf8(writer.into_inner())
        .map_err(|e| WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))
}

fn write_sheet_format_pr_element(
    writer: &mut Writer<Vec<u8>>,
    e: &quick_xml::events::BytesStart<'_>,
    settings: SheetFormatSettings,
    is_empty: bool,
) -> Result<(), WriteError> {
    writer.get_mut().push(b'<');
    writer.get_mut().extend_from_slice(e.name().as_ref());

    let mut wrote_default_row_height = false;
    let mut wrote_default_col_width = false;
    let mut wrote_base_col_width = false;

    for attr in e.attributes() {
        let attr = attr?;
        writer.get_mut().push(b' ');
        writer.get_mut().extend_from_slice(attr.key.as_ref());
        writer.get_mut().extend_from_slice(b"=\"");

        match attr.key.as_ref() {
            b"defaultRowHeight" if settings.default_row_height.is_some() => {
                wrote_default_row_height = true;
                writer.get_mut().extend_from_slice(
                    {
                        let height = settings.default_row_height.expect("checked is_some");
                        let height = if height == 0.0 { 0.0 } else { height };
                        height.to_string()
                    }
                    .as_bytes(),
                );
            }
            b"defaultColWidth" if settings.default_col_width.is_some() => {
                wrote_default_col_width = true;
                writer.get_mut().extend_from_slice(
                    {
                        let width = settings.default_col_width.expect("checked is_some");
                        let width = if width == 0.0 { 0.0 } else { width };
                        width.to_string()
                    }
                    .as_bytes(),
                );
            }
            b"baseColWidth" if settings.base_col_width.is_some() => {
                wrote_base_col_width = true;
                writer.get_mut().extend_from_slice(
                    settings
                        .base_col_width
                        .expect("checked is_some")
                        .to_string()
                        .as_bytes(),
                );
            }
            _ => {
                writer.get_mut().extend_from_slice(
                    escape_attr(&attr.unescape_value()?.into_owned()).as_bytes(),
                );
            }
        }

        writer.get_mut().push(b'"');
    }

    if let Some(base) = settings.base_col_width {
        if !wrote_base_col_width {
            writer.get_mut().extend_from_slice(br#" baseColWidth=""#);
            writer.get_mut().extend_from_slice(base.to_string().as_bytes());
            writer.get_mut().push(b'"');
        }
    }
    if let Some(width) = settings.default_col_width {
        if !wrote_default_col_width {
            writer.get_mut().extend_from_slice(br#" defaultColWidth=""#);
            let width = if width == 0.0 { 0.0 } else { width };
            writer.get_mut().extend_from_slice(width.to_string().as_bytes());
            writer.get_mut().push(b'"');
        }
    }
    if let Some(height) = settings.default_row_height {
        if !wrote_default_row_height {
            writer.get_mut().extend_from_slice(br#" defaultRowHeight=""#);
            let height = if height == 0.0 { 0.0 } else { height };
            writer.get_mut().extend_from_slice(height.to_string().as_bytes());
            writer.get_mut().push(b'"');
        }
    }

    if is_empty {
        writer.get_mut().extend_from_slice(b"/>");
    } else {
        writer.get_mut().push(b'>');
    }

    Ok(())
}
fn parse_xml_bool(val: &str) -> bool {
    val == "1" || val.eq_ignore_ascii_case("true")
}

fn parse_xml_u16_hex(val: &str) -> Option<u16> {
    u16::from_str_radix(val.trim(), 16).ok()
}

fn parse_sheet_protection(xml: &str) -> Result<Option<SheetProtection>, WriteError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"sheetProtection" => {
                // Mirror the parser behavior from `read`: the presence of the element implies
                // protection is enabled unless `sheet="0"` overrides it.
                let mut protection = SheetProtection::default();
                protection.enabled = true;
                for attr in e.attributes() {
                    let attr = attr?;
                    let val = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"sheet" => protection.enabled = parse_xml_bool(&val),
                        b"selectLockedCells" => {
                            protection.select_locked_cells = parse_xml_bool(&val)
                        }
                        b"selectUnlockedCells" => {
                            protection.select_unlocked_cells = parse_xml_bool(&val)
                        }
                        b"formatCells" => protection.format_cells = parse_xml_bool(&val),
                        b"formatColumns" => protection.format_columns = parse_xml_bool(&val),
                        b"formatRows" => protection.format_rows = parse_xml_bool(&val),
                        b"insertColumns" => protection.insert_columns = parse_xml_bool(&val),
                        b"insertRows" => protection.insert_rows = parse_xml_bool(&val),
                        b"insertHyperlinks" => protection.insert_hyperlinks = parse_xml_bool(&val),
                        b"deleteColumns" => protection.delete_columns = parse_xml_bool(&val),
                        b"deleteRows" => protection.delete_rows = parse_xml_bool(&val),
                        b"sort" => protection.sort = parse_xml_bool(&val),
                        b"autoFilter" => protection.auto_filter = parse_xml_bool(&val),
                        b"pivotTables" => protection.pivot_tables = parse_xml_bool(&val),
                        // Inverted "protected" flags.
                        b"objects" => protection.edit_objects = !parse_xml_bool(&val),
                        b"scenarios" => protection.edit_scenarios = !parse_xml_bool(&val),
                        b"password" => {
                            protection.password_hash =
                                parse_xml_u16_hex(&val).filter(|hash| *hash != 0);
                        }
                        _ => {}
                    }
                }
                return Ok(Some(protection));
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(None)
}

fn worksheet_has_data_validations(xml: &str) -> Result<bool, WriteError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"dataValidations" => {
                return Ok(true);
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(false)
}

fn update_sheet_protection_xml(
    sheet_xml: &str,
    protection: &SheetProtection,
) -> Result<String, WriteError> {
    let worksheet_prefix = crate::xml::worksheet_spreadsheetml_prefix(sheet_xml)?;
    let new_section = render_sheet_protection(protection, worksheet_prefix.as_deref());

    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(false);

    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut skip_depth: usize = 0;
    let mut sheet_calc_pr_depth: usize = 0;
    let mut replaced = false;
    let mut inserted = false;
    let mut pending_insert_after_sheet_data = false;

    fn append_attr_raw(writer: &mut Vec<u8>, key: &[u8], val: &str) {
        writer.push(b' ');
        writer.extend_from_slice(key);
        writer.extend_from_slice(b"=\"");
        writer.extend_from_slice(escape_attr(val).as_bytes());
        writer.push(b'"');
    }

    fn append_attr_bool(writer: &mut Vec<u8>, key: &[u8], original: Option<&str>, desired: bool) {
        writer.push(b' ');
        writer.extend_from_slice(key);
        writer.extend_from_slice(b"=\"");
        if let Some(original) = original {
            if parse_xml_bool(original) == desired {
                writer.extend_from_slice(escape_attr(original).as_bytes());
            } else {
                writer.extend_from_slice(if desired { b"1" } else { b"0" });
            }
        } else {
            writer.extend_from_slice(if desired { b"1" } else { b"0" });
        }
        writer.push(b'"');
    }

    fn append_attr_password(
        writer: &mut Vec<u8>,
        key: &[u8],
        original: Option<&str>,
        desired: Option<u16>,
    ) {
        let Some(original) = original else {
            if let Some(hash) = desired {
                writer.push(b' ');
                writer.extend_from_slice(key);
                writer.extend_from_slice(b"=\"");
                writer.extend_from_slice(format!("{:04X}", hash).as_bytes());
                writer.push(b'"');
            }
            return;
        };

        let original_hash = parse_xml_u16_hex(original).filter(|hash| *hash != 0);
        match desired {
            Some(hash) => {
                writer.push(b' ');
                writer.extend_from_slice(key);
                writer.extend_from_slice(b"=\"");
                if original_hash == Some(hash) {
                    writer.extend_from_slice(escape_attr(original).as_bytes());
                } else {
                    writer.extend_from_slice(format!("{:04X}", hash).as_bytes());
                }
                writer.push(b'"');
            }
            None => {
                // Preserve semantically-empty password hashes (e.g. `0000`) but drop real ones.
                if original_hash.is_none() {
                    append_attr_raw(writer, key, original);
                }
            }
        }
    }

    fn write_patched_sheet_protection_start(
        writer: &mut Writer<Vec<u8>>,
        e: &quick_xml::events::BytesStart<'_>,
        protection: &SheetProtection,
        is_empty: bool,
    ) -> Result<(), WriteError> {
        let mut wrote_sheet = false;
        let mut wrote_select_locked_cells = false;
        let mut wrote_select_unlocked_cells = false;
        let mut wrote_format_cells = false;
        let mut wrote_format_columns = false;
        let mut wrote_format_rows = false;
        let mut wrote_insert_columns = false;
        let mut wrote_insert_rows = false;
        let mut wrote_insert_hyperlinks = false;
        let mut wrote_delete_columns = false;
        let mut wrote_delete_rows = false;
        let mut wrote_sort = false;
        let mut wrote_auto_filter = false;
        let mut wrote_pivot_tables = false;
        let mut wrote_objects = false;
        let mut wrote_scenarios = false;
        let mut wrote_password = false;

        let tag = e.name();
        let tag = tag.as_ref();
        let buf = writer.get_mut();
        buf.extend_from_slice(b"<");
        buf.extend_from_slice(tag);

        for attr in e.attributes().with_checks(false) {
            let attr = attr?;
            let key = attr.key.as_ref();
            let val = attr.unescape_value()?.into_owned();

            match key {
                b"sheet" => {
                    wrote_sheet = true;
                    append_attr_bool(buf, key, Some(val.as_str()), protection.enabled);
                }
                b"selectLockedCells" => {
                    wrote_select_locked_cells = true;
                    append_attr_bool(buf, key, Some(val.as_str()), protection.select_locked_cells);
                }
                b"selectUnlockedCells" => {
                    wrote_select_unlocked_cells = true;
                    append_attr_bool(
                        buf,
                        key,
                        Some(val.as_str()),
                        protection.select_unlocked_cells,
                    );
                }
                b"formatCells" => {
                    wrote_format_cells = true;
                    append_attr_bool(buf, key, Some(val.as_str()), protection.format_cells);
                }
                b"formatColumns" => {
                    wrote_format_columns = true;
                    append_attr_bool(buf, key, Some(val.as_str()), protection.format_columns);
                }
                b"formatRows" => {
                    wrote_format_rows = true;
                    append_attr_bool(buf, key, Some(val.as_str()), protection.format_rows);
                }
                b"insertColumns" => {
                    wrote_insert_columns = true;
                    append_attr_bool(buf, key, Some(val.as_str()), protection.insert_columns);
                }
                b"insertRows" => {
                    wrote_insert_rows = true;
                    append_attr_bool(buf, key, Some(val.as_str()), protection.insert_rows);
                }
                b"insertHyperlinks" => {
                    wrote_insert_hyperlinks = true;
                    append_attr_bool(buf, key, Some(val.as_str()), protection.insert_hyperlinks);
                }
                b"deleteColumns" => {
                    wrote_delete_columns = true;
                    append_attr_bool(buf, key, Some(val.as_str()), protection.delete_columns);
                }
                b"deleteRows" => {
                    wrote_delete_rows = true;
                    append_attr_bool(buf, key, Some(val.as_str()), protection.delete_rows);
                }
                b"sort" => {
                    wrote_sort = true;
                    append_attr_bool(buf, key, Some(val.as_str()), protection.sort);
                }
                b"autoFilter" => {
                    wrote_auto_filter = true;
                    append_attr_bool(buf, key, Some(val.as_str()), protection.auto_filter);
                }
                b"pivotTables" => {
                    wrote_pivot_tables = true;
                    append_attr_bool(buf, key, Some(val.as_str()), protection.pivot_tables);
                }
                // Inverted "protected" flags.
                b"objects" => {
                    wrote_objects = true;
                    append_attr_bool(buf, key, Some(val.as_str()), !protection.edit_objects);
                }
                b"scenarios" => {
                    wrote_scenarios = true;
                    append_attr_bool(buf, key, Some(val.as_str()), !protection.edit_scenarios);
                }
                b"password" => {
                    wrote_password = true;
                    append_attr_password(buf, key, Some(val.as_str()), protection.password_hash);
                }
                _ => {
                    append_attr_raw(buf, key, val.as_str());
                }
            }
        }

        // Only add missing allow-list attributes when the desired value differs from the implicit
        // defaults used by our parser when the attribute is absent.
        let mut implied_defaults = SheetProtection::default();
        implied_defaults.enabled = true;

        if protection.enabled != implied_defaults.enabled && !wrote_sheet {
            append_attr_bool(buf, b"sheet", None, protection.enabled);
        }
        if protection.select_locked_cells != implied_defaults.select_locked_cells
            && !wrote_select_locked_cells
        {
            append_attr_bool(
                buf,
                b"selectLockedCells",
                None,
                protection.select_locked_cells,
            );
        }
        if protection.select_unlocked_cells != implied_defaults.select_unlocked_cells
            && !wrote_select_unlocked_cells
        {
            append_attr_bool(
                buf,
                b"selectUnlockedCells",
                None,
                protection.select_unlocked_cells,
            );
        }
        if protection.format_cells != implied_defaults.format_cells && !wrote_format_cells {
            append_attr_bool(buf, b"formatCells", None, protection.format_cells);
        }
        if protection.format_columns != implied_defaults.format_columns && !wrote_format_columns {
            append_attr_bool(buf, b"formatColumns", None, protection.format_columns);
        }
        if protection.format_rows != implied_defaults.format_rows && !wrote_format_rows {
            append_attr_bool(buf, b"formatRows", None, protection.format_rows);
        }
        if protection.insert_columns != implied_defaults.insert_columns && !wrote_insert_columns {
            append_attr_bool(buf, b"insertColumns", None, protection.insert_columns);
        }
        if protection.insert_rows != implied_defaults.insert_rows && !wrote_insert_rows {
            append_attr_bool(buf, b"insertRows", None, protection.insert_rows);
        }
        if protection.insert_hyperlinks != implied_defaults.insert_hyperlinks
            && !wrote_insert_hyperlinks
        {
            append_attr_bool(buf, b"insertHyperlinks", None, protection.insert_hyperlinks);
        }
        if protection.delete_columns != implied_defaults.delete_columns && !wrote_delete_columns {
            append_attr_bool(buf, b"deleteColumns", None, protection.delete_columns);
        }
        if protection.delete_rows != implied_defaults.delete_rows && !wrote_delete_rows {
            append_attr_bool(buf, b"deleteRows", None, protection.delete_rows);
        }
        if protection.sort != implied_defaults.sort && !wrote_sort {
            append_attr_bool(buf, b"sort", None, protection.sort);
        }
        if protection.auto_filter != implied_defaults.auto_filter && !wrote_auto_filter {
            append_attr_bool(buf, b"autoFilter", None, protection.auto_filter);
        }
        if protection.pivot_tables != implied_defaults.pivot_tables && !wrote_pivot_tables {
            append_attr_bool(buf, b"pivotTables", None, protection.pivot_tables);
        }
        if protection.edit_objects != implied_defaults.edit_objects && !wrote_objects {
            append_attr_bool(buf, b"objects", None, !protection.edit_objects);
        }
        if protection.edit_scenarios != implied_defaults.edit_scenarios && !wrote_scenarios {
            append_attr_bool(buf, b"scenarios", None, !protection.edit_scenarios);
        }
        if protection.password_hash.is_some() && !wrote_password {
            append_attr_password(buf, b"password", None, protection.password_hash);
        }

        if is_empty {
            buf.extend_from_slice(b"/>");
        } else {
            buf.push(b'>');
        }

        Ok(())
    }

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            _ if skip_depth > 0 => match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth = skip_depth.saturating_sub(1),
                Event::Empty(_) => {}
                _ => {}
            },
            _ if sheet_calc_pr_depth > 0 => match event {
                Event::Start(ref e) => {
                    sheet_calc_pr_depth = sheet_calc_pr_depth.saturating_add(1);
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
                Event::Empty(ref e) => {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
                Event::End(ref e) => {
                    sheet_calc_pr_depth = sheet_calc_pr_depth.saturating_sub(1);
                    writer.write_event(Event::End(e.to_owned()))?;
                    if sheet_calc_pr_depth == 0 && !replaced && !inserted && !new_section.is_empty()
                    {
                        writer.get_mut().extend_from_slice(new_section.as_bytes());
                        inserted = true;
                    }
                }
                _ => {
                    writer.write_event(event.to_owned())?;
                }
            },
            Event::Start(ref e) if e.local_name().as_ref() == b"sheetProtection" => {
                replaced = true;
                pending_insert_after_sheet_data = false;
                if protection.enabled {
                    write_patched_sheet_protection_start(&mut writer, e, protection, false)?;
                } else {
                    // When protection is disabled, remove any `<sheetProtection>` element (even if
                    // it used `sheet="0"` and includes nested content).
                    skip_depth = 1;
                }
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"sheetProtection" => {
                replaced = true;
                pending_insert_after_sheet_data = false;
                if protection.enabled {
                    write_patched_sheet_protection_start(&mut writer, e, protection, true)?;
                } else {
                    // Drop the element entirely.
                }
            }
            Event::End(ref e) if e.local_name().as_ref() == b"sheetData" => {
                writer.write_event(Event::End(e.to_owned()))?;
                if !replaced && !inserted && !new_section.is_empty() {
                    pending_insert_after_sheet_data = true;
                }
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"sheetData" => {
                writer.write_event(Event::Empty(e.to_owned()))?;
                if !replaced && !inserted && !new_section.is_empty() {
                    pending_insert_after_sheet_data = true;
                }
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"sheetCalcPr" => {
                if pending_insert_after_sheet_data
                    && !replaced
                    && !inserted
                    && !new_section.is_empty()
                {
                    // Schema order is `sheetData`, `sheetCalcPr`, then `sheetProtection`.
                    // If `sheetCalcPr` exists, insert after it instead of immediately after
                    // `sheetData`.
                    pending_insert_after_sheet_data = false;
                    sheet_calc_pr_depth = 1;
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"sheetCalcPr" => {
                writer.write_event(Event::Empty(e.to_owned()))?;
                if pending_insert_after_sheet_data
                    && !replaced
                    && !inserted
                    && !new_section.is_empty()
                {
                    pending_insert_after_sheet_data = false;
                    writer.get_mut().extend_from_slice(new_section.as_bytes());
                    inserted = true;
                }
            }
            Event::Start(ref e) | Event::Empty(ref e) if pending_insert_after_sheet_data => {
                // We just finished `<sheetData>` and the next element is *not* `<sheetCalcPr>`,
                // meaning we can insert `<sheetProtection>` immediately after sheetData while
                // preserving SpreadsheetML element ordering.
                if !replaced && !inserted && !new_section.is_empty() {
                    pending_insert_after_sheet_data = false;
                    writer.get_mut().extend_from_slice(new_section.as_bytes());
                    inserted = true;
                }
                writer.write_event(event.to_owned())?;
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if e.local_name().as_ref() == b"autoFilter" =>
            {
                // If `sheetData` is missing (unexpected), fall back to inserting before autoFilter
                // so the worksheet remains schema-valid.
                if !replaced && !inserted && !new_section.is_empty() {
                    writer.get_mut().extend_from_slice(new_section.as_bytes());
                    inserted = true;
                }
                writer.write_event(event.to_owned())?;
            }
            Event::End(ref e) if e.local_name().as_ref() == b"worksheet" => {
                if pending_insert_after_sheet_data
                    && !replaced
                    && !inserted
                    && !new_section.is_empty()
                {
                    pending_insert_after_sheet_data = false;
                    writer.get_mut().extend_from_slice(new_section.as_bytes());
                    inserted = true;
                } else if !replaced && !inserted && !new_section.is_empty() {
                    writer.get_mut().extend_from_slice(new_section.as_bytes());
                    inserted = true;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            _ => {
                writer.write_event(event.to_owned())?;
            }
        }
        buf.clear();
    }

    String::from_utf8(writer.into_inner())
        .map_err(|e| WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))
}

fn parse_col_properties(
    xml: &str,
    styles_editor: &XlsxStylesEditor,
) -> Result<BTreeMap<u32, formula_model::ColProperties>, WriteError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_cols = false;
    let mut map: BTreeMap<u32, formula_model::ColProperties> = BTreeMap::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) if e.local_name().as_ref() == b"cols" => in_cols = true,
            Event::End(e) if e.local_name().as_ref() == b"cols" => {
                in_cols = false;
                drop(e);
            }
            Event::Start(e) | Event::Empty(e) if in_cols && e.local_name().as_ref() == b"col" => {
                let mut min: Option<u32> = None;
                let mut max: Option<u32> = None;
                let mut width: Option<f32> = None;
                let mut custom_width: Option<bool> = None;
                let mut hidden = false;
                let mut style: Option<u32> = None;
                let mut custom_format: Option<bool> = None;

                for attr in e.attributes() {
                    let attr = attr?;
                    let val = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"min" => min = val.parse().ok(),
                        b"max" => max = val.parse().ok(),
                        b"width" => width = val.parse().ok(),
                        b"customWidth" => custom_width = Some(parse_xml_bool(&val)),
                        b"hidden" => hidden = parse_xml_bool(&val),
                        b"style" => style = val.parse().ok(),
                        b"customFormat" => custom_format = Some(parse_xml_bool(&val)),
                        _ => {}
                    }
                }

                let Some(min) = min else { continue };
                let max = max.unwrap_or(min).min(formula_model::EXCEL_MAX_COLS);
                if min == 0 || max == 0 || min > formula_model::EXCEL_MAX_COLS {
                    continue;
                }

                let width = if custom_width == Some(false) {
                    None
                } else {
                    width
                };

                let clear_style = custom_format == Some(false);
                let style_id = if clear_style {
                    None
                } else {
                    style
                        .map(|xf_index| styles_editor.style_id_for_xf(xf_index))
                        .filter(|style_id| *style_id != 0)
                };

                for idx_1_based in min..=max {
                    let col = idx_1_based - 1;
                    if col >= formula_model::EXCEL_MAX_COLS {
                        continue;
                    }
                    if width.is_none() && !hidden && style_id.is_none() {
                        continue;
                    }
                    let entry = map.entry(col).or_default();
                    if let Some(width) = width {
                        entry.width = Some(width);
                    }
                    if hidden {
                        entry.hidden = true;
                    }
                    if clear_style {
                        entry.style_id = None;
                    } else if let Some(style_id) = style_id {
                        entry.style_id = Some(style_id);
                    }
                    if entry.width.is_none() && !entry.hidden && entry.style_id.is_none() {
                        map.remove(&col);
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(map)
}

fn parse_cols_xml_props(
    xml: &str,
    styles_editor: &XlsxStylesEditor,
) -> Result<BTreeMap<u32, ColXmlProps>, WriteError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    let mut in_cols = false;
    let mut map: BTreeMap<u32, ColXmlProps> = BTreeMap::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) if e.local_name().as_ref() == b"cols" => in_cols = true,
            Event::End(e) if e.local_name().as_ref() == b"cols" => {
                in_cols = false;
                drop(e);
            }
            Event::Start(e) | Event::Empty(e) if in_cols && e.local_name().as_ref() == b"col" => {
                let mut min: Option<u32> = None;
                let mut max: Option<u32> = None;
                let mut width: Option<f32> = None;
                let mut custom_width: Option<bool> = None;
                let mut style_xf: Option<u32> = None;
                let mut custom_format: Option<bool> = None;
                let mut hidden = false;
                let mut outline_level: u8 = 0;
                let mut collapsed = false;

                for attr in e.attributes() {
                    let attr = attr?;
                    let val = attr.unescape_value()?.into_owned();
                    match attr.key.as_ref() {
                        b"min" => min = val.parse().ok(),
                        b"max" => max = val.parse().ok(),
                        b"width" => width = val.parse().ok(),
                        b"customWidth" => custom_width = Some(parse_xml_bool(&val)),
                        b"hidden" => hidden = parse_xml_bool(&val),
                        b"style" => style_xf = val.parse().ok(),
                        b"customFormat" => custom_format = Some(parse_xml_bool(&val)),
                        b"outlineLevel" => outline_level = val.parse().unwrap_or(0),
                        b"collapsed" => collapsed = parse_xml_bool(&val),
                        _ => {}
                    }
                }

                let Some(min) = min else { continue };
                let max = max.unwrap_or(min).min(formula_model::EXCEL_MAX_COLS);
                if min == 0 || max == 0 || min > formula_model::EXCEL_MAX_COLS {
                    continue;
                }

                let width = if custom_width == Some(false) {
                    None
                } else {
                    width
                };

                let clear_style = custom_format == Some(false);
                let style_xf = if clear_style {
                    None
                } else {
                    style_xf.and_then(|xf| {
                        // Treat references to the default style as "no override" so no-op saves do
                        // not spuriously rewrite `<cols>` when a producer emits redundant
                        // `style="0"` (or any other xf index that maps to the default style).
                        //
                        // NOTE: Some producers place custom xfs at index 0; in that case
                        // `style_id_for_xf(0)` will be non-zero and we preserve `style="0"`.
                        (styles_editor.style_id_for_xf(xf) != 0).then_some(xf)
                    })
                };

                if width.is_none()
                    && !hidden
                    && style_xf.is_none()
                    && outline_level == 0
                    && !collapsed
                    && !clear_style
                {
                    continue;
                }

                for col_1_based in min..=max {
                    if col_1_based == 0 || col_1_based > formula_model::EXCEL_MAX_COLS {
                        continue;
                    }
                    let entry = map.entry(col_1_based).or_insert_with(|| ColXmlProps {
                        width: None,
                        hidden: false,
                        outline_level: 0,
                        collapsed: false,
                        style_xf: None,
                    });
                    if let Some(width) = width {
                        entry.width = Some(width);
                    }
                    entry.hidden |= hidden;
                    entry.outline_level = entry.outline_level.max(outline_level);
                    entry.collapsed |= collapsed;
                    if clear_style {
                        entry.style_xf = None;
                    } else if let Some(style_xf) = style_xf {
                        entry.style_xf = Some(style_xf);
                    }
                    if entry.width.is_none()
                        && !entry.hidden
                        && entry.outline_level == 0
                        && !entry.collapsed
                        && entry.style_xf.is_none()
                    {
                        map.remove(&col_1_based);
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(map)
}

fn update_cols_xml(sheet_xml: &str, cols_section: &str) -> Result<String, WriteError> {
    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(false);

    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut skip_depth: usize = 0;
    let mut replaced = false;
    let mut inserted = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            _ if skip_depth > 0 => match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth = skip_depth.saturating_sub(1),
                Event::Empty(_) => {}
                _ => {}
            },
            Event::Start(ref e) if e.local_name().as_ref() == b"cols" => {
                replaced = true;
                if !cols_section.is_empty() {
                    writer.get_mut().extend_from_slice(cols_section.as_bytes());
                }
                skip_depth = 1;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"cols" => {
                replaced = true;
                if !cols_section.is_empty() {
                    writer.get_mut().extend_from_slice(cols_section.as_bytes());
                }
            }
            Event::Start(ref e) if e.local_name().as_ref() == b"sheetData" => {
                if !replaced && !inserted && !cols_section.is_empty() {
                    writer.get_mut().extend_from_slice(cols_section.as_bytes());
                    inserted = true;
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"sheetData" => {
                if !replaced && !inserted && !cols_section.is_empty() {
                    writer.get_mut().extend_from_slice(cols_section.as_bytes());
                    inserted = true;
                }
                writer.write_event(Event::Empty(e.to_owned()))?;
            }
            _ => {
                writer.write_event(event.to_owned())?;
            }
        }
        buf.clear();
    }

    String::from_utf8(writer.into_inner())
        .map_err(|e| WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))
}

fn worksheet_drawing_rel_id(sheet_xml: &str) -> Result<Option<String>, WriteError> {
    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) | Event::Empty(e) if e.local_name().as_ref() == b"drawing" => {
                for attr in e.attributes() {
                    let attr = attr?;
                    if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"id") {
                        return Ok(Some(attr.unescape_value()?.into_owned()));
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }
    Ok(None)
}

fn insert_drawing_before_tag(name: &[u8]) -> bool {
    matches!(
        name,
        // Elements that come after <drawing> in the SpreadsheetML schema.
        b"drawingHF"
            | b"picture"
            | b"oleObjects"
            | b"controls"
            | b"webPublishItems"
            | b"tableParts"
            | b"extLst"
    )
}

fn remove_worksheet_drawing_xml(sheet_xml: &str) -> Result<String, WriteError> {
    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(false);

    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut skip_depth: usize = 0;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            _ if skip_depth > 0 => match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth = skip_depth.saturating_sub(1),
                Event::Empty(_) => {}
                _ => {}
            },
            Event::Start(ref e) if e.local_name().as_ref() == b"drawing" => {
                skip_depth = 1;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"drawing" => {
                // Drop the element.
            }
            _ => {
                writer.write_event(event.to_owned())?;
            }
        }
        buf.clear();
    }

    String::from_utf8(writer.into_inner())
        .map_err(|e| WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))
}

fn update_worksheet_drawing_xml(
    sheet_xml: &str,
    drawing_rel_id: &str,
) -> Result<String, WriteError> {
    let worksheet_prefix = crate::xml::worksheet_spreadsheetml_prefix(sheet_xml)?;
    let drawing_tag = prefixed_tag(worksheet_prefix.as_deref(), "drawing");

    let mut reader = Reader::from_str(sheet_xml);
    reader.config_mut().trim_text(false);

    let mut writer = Writer::new(Vec::new());
    let mut buf = Vec::new();

    let mut skip_depth: usize = 0;
    let mut replaced = false;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            _ if skip_depth > 0 => match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth = skip_depth.saturating_sub(1),
                Event::Empty(_) => {}
                _ => {}
            },
            Event::Start(ref e) if e.local_name().as_ref() == b"drawing" => {
                replaced = true;
                write_drawing_block(&mut writer, drawing_rel_id, &drawing_tag)?;
                skip_depth = 1;
            }
            Event::Empty(ref e) if e.local_name().as_ref() == b"drawing" => {
                replaced = true;
                write_drawing_block(&mut writer, drawing_rel_id, &drawing_tag)?;
            }
            Event::Start(ref e) | Event::Empty(ref e)
                if !replaced && insert_drawing_before_tag(e.local_name().as_ref()) =>
            {
                write_drawing_block(&mut writer, drawing_rel_id, &drawing_tag)?;
                replaced = true;
                writer.write_event(event.to_owned())?;
            }
            Event::End(ref e) if e.local_name().as_ref() == b"worksheet" => {
                if !replaced {
                    write_drawing_block(&mut writer, drawing_rel_id, &drawing_tag)?;
                    replaced = true;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            _ => {
                writer.write_event(event.to_owned())?;
            }
        }
        buf.clear();
    }

    String::from_utf8(writer.into_inner())
        .map_err(|e| WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))
}

fn write_drawing_block<W: std::io::Write>(
    writer: &mut Writer<W>,
    drawing_rel_id: &str,
    drawing_tag: &str,
) -> Result<(), WriteError> {
    let mut elem = quick_xml::events::BytesStart::new(drawing_tag);
    // Declare the `r:` prefix locally so we can always emit `r:id`.
    elem.push_attribute(("xmlns:r", crate::xml::OFFICE_RELATIONSHIPS_NS));
    elem.push_attribute(("r:id", drawing_rel_id));
    writer.write_event(Event::Empty(elem))?;
    Ok(())
}

fn assign_hyperlink_rel_ids(hyperlinks: &[Hyperlink], rels_xml: Option<&str>) -> Vec<Hyperlink> {
    let mut next_id = rels_xml.map(next_relationship_id_in_xml).unwrap_or(1);
    let mut used: HashSet<String> = hyperlinks.iter().filter_map(|l| l.rel_id.clone()).collect();

    hyperlinks
        .iter()
        .cloned()
        .map(|mut link| {
            match link.target {
                HyperlinkTarget::ExternalUrl { .. } | HyperlinkTarget::Email { .. } => {
                    if link.rel_id.is_none() {
                        loop {
                            let id = format!("rId{next_id}");
                            next_id += 1;
                            if used.insert(id.clone()) {
                                link.rel_id = Some(id);
                                break;
                            }
                        }
                    }
                }
                HyperlinkTarget::Internal { .. } => {}
            }
            link
        })
        .collect()
}

fn cmp_hyperlink(a: &Hyperlink, b: &Hyperlink) -> std::cmp::Ordering {
    use std::cmp::Ordering;

    a.range
        .start
        .row
        .cmp(&b.range.start.row)
        .then(a.range.start.col.cmp(&b.range.start.col))
        .then(a.range.end.row.cmp(&b.range.end.row))
        .then(a.range.end.col.cmp(&b.range.end.col))
        .then(cmp_hyperlink_target(&a.target, &b.target))
        .then(a.display.cmp(&b.display))
        .then(a.tooltip.cmp(&b.tooltip))
        .then(a.rel_id.cmp(&b.rel_id))
        // Keep the ordering total even if we add new fields later.
        .then_with(|| Ordering::Equal)
}

fn cmp_hyperlink_target(a: &HyperlinkTarget, b: &HyperlinkTarget) -> std::cmp::Ordering {
    use std::cmp::Ordering;

    fn rank(target: &HyperlinkTarget) -> u8 {
        match target {
            HyperlinkTarget::ExternalUrl { .. } => 0,
            HyperlinkTarget::Email { .. } => 1,
            HyperlinkTarget::Internal { .. } => 2,
        }
    }

    let rank_cmp = rank(a).cmp(&rank(b));
    if rank_cmp != Ordering::Equal {
        return rank_cmp;
    }

    match (a, b) {
        (HyperlinkTarget::ExternalUrl { uri: a }, HyperlinkTarget::ExternalUrl { uri: b }) => {
            a.cmp(b)
        }
        (HyperlinkTarget::Email { uri: a }, HyperlinkTarget::Email { uri: b }) => a.cmp(b),
        (
            HyperlinkTarget::Internal {
                sheet: a_sheet,
                cell: a_cell,
            },
            HyperlinkTarget::Internal {
                sheet: b_sheet,
                cell: b_cell,
            },
        ) => a_sheet
            .cmp(b_sheet)
            .then(a_cell.row.cmp(&b_cell.row))
            .then(a_cell.col.cmp(&b_cell.col)),
        _ => Ordering::Equal,
    }
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
struct SharedStringKey {
    text: String,
    runs: Vec<SharedStringRunKey>,
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
struct SharedStringRunKey {
    start: usize,
    end: usize,
    style: SharedStringRunStyleKey,
}

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
struct SharedStringRunStyleKey {
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<u8>,
    color: Option<u32>,
    font: Option<String>,
    size_100pt: Option<u16>,
}

impl SharedStringKey {
    fn plain(text: &str) -> Self {
        Self {
            text: text.to_string(),
            runs: Vec::new(),
        }
    }

    fn from_rich_text(rich: &RichText) -> Self {
        let runs = rich
            .runs
            .iter()
            .map(|run| SharedStringRunKey {
                start: run.start,
                end: run.end,
                style: SharedStringRunStyleKey {
                    bold: run.style.bold,
                    italic: run.style.italic,
                    underline: run.style.underline.map(underline_key),
                    color: run.style.color.and_then(|c| c.argb()),
                    font: run.style.font.clone(),
                    size_100pt: run.style.size_100pt,
                },
            })
            .collect();
        Self {
            text: rich.text.clone(),
            runs,
        }
    }
}

fn underline_key(underline: Underline) -> u8 {
    match underline {
        Underline::None => 0,
        Underline::Single => 1,
        Underline::Double => 2,
        Underline::SingleAccounting => 3,
        Underline::DoubleAccounting => 4,
    }
}

fn lookup_cell_meta<'a>(
    doc: &'a XlsxDocument,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    worksheet_id: WorksheetId,
    cell_ref: CellRef,
) -> Option<&'a crate::CellMeta> {
    let meta_sheet_id = cell_meta_sheet_ids
        .get(&worksheet_id)
        .copied()
        .unwrap_or(worksheet_id);
    doc.meta
        .cell_meta
        .get(&(meta_sheet_id, cell_ref))
        .or_else(|| {
            if meta_sheet_id != worksheet_id {
                doc.meta.cell_meta.get(&(worksheet_id, cell_ref))
            } else {
                None
            }
        })
}

fn build_shared_strings_xml(
    doc: &XlsxDocument,
    sheets: &[SheetMeta],
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    original_xml: Option<&[u8]>,
) -> Result<(Vec<u8>, HashMap<SharedStringKey, u32>), WriteError> {
    let mut table: Vec<RichText> = doc.shared_strings.clone();
    let mut lookup: HashMap<SharedStringKey, u32> = HashMap::new();
    for (idx, rich) in table.iter().enumerate() {
        lookup
            .entry(SharedStringKey::from_rich_text(rich))
            .or_insert(idx as u32);
    }

    let mut ref_count: u32 = 0;

    for sheet_meta in sheets {
        let sheet = match doc.workbook.sheet(sheet_meta.worksheet_id) {
            Some(s) => s,
            None => continue,
        };

        let mut cells: Vec<(CellRef, &formula_model::Cell)> = sheet.iter_cells().collect();
        cells.sort_by_key(|(r, _)| (r.row, r.col));
        for (cell_ref, cell) in cells {
            let meta =
                lookup_cell_meta(doc, cell_meta_sheet_ids, sheet_meta.worksheet_id, cell_ref);
            let kind = effective_value_kind(meta, cell);
            let CellValueKind::SharedString { .. } = kind else {
                continue;
            };

            match &cell.value {
                CellValue::String(text) => {
                    ref_count += 1;
                    if meta
                        .and_then(|m| m.value_kind.clone())
                        .and_then(|k| match k {
                            CellValueKind::SharedString { index } => Some(index),
                            _ => None,
                        })
                        .and_then(|idx| doc.shared_strings.get(idx as usize))
                        .map(|rt| rt.text.as_str() == text.as_str())
                        .unwrap_or(false)
                    {
                        // Preserve the original shared string index even if the entry
                        // contains rich formatting.
                        continue;
                    }

                    let key = SharedStringKey::plain(text);
                    if !lookup.contains_key(&key) {
                        let new_index = table.len() as u32;
                        table.push(RichText::new(text.clone()));
                        lookup.insert(key, new_index);
                    }
                }
                CellValue::Entity(entity) => {
                    let text = entity.display_value.as_str();
                    ref_count += 1;
                    if meta
                        .and_then(|m| m.value_kind.clone())
                        .and_then(|k| match k {
                            CellValueKind::SharedString { index } => Some(index),
                            _ => None,
                        })
                        .and_then(|idx| doc.shared_strings.get(idx as usize))
                        .map(|rt| rt.text.as_str() == text)
                        .unwrap_or(false)
                    {
                        continue;
                    }

                    let key = SharedStringKey::plain(text);
                    if !lookup.contains_key(&key) {
                        let new_index = table.len() as u32;
                        table.push(RichText::new(text.to_string()));
                        lookup.insert(key, new_index);
                    }
                }
                CellValue::Record(record) => {
                    let text = record.to_string();
                    let text_str = text.as_str();
                    ref_count += 1;
                    if meta
                        .and_then(|m| m.value_kind.clone())
                        .and_then(|k| match k {
                            CellValueKind::SharedString { index } => Some(index),
                            _ => None,
                        })
                        .and_then(|idx| doc.shared_strings.get(idx as usize))
                        .map(|rt| rt.text.as_str() == text_str)
                        .unwrap_or(false)
                    {
                        continue;
                    }

                    let key = SharedStringKey::plain(text_str);
                    if !lookup.contains_key(&key) {
                        let new_index = table.len() as u32;
                        table.push(RichText::new(text));
                        lookup.insert(key, new_index);
                    }
                }
                CellValue::Image(image) => {
                    let Some(text) = image.alt_text.as_deref().filter(|s| !s.is_empty()) else {
                        continue;
                    };
                    ref_count += 1;
                    if meta
                        .and_then(|m| m.value_kind.clone())
                        .and_then(|k| match k {
                            CellValueKind::SharedString { index } => Some(index),
                            _ => None,
                        })
                        .and_then(|idx| doc.shared_strings.get(idx as usize))
                        .map(|rt| rt.text.as_str() == text)
                        .unwrap_or(false)
                    {
                        continue;
                    }

                    let key = SharedStringKey::plain(text);
                    if !lookup.contains_key(&key) {
                        let new_index = table.len() as u32;
                        table.push(RichText::new(text.to_string()));
                        lookup.insert(key, new_index);
                    }
                }
                CellValue::RichText(rich) => {
                    ref_count += 1;
                    if meta
                        .and_then(|m| m.value_kind.clone())
                        .and_then(|k| match k {
                            CellValueKind::SharedString { index } => Some(index),
                            _ => None,
                        })
                        .and_then(|idx| doc.shared_strings.get(idx as usize))
                        .map(|rt| rt == rich)
                        .unwrap_or(false)
                    {
                        continue;
                    }

                    let key = SharedStringKey::from_rich_text(rich);
                    if !lookup.contains_key(&key) {
                        let new_index = table.len() as u32;
                        table.push(rich.clone());
                        lookup.insert(key, new_index);
                    }
                }
                _ => {
                    // Non-string values ignore shared string bookkeeping.
                }
            }
        }

        // Include shared strings from columnar-backed worksheets.
        if let Some((_, rows, cols)) = sheet.columnar_table_extent() {
            if let Some(columnar) = sheet.columnar_table() {
                let columnar = columnar.as_ref();
                for row in 0..rows {
                    for col in 0..cols {
                        if let ColumnarValue::String(s) = columnar.get_cell(row, col) {
                            ref_count += 1;
                            let text = s.as_ref();
                            let key = SharedStringKey::plain(text);
                            if !lookup.contains_key(&key) {
                                let new_index = table.len() as u32;
                                table.push(RichText::new(text.to_string()));
                                lookup.insert(key, new_index);
                            }
                        }
                    }
                }
            }
        }
    }

    // If we started from an existing `sharedStrings.xml` and we didn't add any new entries,
    // preserve the original bytes byte-for-byte to avoid dropping unsupported substructures
    // (phonetic runs, extensions, mc:AlternateContent, etc.).
    if let Some(original_xml) = original_xml {
        if table.len() == doc.shared_strings.len() {
            return Ok((original_xml.to_vec(), lookup));
        }

        let mut editor = SharedStringsEditor::parse(original_xml).map_err(|e| {
            WriteError::Io(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!("sharedStrings.xml parse error: {e}"),
            ))
        })?;

        // Append only the newly created entries (do not rewrite existing `<si>` blocks).
        for rich in &table[doc.shared_strings.len()..] {
            if rich.runs.is_empty() {
                editor.get_or_insert_plain(&rich.text);
            } else {
                editor.get_or_insert_rich(rich);
            }
        }

        let patched = editor.to_xml_bytes(Some(ref_count)).map_err(|e| {
            WriteError::Io(std::io::Error::new(
                std::io::ErrorKind::InvalidData,
                format!("sharedStrings.xml write error: {e}"),
            ))
        })?;

        return Ok((patched, lookup));
    }

    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#);
    xml.push_str(&format!(
        r#" count="{ref_count}" uniqueCount="{}">"#,
        table.len()
    ));
    for rich in &table {
        xml.push_str("<si>");
        if rich.runs.is_empty() {
            write_shared_string_t(&mut xml, &rich.text);
        } else {
            for run in &rich.runs {
                xml.push_str("<r>");
                if !run.style.is_empty() {
                    xml.push_str("<rPr>");
                    write_shared_string_rpr(&mut xml, &run.style);
                    xml.push_str("</rPr>");
                }
                let segment = rich.slice_run_text(run);
                write_shared_string_t(&mut xml, segment);
                xml.push_str("</r>");
            }
        }
        xml.push_str("</si>");
    }
    xml.push_str("</sst>");

    Ok((xml.into_bytes(), lookup))
}

fn write_shared_string_t(xml: &mut String, text: &str) {
    xml.push_str("<t");
    if needs_space_preserve(text) {
        xml.push_str(r#" xml:space="preserve""#);
    }
    xml.push('>');
    xml.push_str(&escape_text(text));
    xml.push_str("</t>");
}

fn write_shared_string_rpr(xml: &mut String, style: &formula_model::rich_text::RichTextRunStyle) {
    if let Some(font) = &style.font {
        xml.push_str(r#"<rFont val=""#);
        xml.push_str(&escape_attr(font));
        xml.push_str(r#""/>"#);
    }

    if let Some(size_100pt) = style.size_100pt {
        xml.push_str(r#"<sz val=""#);
        xml.push_str(&format_size_100pt(size_100pt));
        xml.push_str(r#""/>"#);
    }

    if let Some(color) = style.color.and_then(|c| c.argb()) {
        xml.push_str(r#"<color rgb=""#);
        xml.push_str(&format!("{:08X}", color));
        xml.push_str(r#""/>"#);
    }

    if let Some(bold) = style.bold {
        if bold {
            xml.push_str("<b/>");
        } else {
            xml.push_str(r#"<b val="0"/>"#);
        }
    }

    if let Some(italic) = style.italic {
        if italic {
            xml.push_str("<i/>");
        } else {
            xml.push_str(r#"<i val="0"/>"#);
        }
    }

    if let Some(underline) = style.underline {
        match underline {
            Underline::Single => xml.push_str("<u/>"),
            other => {
                xml.push_str(r#"<u val=""#);
                xml.push_str(other.to_ooxml().unwrap_or("single"));
                xml.push_str(r#""/>"#);
            }
        }
    }
}

fn format_size_100pt(size_100pt: u16) -> String {
    let int = size_100pt / 100;
    let frac = size_100pt % 100;
    if frac == 0 {
        return int.to_string();
    }

    let mut s = format!("{int}.{frac:02}");
    while s.ends_with('0') {
        s.pop();
    }
    s
}

fn needs_space_preserve(s: &str) -> bool {
    s.starts_with(char::is_whitespace) || s.ends_with(char::is_whitespace)
}

fn escape_text(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn escape_attr(s: &str) -> String {
    escape_text(s)
        .replace('\"', "&quot;")
        .replace('\'', "&apos;")
}

fn write_workbook_xml(
    doc: &XlsxDocument,
    original: Option<&[u8]>,
    sheets: &[SheetMeta],
) -> Result<Vec<u8>, WriteError> {
    if let Some(original) = original {
        return patch_workbook_xml(doc, original, sheets);
    }

    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );
    xml.push_str("<workbookPr");
    if doc.meta.date_system == DateSystem::V1904 {
        xml.push_str(r#" date1904="1""#);
    }
    xml.push_str("/>");

    if !WorkbookProtection::is_default(&doc.workbook.workbook_protection) {
        let protection = &doc.workbook.workbook_protection;
        xml.push_str("<workbookProtection");
        if protection.lock_structure {
            xml.push_str(r#" lockStructure="1""#);
        }
        if protection.lock_windows {
            xml.push_str(r#" lockWindows="1""#);
        }
        if let Some(hash) = protection.password_hash {
            xml.push_str(&format!(r#" workbookPassword="{:04X}""#, hash));
        }
        xml.push_str("/>");
    }

    // Workbook view state (`bookViews/workbookView`) is optional. Only emit it for new workbooks
    // when the view is meaningfully non-default (e.g. active tab is not the first sheet, or window
    // geometry/state is explicitly set). This keeps new document output minimal while still
    // preserving `.xls`-imported window metadata when exporting to `.xlsx`.
    let window = doc.workbook.view.window.as_ref().filter(|window| {
        window.x.is_some()
            || window.y.is_some()
            || window.width.is_some()
            || window.height.is_some()
            || window.state.is_some()
    });
    let active_tab_idx = doc
        .workbook
        .active_sheet_id()
        .and_then(|active| sheets.iter().position(|meta| meta.worksheet_id == active))
        .unwrap_or(0);
    let include_book_views = window.is_some() || active_tab_idx != 0;
    if include_book_views {
        xml.push_str("<bookViews><workbookView");
        if active_tab_idx != 0 {
            xml.push_str(&format!(r#" activeTab="{active_tab_idx}""#));
        }
        if let Some(window) = window {
            if let Some(x) = window.x {
                xml.push_str(&format!(r#" xWindow="{x}""#));
            }
            if let Some(y) = window.y {
                xml.push_str(&format!(r#" yWindow="{y}""#));
            }
            if let Some(width) = window.width {
                xml.push_str(&format!(r#" windowWidth="{width}""#));
            }
            if let Some(height) = window.height {
                xml.push_str(&format!(r#" windowHeight="{height}""#));
            }
            if let Some(state) = window.state {
                match state {
                    WorkbookWindowState::Normal => {}
                    WorkbookWindowState::Minimized => {
                        xml.push_str(r#" windowState="minimized""#);
                    }
                    WorkbookWindowState::Maximized => {
                        xml.push_str(r#" windowState="maximized""#);
                    }
                }
            }
        }
        xml.push_str("/></bookViews>");
    }

    xml.push_str("<sheets>");
    for sheet_meta in sheets {
        let sheet = doc.workbook.sheet(sheet_meta.worksheet_id);
        let name = sheet.map(|s| s.name.as_str()).unwrap_or("Sheet");
        let visibility = sheet
            .map(|s| s.visibility)
            .unwrap_or(SheetVisibility::Visible);
        xml.push_str("<sheet");
        xml.push_str(&format!(r#" name="{}""#, escape_attr(name)));
        xml.push_str(&format!(r#" sheetId="{}""#, sheet_meta.sheet_id));
        xml.push_str(&format!(
            r#" r:id="{}""#,
            escape_attr(&sheet_meta.relationship_id)
        ));
        match visibility {
            SheetVisibility::Visible => {}
            SheetVisibility::Hidden => xml.push_str(r#" state="hidden""#),
            SheetVisibility::VeryHidden => xml.push_str(r#" state="veryHidden""#),
        }
        xml.push_str("/>");
    }
    xml.push_str("</sheets>");
    xml.push_str("</workbook>");
    Ok(xml.into_bytes())
}

fn patch_workbook_xml(
    doc: &XlsxDocument,
    original: &[u8],
    sheets: &[SheetMeta],
) -> Result<Vec<u8>, WriteError> {
    let mut rel_id_to_index: HashMap<&str, usize> = HashMap::with_capacity(sheets.len());
    for (idx, sheet) in sheets.iter().enumerate() {
        rel_id_to_index.insert(sheet.relationship_id.as_str(), idx);
    }
    let old_sheet_index_to_new_index: Vec<Option<usize>> = doc
        .meta
        .sheets
        .iter()
        .map(|meta| rel_id_to_index.get(meta.relationship_id.as_str()).copied())
        .collect();
    let new_sheet_len = sheets.len();

    let mut reader = Reader::from_reader(original);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut writer = Writer::new(Vec::with_capacity(original.len()));

    let mut spreadsheetml_prefix: Option<String> = None;
    let mut office_rels_prefix: Option<String> = None;

    let want_workbook_protection =
        !WorkbookProtection::is_default(&doc.workbook.workbook_protection);
    let mut saw_workbook_protection = false;
    let mut inserted_workbook_protection = false;
    let mut skipping_workbook_protection = false;

    let mut skipping_sheets = false;
    let mut skipping_workbook_pr = false;
    let mut skipping_calc_pr = false;
    let mut skipping_defined_name: usize = 0;

    let view_window = doc.workbook.view.window.as_ref().filter(|window| {
        window.x.is_some()
            || window.y.is_some()
            || window.width.is_some()
            || window.height.is_some()
            || window.state.is_some()
    });
    let active_tab_idx = doc
        .workbook
        .active_sheet_id()
        .and_then(|active| sheets.iter().position(|meta| meta.worksheet_id == active))
        .unwrap_or(0);
    let want_book_views = view_window.is_some() || active_tab_idx != 0;
    let mut saw_book_views = false;
    let mut inserted_book_views = false;
    let mut in_book_views = false;
    let mut saw_workbook_view = false;
    loop {
        let event = reader.read_event_into(&mut buf)?;

        if skipping_defined_name > 0 {
            match event {
                Event::Start(_) => {
                    skipping_defined_name += 1;
                }
                Event::End(_) => {
                    skipping_defined_name = skipping_defined_name.saturating_sub(1);
                }
                _ => {}
            }
            buf.clear();
            continue;
        }

        match event {
            Event::Start(e) if e.local_name().as_ref() == b"workbook" => {
                let ns = crate::xml::workbook_xml_namespaces_from_workbook_start(&e)?;
                spreadsheetml_prefix = ns.spreadsheetml_prefix;
                office_rels_prefix = ns.office_relationships_prefix;
                writer.write_event(Event::Start(e.into_owned()))?;
            }
            Event::Empty(e) if e.local_name().as_ref() == b"workbook" => {
                let ns = crate::xml::workbook_xml_namespaces_from_workbook_start(&e)?;
                spreadsheetml_prefix = ns.spreadsheetml_prefix;
                office_rels_prefix = ns.office_relationships_prefix;
                // Expand self-closing `<workbook/>` roots so we can synthesize required children
                // like `<workbookPr/>` and `<sheets>...</sheets>`.
                //
                // Some producers emit degenerate workbooks that self-close the root element.
                // If we preserved that byte-for-byte we would have nowhere to insert required
                // SpreadsheetML children and would output an invalid workbook.
                let workbook_tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();

                // Preserve the qualified name + all existing attributes/namespace declarations.
                writer.get_mut().extend_from_slice(b"<");
                writer.get_mut().extend_from_slice(workbook_tag.as_bytes());
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    writer.get_mut().push(b' ');
                    writer.get_mut().extend_from_slice(attr.key.as_ref());
                    writer.get_mut().extend_from_slice(b"=\"");
                    writer.get_mut().extend_from_slice(
                        escape_attr(&attr.unescape_value()?.into_owned()).as_bytes(),
                    );
                    writer.get_mut().push(b'"');
                }
                writer.get_mut().push(b'>');

                // Always emit a `<workbookPr/>` so Excel considers the workbook structure valid.
                let workbook_pr_tag = prefixed_tag(spreadsheetml_prefix.as_deref(), "workbookPr");
                writer.get_mut().extend_from_slice(b"<");
                writer
                    .get_mut()
                    .extend_from_slice(workbook_pr_tag.as_bytes());
                if doc.meta.date_system == DateSystem::V1904 {
                    writer.get_mut().extend_from_slice(br#" date1904="1""#);
                }
                writer.get_mut().extend_from_slice(b"/>");

                // If the document wants workbookProtection but the original workbook has no
                // children at all, synthesize it (matching the new-workbook writer behavior).
                if want_workbook_protection {
                    let tag = prefixed_tag(spreadsheetml_prefix.as_deref(), "workbookProtection");
                    write_new_workbook_protection(doc, &mut writer, tag.as_str())?;
                    inserted_workbook_protection = true;
                }

                // Optionally emit bookViews/workbookView when the view is meaningfully non-default
                // (mirroring `write_workbook_xml` for new workbooks). This is not strictly required
                // for validity, but keeps `.xls`-imported window metadata round-trippable even when
                // the workbook root was self-closing.
                let active_tab_idx = doc
                    .workbook
                    .active_sheet_id()
                    .and_then(|active| sheets.iter().position(|meta| meta.worksheet_id == active))
                    .unwrap_or(0);
                let include_book_views = view_window.is_some() || active_tab_idx != 0;
                if include_book_views {
                    let book_views_tag = prefixed_tag(spreadsheetml_prefix.as_deref(), "bookViews");
                    let workbook_view_tag =
                        prefixed_tag(spreadsheetml_prefix.as_deref(), "workbookView");
                    writer.get_mut().extend_from_slice(b"<");
                    writer
                        .get_mut()
                        .extend_from_slice(book_views_tag.as_bytes());
                    writer.get_mut().extend_from_slice(b"><");
                    writer
                        .get_mut()
                        .extend_from_slice(workbook_view_tag.as_bytes());
                    if active_tab_idx != 0 {
                        writer.get_mut().extend_from_slice(b" activeTab=\"");
                        writer
                            .get_mut()
                            .extend_from_slice(active_tab_idx.to_string().as_bytes());
                        writer.get_mut().push(b'"');
                    }
                    if let Some(window) = view_window {
                        if let Some(x) = window.x {
                            writer.get_mut().extend_from_slice(b" xWindow=\"");
                            writer.get_mut().extend_from_slice(x.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(y) = window.y {
                            writer.get_mut().extend_from_slice(b" yWindow=\"");
                            writer.get_mut().extend_from_slice(y.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(width) = window.width {
                            writer.get_mut().extend_from_slice(b" windowWidth=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(width.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(height) = window.height {
                            writer.get_mut().extend_from_slice(b" windowHeight=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(height.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(state) = window.state {
                            match state {
                                WorkbookWindowState::Normal => {}
                                WorkbookWindowState::Minimized => {
                                    writer
                                        .get_mut()
                                        .extend_from_slice(br#" windowState="minimized""#);
                                }
                                WorkbookWindowState::Maximized => {
                                    writer
                                        .get_mut()
                                        .extend_from_slice(br#" windowState="maximized""#);
                                }
                            }
                        }
                    }
                    writer.get_mut().extend_from_slice(b"/></");
                    writer
                        .get_mut()
                        .extend_from_slice(book_views_tag.as_bytes());
                    writer.get_mut().extend_from_slice(b">");
                }

                let sheets_tag = prefixed_tag(spreadsheetml_prefix.as_deref(), "sheets");
                let sheet_tag = prefixed_tag(spreadsheetml_prefix.as_deref(), "sheet");
                let needs_r_namespace = office_rels_prefix.is_none();
                if needs_r_namespace {
                    office_rels_prefix = Some("r".to_string());
                }
                let rel_id_attr = prefixed_tag(office_rels_prefix.as_deref(), "id");

                writer.get_mut().extend_from_slice(b"<");
                writer.get_mut().extend_from_slice(sheets_tag.as_bytes());
                if needs_r_namespace {
                    writer.get_mut().extend_from_slice(br#" xmlns:r=""#);
                    writer
                        .get_mut()
                        .extend_from_slice(crate::xml::OFFICE_RELATIONSHIPS_NS.as_bytes());
                    writer.get_mut().push(b'"');
                }
                writer.get_mut().push(b'>');
                for sheet_meta in sheets {
                    let sheet = doc.workbook.sheet(sheet_meta.worksheet_id);
                    let name = sheet.map(|s| s.name.as_str()).unwrap_or("Sheet");
                    let visibility = sheet
                        .map(|s| s.visibility)
                        .unwrap_or(SheetVisibility::Visible);
                    writer.get_mut().extend_from_slice(b"<");
                    writer.get_mut().extend_from_slice(sheet_tag.as_bytes());
                    writer.get_mut().extend_from_slice(b" name=\"");
                    writer
                        .get_mut()
                        .extend_from_slice(escape_attr(name).as_bytes());
                    writer.get_mut().push(b'"');
                    writer.get_mut().extend_from_slice(b" sheetId=\"");
                    writer
                        .get_mut()
                        .extend_from_slice(sheet_meta.sheet_id.to_string().as_bytes());
                    writer.get_mut().push(b'"');
                    writer.get_mut().push(b' ');
                    writer.get_mut().extend_from_slice(rel_id_attr.as_bytes());
                    writer.get_mut().extend_from_slice(b"=\"");
                    writer
                        .get_mut()
                        .extend_from_slice(escape_attr(&sheet_meta.relationship_id).as_bytes());
                    writer.get_mut().push(b'"');
                    match visibility {
                        SheetVisibility::Visible => {}
                        SheetVisibility::Hidden => {
                            writer.get_mut().extend_from_slice(b" state=\"hidden\"");
                        }
                        SheetVisibility::VeryHidden => {
                            writer.get_mut().extend_from_slice(b" state=\"veryHidden\"");
                        }
                    }
                    writer.get_mut().extend_from_slice(b"/>");
                }
                writer.get_mut().extend_from_slice(b"</");
                writer.get_mut().extend_from_slice(sheets_tag.as_bytes());
                writer.get_mut().extend_from_slice(b">");

                writer.get_mut().extend_from_slice(b"</");
                writer.get_mut().extend_from_slice(workbook_tag.as_bytes());
                writer.get_mut().extend_from_slice(b">");
            }

            Event::Start(e) if e.local_name().as_ref() == b"workbookPr" => {
                skipping_workbook_pr = true;
                let empty = Event::Empty(e.into_owned());
                match empty {
                    Event::Empty(e) => write_workbook_pr(doc, &mut writer, &e)?,
                    _ => unreachable!(),
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == b"workbookPr" => {
                write_workbook_pr(doc, &mut writer, &e)?
            }
            Event::End(e) if e.local_name().as_ref() == b"workbookPr" => {
                if skipping_workbook_pr {
                    skipping_workbook_pr = false;
                } else {
                    writer.write_event(Event::End(e.into_owned()))?;
                }
            }

            Event::Start(e) if e.local_name().as_ref() == b"workbookProtection" => {
                saw_workbook_protection = true;
                skipping_workbook_protection = true;
                if want_workbook_protection && !inserted_workbook_protection {
                    let empty = Event::Empty(e.into_owned());
                    match empty {
                        Event::Empty(e) => write_workbook_protection(doc, &mut writer, &e)?,
                        _ => unreachable!(),
                    }
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == b"workbookProtection" => {
                saw_workbook_protection = true;
                if want_workbook_protection && !inserted_workbook_protection {
                    write_workbook_protection(doc, &mut writer, &e)?;
                }
            }
            Event::End(e) if e.local_name().as_ref() == b"workbookProtection" => {
                if skipping_workbook_protection {
                    skipping_workbook_protection = false;
                } else {
                    writer.write_event(Event::End(e.into_owned()))?;
                }
            }

            Event::Start(e) if e.local_name().as_ref() == b"calcPr" => {
                skipping_calc_pr = true;
                let empty = Event::Empty(e.into_owned());
                match empty {
                    Event::Empty(e) => write_calc_pr(doc, &mut writer, &e)?,
                    _ => unreachable!(),
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == b"calcPr" => {
                write_calc_pr(doc, &mut writer, &e)?
            }
            Event::End(e) if e.local_name().as_ref() == b"calcPr" => {
                if skipping_calc_pr {
                    skipping_calc_pr = false;
                } else {
                    writer.write_event(Event::End(e.into_owned()))?;
                }
            }

            Event::Start(e) if e.local_name().as_ref() == b"bookViews" => {
                saw_book_views = true;
                in_book_views = true;
                saw_workbook_view = false;
                if want_workbook_protection
                    && !saw_workbook_protection
                    && !inserted_workbook_protection
                {
                    let tag = prefixed_tag(spreadsheetml_prefix.as_deref(), "workbookProtection");
                    write_new_workbook_protection(doc, &mut writer, tag.as_str())?;
                    inserted_workbook_protection = true;
                }
                writer.write_event(Event::Start(e.into_owned()))?;
            }
            Event::Empty(e) if e.local_name().as_ref() == b"bookViews" => {
                saw_book_views = true;
                if want_workbook_protection
                    && !saw_workbook_protection
                    && !inserted_workbook_protection
                {
                    let tag = prefixed_tag(spreadsheetml_prefix.as_deref(), "workbookProtection");
                    write_new_workbook_protection(doc, &mut writer, tag.as_str())?;
                    inserted_workbook_protection = true;
                }
                if want_book_views {
                    // Replace `<bookViews/>` with a full section containing a `workbookView`.
                    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    let workbook_view_tag =
                        prefixed_tag(spreadsheetml_prefix.as_deref(), "workbookView");

                    writer.get_mut().extend_from_slice(b"<");
                    writer.get_mut().extend_from_slice(tag.as_bytes());
                    for attr in e.attributes() {
                        let attr = attr?;
                        writer.get_mut().push(b' ');
                        writer.get_mut().extend_from_slice(attr.key.as_ref());
                        writer.get_mut().extend_from_slice(b"=\"");
                        writer.get_mut().extend_from_slice(
                            escape_attr(&attr.unescape_value()?.into_owned()).as_bytes(),
                        );
                        writer.get_mut().push(b'"');
                    }
                    writer.get_mut().push(b'>');

                    writer.get_mut().extend_from_slice(b"<");
                    writer
                        .get_mut()
                        .extend_from_slice(workbook_view_tag.as_bytes());
                    if active_tab_idx != 0 {
                        writer.get_mut().extend_from_slice(b" activeTab=\"");
                        writer
                            .get_mut()
                            .extend_from_slice(active_tab_idx.to_string().as_bytes());
                        writer.get_mut().push(b'"');
                    }
                    if let Some(window) = view_window {
                        if let Some(x) = window.x {
                            writer.get_mut().extend_from_slice(b" xWindow=\"");
                            writer.get_mut().extend_from_slice(x.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(y) = window.y {
                            writer.get_mut().extend_from_slice(b" yWindow=\"");
                            writer.get_mut().extend_from_slice(y.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(width) = window.width {
                            writer.get_mut().extend_from_slice(b" windowWidth=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(width.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(height) = window.height {
                            writer.get_mut().extend_from_slice(b" windowHeight=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(height.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(state) = window.state {
                            match state {
                                WorkbookWindowState::Normal => {}
                                WorkbookWindowState::Minimized => writer
                                    .get_mut()
                                    .extend_from_slice(b" windowState=\"minimized\""),
                                WorkbookWindowState::Maximized => writer
                                    .get_mut()
                                    .extend_from_slice(b" windowState=\"maximized\""),
                            }
                        }
                    }
                    writer.get_mut().extend_from_slice(b"/>");

                    writer.get_mut().extend_from_slice(b"</");
                    writer.get_mut().extend_from_slice(tag.as_bytes());
                    writer.get_mut().push(b'>');
                } else {
                    writer.write_event(Event::Empty(e.into_owned()))?;
                }
            }
            Event::End(e) if e.local_name().as_ref() == b"bookViews" => {
                if want_book_views && in_book_views && !saw_workbook_view {
                    // Insert a workbookView if the workbook declared bookViews but did not include
                    // one. This keeps non-default model view state round-trip safe.
                    let workbook_view_tag =
                        prefixed_tag(spreadsheetml_prefix.as_deref(), "workbookView");
                    writer.get_mut().extend_from_slice(b"<");
                    writer
                        .get_mut()
                        .extend_from_slice(workbook_view_tag.as_bytes());
                    if active_tab_idx != 0 {
                        writer.get_mut().extend_from_slice(b" activeTab=\"");
                        writer
                            .get_mut()
                            .extend_from_slice(active_tab_idx.to_string().as_bytes());
                        writer.get_mut().push(b'"');
                    }
                    if let Some(window) = view_window {
                        if let Some(x) = window.x {
                            writer.get_mut().extend_from_slice(b" xWindow=\"");
                            writer.get_mut().extend_from_slice(x.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(y) = window.y {
                            writer.get_mut().extend_from_slice(b" yWindow=\"");
                            writer.get_mut().extend_from_slice(y.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(width) = window.width {
                            writer.get_mut().extend_from_slice(b" windowWidth=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(width.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(height) = window.height {
                            writer.get_mut().extend_from_slice(b" windowHeight=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(height.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(state) = window.state {
                            match state {
                                WorkbookWindowState::Normal => {}
                                WorkbookWindowState::Minimized => writer
                                    .get_mut()
                                    .extend_from_slice(b" windowState=\"minimized\""),
                                WorkbookWindowState::Maximized => writer
                                    .get_mut()
                                    .extend_from_slice(b" windowState=\"maximized\""),
                            }
                        }
                    }
                    writer.get_mut().extend_from_slice(b"/>");
                }
                in_book_views = false;
                writer.write_event(Event::End(e.into_owned()))?;
            }

            Event::Start(e) if e.local_name().as_ref() == b"workbookView" => {
                saw_workbook_view = true;
                let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                writer.get_mut().extend_from_slice(b"<");
                writer.get_mut().extend_from_slice(tag.as_bytes());
                let mut saw_x_window = false;
                let mut saw_y_window = false;
                let mut saw_window_width = false;
                let mut saw_window_height = false;
                let mut saw_window_state = false;
                let mut saw_active_tab = false;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"activeTab" => saw_active_tab = true,
                        b"xWindow" => saw_x_window = true,
                        b"yWindow" => saw_y_window = true,
                        b"windowWidth" => saw_window_width = true,
                        b"windowHeight" => saw_window_height = true,
                        b"windowState" => saw_window_state = true,
                        _ => {}
                    }
                    writer.get_mut().push(b' ');
                    writer.get_mut().extend_from_slice(attr.key.as_ref());
                    writer.get_mut().extend_from_slice(b"=\"");
                    let value = match attr.key.as_ref() {
                        b"activeTab" => active_tab_idx.to_string(),
                        b"firstSheet" => {
                            let old = attr.unescape_value()?.trim().parse::<usize>().ok();
                            old.and_then(|idx| {
                                old_sheet_index_to_new_index
                                    .get(idx)
                                    .copied()
                                    .flatten()
                                    .or_else(|| (new_sheet_len > 0).then_some(0))
                            })
                            .map(|idx| idx.to_string())
                            .unwrap_or_else(|| {
                                attr.unescape_value()
                                    .map(|v| v.into_owned())
                                    .unwrap_or_default()
                            })
                        }
                        b"xWindow" => view_window
                            .and_then(|window| window.x)
                            .map(|x| x.to_string())
                            .unwrap_or_else(|| {
                                attr.unescape_value()
                                    .map(|v| v.into_owned())
                                    .unwrap_or_default()
                            }),
                        b"yWindow" => view_window
                            .and_then(|window| window.y)
                            .map(|y| y.to_string())
                            .unwrap_or_else(|| {
                                attr.unescape_value()
                                    .map(|v| v.into_owned())
                                    .unwrap_or_default()
                            }),
                        b"windowWidth" => view_window
                            .and_then(|window| window.width)
                            .map(|w| w.to_string())
                            .unwrap_or_else(|| {
                                attr.unescape_value()
                                    .map(|v| v.into_owned())
                                    .unwrap_or_default()
                            }),
                        b"windowHeight" => view_window
                            .and_then(|window| window.height)
                            .map(|h| h.to_string())
                            .unwrap_or_else(|| {
                                attr.unescape_value()
                                    .map(|v| v.into_owned())
                                    .unwrap_or_default()
                            }),
                        b"windowState" => view_window
                            .and_then(|window| window.state)
                            .and_then(|state| match state {
                                WorkbookWindowState::Normal => None,
                                WorkbookWindowState::Minimized => Some("minimized".to_string()),
                                WorkbookWindowState::Maximized => Some("maximized".to_string()),
                            })
                            .unwrap_or_else(|| {
                                attr.unescape_value()
                                    .map(|v| v.into_owned())
                                    .unwrap_or_default()
                            }),
                        _ => attr.unescape_value()?.into_owned(),
                    };
                    writer
                        .get_mut()
                        .extend_from_slice(escape_attr(&value).as_bytes());
                    writer.get_mut().push(b'"');
                }

                if let Some(window) = view_window {
                    if !saw_x_window {
                        if let Some(x) = window.x {
                            writer.get_mut().extend_from_slice(b" xWindow=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(escape_attr(&x.to_string()).as_bytes());
                            writer.get_mut().push(b'"');
                        }
                    }
                    if !saw_y_window {
                        if let Some(y) = window.y {
                            writer.get_mut().extend_from_slice(b" yWindow=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(escape_attr(&y.to_string()).as_bytes());
                            writer.get_mut().push(b'"');
                        }
                    }
                    if !saw_window_width {
                        if let Some(width) = window.width {
                            writer.get_mut().extend_from_slice(b" windowWidth=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(escape_attr(&width.to_string()).as_bytes());
                            writer.get_mut().push(b'"');
                        }
                    }
                    if !saw_window_height {
                        if let Some(height) = window.height {
                            writer.get_mut().extend_from_slice(b" windowHeight=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(escape_attr(&height.to_string()).as_bytes());
                            writer.get_mut().push(b'"');
                        }
                    }
                    if !saw_window_state {
                        if let Some(state) = window.state {
                            let state_str = match state {
                                WorkbookWindowState::Normal => None,
                                WorkbookWindowState::Minimized => Some("minimized"),
                                WorkbookWindowState::Maximized => Some("maximized"),
                            };
                            if let Some(state_str) = state_str {
                                writer.get_mut().extend_from_slice(b" windowState=\"");
                                writer.get_mut().extend_from_slice(state_str.as_bytes());
                                writer.get_mut().push(b'"');
                            }
                        }
                    }
                }
                if !saw_active_tab && active_tab_idx != 0 {
                    writer.get_mut().extend_from_slice(b" activeTab=\"");
                    writer
                        .get_mut()
                        .extend_from_slice(active_tab_idx.to_string().as_bytes());
                    writer.get_mut().push(b'"');
                }
                writer.get_mut().push(b'>');
            }
            Event::Empty(e) if e.local_name().as_ref() == b"workbookView" => {
                saw_workbook_view = true;
                let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                writer.get_mut().extend_from_slice(b"<");
                writer.get_mut().extend_from_slice(tag.as_bytes());
                let mut saw_x_window = false;
                let mut saw_y_window = false;
                let mut saw_window_width = false;
                let mut saw_window_height = false;
                let mut saw_window_state = false;
                let mut saw_active_tab = false;
                for attr in e.attributes() {
                    let attr = attr?;
                    match attr.key.as_ref() {
                        b"activeTab" => saw_active_tab = true,
                        b"xWindow" => saw_x_window = true,
                        b"yWindow" => saw_y_window = true,
                        b"windowWidth" => saw_window_width = true,
                        b"windowHeight" => saw_window_height = true,
                        b"windowState" => saw_window_state = true,
                        _ => {}
                    }
                    writer.get_mut().push(b' ');
                    writer.get_mut().extend_from_slice(attr.key.as_ref());
                    writer.get_mut().extend_from_slice(b"=\"");
                    let value = match attr.key.as_ref() {
                        b"activeTab" => active_tab_idx.to_string(),
                        b"firstSheet" => {
                            let old = attr.unescape_value()?.trim().parse::<usize>().ok();
                            old.and_then(|idx| {
                                old_sheet_index_to_new_index
                                    .get(idx)
                                    .copied()
                                    .flatten()
                                    .or_else(|| (new_sheet_len > 0).then_some(0))
                            })
                            .map(|idx| idx.to_string())
                            .unwrap_or_else(|| {
                                attr.unescape_value()
                                    .map(|v| v.into_owned())
                                    .unwrap_or_default()
                            })
                        }
                        b"xWindow" => view_window
                            .and_then(|window| window.x)
                            .map(|x| x.to_string())
                            .unwrap_or_else(|| {
                                attr.unescape_value()
                                    .map(|v| v.into_owned())
                                    .unwrap_or_default()
                            }),
                        b"yWindow" => view_window
                            .and_then(|window| window.y)
                            .map(|y| y.to_string())
                            .unwrap_or_else(|| {
                                attr.unescape_value()
                                    .map(|v| v.into_owned())
                                    .unwrap_or_default()
                            }),
                        b"windowWidth" => view_window
                            .and_then(|window| window.width)
                            .map(|w| w.to_string())
                            .unwrap_or_else(|| {
                                attr.unescape_value()
                                    .map(|v| v.into_owned())
                                    .unwrap_or_default()
                            }),
                        b"windowHeight" => view_window
                            .and_then(|window| window.height)
                            .map(|h| h.to_string())
                            .unwrap_or_else(|| {
                                attr.unescape_value()
                                    .map(|v| v.into_owned())
                                    .unwrap_or_default()
                            }),
                        b"windowState" => view_window
                            .and_then(|window| window.state)
                            .and_then(|state| match state {
                                WorkbookWindowState::Normal => None,
                                WorkbookWindowState::Minimized => Some("minimized".to_string()),
                                WorkbookWindowState::Maximized => Some("maximized".to_string()),
                            })
                            .unwrap_or_else(|| {
                                attr.unescape_value()
                                    .map(|v| v.into_owned())
                                    .unwrap_or_default()
                            }),
                        _ => attr.unescape_value()?.into_owned(),
                    };
                    writer
                        .get_mut()
                        .extend_from_slice(escape_attr(&value).as_bytes());
                    writer.get_mut().push(b'"');
                }

                if let Some(window) = view_window {
                    if !saw_x_window {
                        if let Some(x) = window.x {
                            writer.get_mut().extend_from_slice(b" xWindow=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(escape_attr(&x.to_string()).as_bytes());
                            writer.get_mut().push(b'"');
                        }
                    }
                    if !saw_y_window {
                        if let Some(y) = window.y {
                            writer.get_mut().extend_from_slice(b" yWindow=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(escape_attr(&y.to_string()).as_bytes());
                            writer.get_mut().push(b'"');
                        }
                    }
                    if !saw_window_width {
                        if let Some(width) = window.width {
                            writer.get_mut().extend_from_slice(b" windowWidth=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(escape_attr(&width.to_string()).as_bytes());
                            writer.get_mut().push(b'"');
                        }
                    }
                    if !saw_window_height {
                        if let Some(height) = window.height {
                            writer.get_mut().extend_from_slice(b" windowHeight=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(escape_attr(&height.to_string()).as_bytes());
                            writer.get_mut().push(b'"');
                        }
                    }
                    if !saw_window_state {
                        if let Some(state) = window.state {
                            let state_str = match state {
                                WorkbookWindowState::Normal => None,
                                WorkbookWindowState::Minimized => Some("minimized"),
                                WorkbookWindowState::Maximized => Some("maximized"),
                            };
                            if let Some(state_str) = state_str {
                                writer.get_mut().extend_from_slice(b" windowState=\"");
                                writer.get_mut().extend_from_slice(state_str.as_bytes());
                                writer.get_mut().push(b'"');
                            }
                        }
                    }
                }
                if !saw_active_tab && active_tab_idx != 0 {
                    writer.get_mut().extend_from_slice(b" activeTab=\"");
                    writer
                        .get_mut()
                        .extend_from_slice(active_tab_idx.to_string().as_bytes());
                    writer.get_mut().push(b'"');
                }
                writer.get_mut().extend_from_slice(b"/>");
            }

            Event::Start(e) if e.local_name().as_ref() == b"sheets" => {
                if want_workbook_protection
                    && !saw_workbook_protection
                    && !inserted_workbook_protection
                {
                    let tag = prefixed_tag(spreadsheetml_prefix.as_deref(), "workbookProtection");
                    write_new_workbook_protection(doc, &mut writer, tag.as_str())?;
                    inserted_workbook_protection = true;
                }
                if want_book_views && !saw_book_views && !inserted_book_views {
                    let book_views_tag = prefixed_tag(spreadsheetml_prefix.as_deref(), "bookViews");
                    let workbook_view_tag =
                        prefixed_tag(spreadsheetml_prefix.as_deref(), "workbookView");
                    writer.get_mut().extend_from_slice(b"<");
                    writer
                        .get_mut()
                        .extend_from_slice(book_views_tag.as_bytes());
                    writer.get_mut().push(b'>');
                    writer.get_mut().extend_from_slice(b"<");
                    writer
                        .get_mut()
                        .extend_from_slice(workbook_view_tag.as_bytes());
                    if active_tab_idx != 0 {
                        writer.get_mut().extend_from_slice(b" activeTab=\"");
                        writer
                            .get_mut()
                            .extend_from_slice(active_tab_idx.to_string().as_bytes());
                        writer.get_mut().push(b'"');
                    }
                    if let Some(window) = view_window {
                        if let Some(x) = window.x {
                            writer.get_mut().extend_from_slice(b" xWindow=\"");
                            writer.get_mut().extend_from_slice(x.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(y) = window.y {
                            writer.get_mut().extend_from_slice(b" yWindow=\"");
                            writer.get_mut().extend_from_slice(y.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(width) = window.width {
                            writer.get_mut().extend_from_slice(b" windowWidth=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(width.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(height) = window.height {
                            writer.get_mut().extend_from_slice(b" windowHeight=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(height.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(state) = window.state {
                            match state {
                                WorkbookWindowState::Normal => {}
                                WorkbookWindowState::Minimized => writer
                                    .get_mut()
                                    .extend_from_slice(b" windowState=\"minimized\""),
                                WorkbookWindowState::Maximized => writer
                                    .get_mut()
                                    .extend_from_slice(b" windowState=\"maximized\""),
                            }
                        }
                    }
                    writer.get_mut().extend_from_slice(b"/>");
                    writer.get_mut().extend_from_slice(b"</");
                    writer
                        .get_mut()
                        .extend_from_slice(book_views_tag.as_bytes());
                    writer.get_mut().push(b'>');
                    inserted_book_views = true;
                }
                skipping_sheets = true;
                let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                let sheet_tag = prefixed_tag(spreadsheetml_prefix.as_deref(), "sheet");
                if office_rels_prefix.is_none() {
                    office_rels_prefix = office_relationships_prefix_from_xmlns(&e)?;
                }
                let rel_id_attr = prefixed_tag(office_rels_prefix.as_deref(), "id");

                writer.get_mut().extend_from_slice(b"<");
                writer.get_mut().extend_from_slice(tag.as_bytes());
                for attr in e.attributes() {
                    let attr = attr?;
                    writer.get_mut().push(b' ');
                    writer.get_mut().extend_from_slice(attr.key.as_ref());
                    writer.get_mut().extend_from_slice(b"=\"");
                    writer.get_mut().extend_from_slice(
                        escape_attr(&attr.unescape_value()?.into_owned()).as_bytes(),
                    );
                    writer.get_mut().push(b'"');
                }
                writer.get_mut().push(b'>');

                for sheet_meta in sheets {
                    let sheet = doc.workbook.sheet(sheet_meta.worksheet_id);
                    let name = sheet.map(|s| s.name.as_str()).unwrap_or("Sheet");
                    let visibility = sheet
                        .map(|s| s.visibility)
                        .unwrap_or(SheetVisibility::Visible);
                    writer.get_mut().extend_from_slice(b"<");
                    writer.get_mut().extend_from_slice(sheet_tag.as_bytes());
                    writer.get_mut().extend_from_slice(b" name=\"");
                    writer
                        .get_mut()
                        .extend_from_slice(escape_attr(name).as_bytes());
                    writer.get_mut().push(b'"');
                    writer.get_mut().extend_from_slice(b" sheetId=\"");
                    writer
                        .get_mut()
                        .extend_from_slice(sheet_meta.sheet_id.to_string().as_bytes());
                    writer.get_mut().push(b'"');
                    writer.get_mut().push(b' ');
                    writer.get_mut().extend_from_slice(rel_id_attr.as_bytes());
                    writer.get_mut().extend_from_slice(b"=\"");
                    writer
                        .get_mut()
                        .extend_from_slice(escape_attr(&sheet_meta.relationship_id).as_bytes());
                    writer.get_mut().push(b'"');
                    match visibility {
                        SheetVisibility::Visible => {}
                        SheetVisibility::Hidden => {
                            writer.get_mut().extend_from_slice(b" state=\"hidden\"");
                        }
                        SheetVisibility::VeryHidden => {
                            writer.get_mut().extend_from_slice(b" state=\"veryHidden\"");
                        }
                    }
                    writer.get_mut().extend_from_slice(b"/>");
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == b"sheets" => {
                if want_workbook_protection
                    && !saw_workbook_protection
                    && !inserted_workbook_protection
                {
                    let tag = prefixed_tag(spreadsheetml_prefix.as_deref(), "workbookProtection");
                    write_new_workbook_protection(doc, &mut writer, tag.as_str())?;
                    inserted_workbook_protection = true;
                }
                if want_book_views && !saw_book_views && !inserted_book_views {
                    let book_views_tag = prefixed_tag(spreadsheetml_prefix.as_deref(), "bookViews");
                    let workbook_view_tag =
                        prefixed_tag(spreadsheetml_prefix.as_deref(), "workbookView");
                    writer.get_mut().extend_from_slice(b"<");
                    writer
                        .get_mut()
                        .extend_from_slice(book_views_tag.as_bytes());
                    writer.get_mut().push(b'>');
                    writer.get_mut().extend_from_slice(b"<");
                    writer
                        .get_mut()
                        .extend_from_slice(workbook_view_tag.as_bytes());
                    if active_tab_idx != 0 {
                        writer.get_mut().extend_from_slice(b" activeTab=\"");
                        writer
                            .get_mut()
                            .extend_from_slice(active_tab_idx.to_string().as_bytes());
                        writer.get_mut().push(b'"');
                    }
                    if let Some(window) = view_window {
                        if let Some(x) = window.x {
                            writer.get_mut().extend_from_slice(b" xWindow=\"");
                            writer.get_mut().extend_from_slice(x.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(y) = window.y {
                            writer.get_mut().extend_from_slice(b" yWindow=\"");
                            writer.get_mut().extend_from_slice(y.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(width) = window.width {
                            writer.get_mut().extend_from_slice(b" windowWidth=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(width.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(height) = window.height {
                            writer.get_mut().extend_from_slice(b" windowHeight=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(height.to_string().as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        if let Some(state) = window.state {
                            match state {
                                WorkbookWindowState::Normal => {}
                                WorkbookWindowState::Minimized => writer
                                    .get_mut()
                                    .extend_from_slice(b" windowState=\"minimized\""),
                                WorkbookWindowState::Maximized => writer
                                    .get_mut()
                                    .extend_from_slice(b" windowState=\"maximized\""),
                            }
                        }
                    }
                    writer.get_mut().extend_from_slice(b"/>");
                    writer.get_mut().extend_from_slice(b"</");
                    writer
                        .get_mut()
                        .extend_from_slice(book_views_tag.as_bytes());
                    writer.get_mut().push(b'>');
                    inserted_book_views = true;
                }
                // Replace `<sheets/>` with a full section.
                let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                let sheet_tag = prefixed_tag(spreadsheetml_prefix.as_deref(), "sheet");
                if office_rels_prefix.is_none() {
                    office_rels_prefix = office_relationships_prefix_from_xmlns(&e)?;
                }
                let rel_id_attr = prefixed_tag(office_rels_prefix.as_deref(), "id");

                writer.get_mut().extend_from_slice(b"<");
                writer.get_mut().extend_from_slice(tag.as_bytes());
                for attr in e.attributes() {
                    let attr = attr?;
                    writer.get_mut().push(b' ');
                    writer.get_mut().extend_from_slice(attr.key.as_ref());
                    writer.get_mut().extend_from_slice(b"=\"");
                    writer.get_mut().extend_from_slice(
                        escape_attr(&attr.unescape_value()?.into_owned()).as_bytes(),
                    );
                    writer.get_mut().push(b'"');
                }
                writer.get_mut().push(b'>');
                for sheet_meta in sheets {
                    let sheet = doc.workbook.sheet(sheet_meta.worksheet_id);
                    let name = sheet.map(|s| s.name.as_str()).unwrap_or("Sheet");
                    let visibility = sheet
                        .map(|s| s.visibility)
                        .unwrap_or(SheetVisibility::Visible);
                    writer.get_mut().extend_from_slice(b"<");
                    writer.get_mut().extend_from_slice(sheet_tag.as_bytes());
                    writer.get_mut().extend_from_slice(b" name=\"");
                    writer
                        .get_mut()
                        .extend_from_slice(escape_attr(name).as_bytes());
                    writer.get_mut().push(b'"');
                    writer.get_mut().extend_from_slice(b" sheetId=\"");
                    writer
                        .get_mut()
                        .extend_from_slice(sheet_meta.sheet_id.to_string().as_bytes());
                    writer.get_mut().push(b'"');
                    writer.get_mut().push(b' ');
                    writer.get_mut().extend_from_slice(rel_id_attr.as_bytes());
                    writer.get_mut().extend_from_slice(b"=\"");
                    writer
                        .get_mut()
                        .extend_from_slice(escape_attr(&sheet_meta.relationship_id).as_bytes());
                    writer.get_mut().push(b'"');
                    match visibility {
                        SheetVisibility::Visible => {}
                        SheetVisibility::Hidden => {
                            writer.get_mut().extend_from_slice(b" state=\"hidden\"");
                        }
                        SheetVisibility::VeryHidden => {
                            writer.get_mut().extend_from_slice(b" state=\"veryHidden\"");
                        }
                    }
                    writer.get_mut().extend_from_slice(b"/>");
                }
                writer.get_mut().extend_from_slice(b"</");
                writer.get_mut().extend_from_slice(tag.as_bytes());
                writer.get_mut().extend_from_slice(b">");
            }
            Event::End(e) if e.local_name().as_ref() == b"sheets" => {
                skipping_sheets = false;
                let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                writer.get_mut().extend_from_slice(b"</");
                writer.get_mut().extend_from_slice(tag.as_bytes());
                writer.get_mut().extend_from_slice(b">");
            }

            Event::Start(e) if e.local_name().as_ref() == b"definedName" => {
                let mut local_sheet_id: Option<usize> = None;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    if attr.key.as_ref() == b"localSheetId" {
                        local_sheet_id = attr.unescape_value()?.trim().parse::<usize>().ok();
                        break;
                    }
                }

                if let Some(old_idx) = local_sheet_id {
                    if let Some(new_idx) =
                        old_sheet_index_to_new_index.get(old_idx).copied().flatten()
                    {
                        let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                        writer.get_mut().extend_from_slice(b"<");
                        writer.get_mut().extend_from_slice(tag.as_bytes());
                        for attr in e.attributes().with_checks(false) {
                            let attr = attr?;
                            writer.get_mut().push(b' ');
                            writer.get_mut().extend_from_slice(attr.key.as_ref());
                            writer.get_mut().extend_from_slice(b"=\"");
                            let value = if attr.key.as_ref() == b"localSheetId" {
                                new_idx.to_string()
                            } else {
                                attr.unescape_value()?.into_owned()
                            };
                            writer
                                .get_mut()
                                .extend_from_slice(escape_attr(&value).as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        writer.get_mut().push(b'>');
                    } else {
                        skipping_defined_name = 1;
                    }
                } else {
                    writer.write_event(Event::Start(e.into_owned()))?;
                }
            }
            Event::Empty(e) if e.local_name().as_ref() == b"definedName" => {
                let mut local_sheet_id: Option<usize> = None;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    if attr.key.as_ref() == b"localSheetId" {
                        local_sheet_id = attr.unescape_value()?.trim().parse::<usize>().ok();
                        break;
                    }
                }

                if let Some(old_idx) = local_sheet_id {
                    if let Some(new_idx) =
                        old_sheet_index_to_new_index.get(old_idx).copied().flatten()
                    {
                        let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                        writer.get_mut().extend_from_slice(b"<");
                        writer.get_mut().extend_from_slice(tag.as_bytes());
                        for attr in e.attributes().with_checks(false) {
                            let attr = attr?;
                            writer.get_mut().push(b' ');
                            writer.get_mut().extend_from_slice(attr.key.as_ref());
                            writer.get_mut().extend_from_slice(b"=\"");
                            let value = if attr.key.as_ref() == b"localSheetId" {
                                new_idx.to_string()
                            } else {
                                attr.unescape_value()?.into_owned()
                            };
                            writer
                                .get_mut()
                                .extend_from_slice(escape_attr(&value).as_bytes());
                            writer.get_mut().push(b'"');
                        }
                        writer.get_mut().extend_from_slice(b"/>");
                    }
                } else {
                    writer.write_event(Event::Empty(e.into_owned()))?;
                }
            }

            Event::Eof => break,
            ev if skipping_workbook_pr || skipping_workbook_protection || skipping_calc_pr => {
                drop(ev)
            }
            ev if skipping_sheets => drop(ev),
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn write_workbook_pr(
    doc: &XlsxDocument,
    writer: &mut Writer<Vec<u8>>,
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<(), WriteError> {
    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
    let had_date1904 = e
        .attributes()
        .flatten()
        .any(|a| a.key.as_ref() == b"date1904");

    writer.get_mut().extend_from_slice(b"<");
    writer.get_mut().extend_from_slice(tag.as_bytes());
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"date1904" {
            continue;
        }
        writer.get_mut().push(b' ');
        writer.get_mut().extend_from_slice(attr.key.as_ref());
        writer.get_mut().extend_from_slice(b"=\"");
        writer
            .get_mut()
            .extend_from_slice(escape_attr(&attr.unescape_value()?.into_owned()).as_bytes());
        writer.get_mut().push(b'"');
    }

    if doc.meta.date_system == DateSystem::V1904 {
        writer.get_mut().extend_from_slice(b" date1904=\"1\"");
    } else if had_date1904 {
        writer.get_mut().extend_from_slice(b" date1904=\"0\"");
    }
    writer.get_mut().extend_from_slice(b"/>");
    Ok(())
}

fn write_calc_pr(
    doc: &XlsxDocument,
    writer: &mut Writer<Vec<u8>>,
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<(), WriteError> {
    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
    writer.get_mut().extend_from_slice(b"<");
    writer.get_mut().extend_from_slice(tag.as_bytes());
    for attr in e.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"calcId" | b"calcMode" | b"fullCalcOnLoad" => continue,
            _ => {}
        }
        writer.get_mut().push(b' ');
        writer.get_mut().extend_from_slice(attr.key.as_ref());
        writer.get_mut().extend_from_slice(b"=\"");
        writer
            .get_mut()
            .extend_from_slice(escape_attr(&attr.unescape_value()?.into_owned()).as_bytes());
        writer.get_mut().push(b'"');
    }

    if let Some(calc_id) = &doc.meta.calc_pr.calc_id {
        writer.get_mut().extend_from_slice(b" calcId=\"");
        writer
            .get_mut()
            .extend_from_slice(escape_attr(calc_id).as_bytes());
        writer.get_mut().push(b'"');
    }
    if let Some(calc_mode) = &doc.meta.calc_pr.calc_mode {
        writer.get_mut().extend_from_slice(b" calcMode=\"");
        writer
            .get_mut()
            .extend_from_slice(escape_attr(calc_mode).as_bytes());
        writer.get_mut().push(b'"');
    }
    if let Some(full) = doc.meta.calc_pr.full_calc_on_load {
        writer.get_mut().extend_from_slice(b" fullCalcOnLoad=\"");
        writer
            .get_mut()
            .extend_from_slice(if full { b"1" } else { b"0" });
        writer.get_mut().push(b'"');
    }
    writer.get_mut().extend_from_slice(b"/>");
    Ok(())
}

fn write_workbook_protection(
    doc: &XlsxDocument,
    writer: &mut Writer<Vec<u8>>,
    e: &quick_xml::events::BytesStart<'_>,
) -> Result<(), WriteError> {
    let protection = &doc.workbook.workbook_protection;
    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();

    let mut wrote_lock_structure = false;
    let mut wrote_lock_windows = false;
    let mut wrote_password = false;

    writer.get_mut().extend_from_slice(b"<");
    writer.get_mut().extend_from_slice(tag.as_bytes());
    for attr in e.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"lockStructure" => {
                wrote_lock_structure = true;
                let val = attr.unescape_value()?.into_owned();
                let original = parse_xml_bool(&val);
                let desired = protection.lock_structure;

                writer.get_mut().push(b' ');
                writer.get_mut().extend_from_slice(attr.key.as_ref());
                writer.get_mut().extend_from_slice(b"=\"");
                if original == desired {
                    writer
                        .get_mut()
                        .extend_from_slice(escape_attr(&val).as_bytes());
                } else {
                    writer
                        .get_mut()
                        .extend_from_slice(if desired { b"1" } else { b"0" });
                }
                writer.get_mut().push(b'"');
            }
            b"lockWindows" => {
                wrote_lock_windows = true;
                let val = attr.unescape_value()?.into_owned();
                let original = parse_xml_bool(&val);
                let desired = protection.lock_windows;

                writer.get_mut().push(b' ');
                writer.get_mut().extend_from_slice(attr.key.as_ref());
                writer.get_mut().extend_from_slice(b"=\"");
                if original == desired {
                    writer
                        .get_mut()
                        .extend_from_slice(escape_attr(&val).as_bytes());
                } else {
                    writer
                        .get_mut()
                        .extend_from_slice(if desired { b"1" } else { b"0" });
                }
                writer.get_mut().push(b'"');
            }
            b"workbookPassword" => {
                wrote_password = true;
                let val = attr.unescape_value()?.into_owned();
                let original_hash = parse_xml_u16_hex(&val).filter(|hash| *hash != 0);
                match protection.password_hash {
                    Some(hash) => {
                        writer.get_mut().extend_from_slice(b" workbookPassword=\"");
                        if original_hash == Some(hash) {
                            writer
                                .get_mut()
                                .extend_from_slice(escape_attr(&val).as_bytes());
                        } else {
                            writer
                                .get_mut()
                                .extend_from_slice(format!("{:04X}", hash).as_bytes());
                        }
                        writer.get_mut().push(b'"');
                    }
                    None => {
                        // Preserve the attribute if it was already semantically equivalent to
                        // `None` (e.g. `0000`), but drop it when clearing a real password hash.
                        if original_hash.is_none() {
                            writer.get_mut().push(b' ');
                            writer.get_mut().extend_from_slice(attr.key.as_ref());
                            writer.get_mut().extend_from_slice(b"=\"");
                            writer
                                .get_mut()
                                .extend_from_slice(escape_attr(&val).as_bytes());
                            writer.get_mut().push(b'"');
                        }
                    }
                }
            }
            _ => {
                writer.get_mut().push(b' ');
                writer.get_mut().extend_from_slice(attr.key.as_ref());
                writer.get_mut().extend_from_slice(b"=\"");
                writer.get_mut().extend_from_slice(
                    escape_attr(&attr.unescape_value()?.into_owned()).as_bytes(),
                );
                writer.get_mut().push(b'"');
            }
        }
    }

    if protection.lock_structure && !wrote_lock_structure {
        writer.get_mut().extend_from_slice(br#" lockStructure="1""#);
    }
    if protection.lock_windows && !wrote_lock_windows {
        writer.get_mut().extend_from_slice(br#" lockWindows="1""#);
    }
    if let Some(hash) = protection.password_hash {
        if !wrote_password {
            writer
                .get_mut()
                .extend_from_slice(br#" workbookPassword=""#);
            writer
                .get_mut()
                .extend_from_slice(format!("{:04X}", hash).as_bytes());
            writer.get_mut().push(b'"');
        }
    }

    writer.get_mut().extend_from_slice(b"/>");
    Ok(())
}

fn write_new_workbook_protection(
    doc: &XlsxDocument,
    writer: &mut Writer<Vec<u8>>,
    tag: &str,
) -> Result<(), WriteError> {
    let protection = &doc.workbook.workbook_protection;
    writer.get_mut().extend_from_slice(b"<");
    writer.get_mut().extend_from_slice(tag.as_bytes());
    if protection.lock_structure {
        writer.get_mut().extend_from_slice(br#" lockStructure="1""#);
    }
    if protection.lock_windows {
        writer.get_mut().extend_from_slice(br#" lockWindows="1""#);
    }
    if let Some(hash) = protection.password_hash {
        writer
            .get_mut()
            .extend_from_slice(br#" workbookPassword=""#);
        writer
            .get_mut()
            .extend_from_slice(format!("{:04X}", hash).as_bytes());
        writer.get_mut().push(b'"');
    }
    writer.get_mut().extend_from_slice(b"/>");
    Ok(())
}

fn write_worksheet_xml(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    sheet: &Worksheet,
    original: Option<&[u8]>,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    local_to_global_dxf: Option<&[u32]>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
) -> Result<Vec<u8>, WriteError> {
    if let Some(original) = original {
        return patch_worksheet_xml(
            doc,
            sheet_meta,
            sheet,
            original,
            shared_lookup,
            style_to_xf,
            cell_meta_sheet_ids,
            changed_formula_cells,
        );
    }

    let dimension = dimension::worksheet_dimension_range(sheet).to_string();
    let sheet_format_pr_xml = render_sheet_format_pr(SheetFormatSettings::from_sheet(sheet), None);
    let cols_xml = render_cols(sheet, None, style_to_xf);
    let sheet_protection_xml = render_sheet_protection(&sheet.sheet_protection, None);
    let conditional_formatting_xml = render_conditional_formatting(sheet, local_to_global_dxf);
    let sheet_data_xml = render_sheet_data(
        doc,
        sheet_meta,
        sheet,
        shared_lookup,
        style_to_xf,
        cell_meta_sheet_ids,
        Some(&sheet.outline),
        changed_formula_cells,
    );

    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );
    if sheet.outline != Outline::default() {
        xml.push_str("<sheetPr>");
        xml.push_str("<outlinePr");
        xml.push_str(if sheet.outline.pr.summary_below {
            r#" summaryBelow="1""#
        } else {
            r#" summaryBelow="0""#
        });
        xml.push_str(if sheet.outline.pr.summary_right {
            r#" summaryRight="1""#
        } else {
            r#" summaryRight="0""#
        });
        xml.push_str(if sheet.outline.pr.show_outline_symbols {
            r#" showOutlineSymbols="1""#
        } else {
            r#" showOutlineSymbols="0""#
        });
        xml.push_str("/></sheetPr>");
    }
    xml.push_str(&format!(r#"<dimension ref="{dimension}"/>"#));
    if !sheet_format_pr_xml.is_empty() {
        xml.push_str(&sheet_format_pr_xml);
    }
    if !cols_xml.is_empty() {
        xml.push_str(&cols_xml);
    }
    xml.push_str(&sheet_data_xml);
    if !sheet_protection_xml.is_empty() {
        xml.push_str(&sheet_protection_xml);
    }
    if !conditional_formatting_xml.is_empty() {
        xml.push_str(&conditional_formatting_xml);
    }
    xml.push_str("</worksheet>");
    Ok(xml.into_bytes())
}

fn render_conditional_formatting(sheet: &Worksheet, local_to_global_dxf: Option<&[u32]>) -> String {
    if sheet.conditional_formatting_rules.is_empty() {
        return String::new();
    }

    let priorities =
        crate::conditional_formatting::normalize_cf_priorities(&sheet.conditional_formatting_rules);

    let mut out = String::new();
    for (idx, rule) in sheet.conditional_formatting_rules.iter().enumerate() {
        let priority = priorities.get(idx).copied().unwrap_or(1);
        let Some(cf_rule_xml) = render_cf_rule(rule, priority, local_to_global_dxf) else {
            continue;
        };

        let sqref = rule
            .applies_to
            .iter()
            .map(|r| r.to_string())
            .collect::<Vec<_>>()
            .join(" ");
        if sqref.is_empty() {
            continue;
        }

        out.push_str(r#"<conditionalFormatting sqref=""#);
        out.push_str(&escape_attr(&sqref));
        out.push_str(r#"">"#);
        out.push_str(&cf_rule_xml);
        out.push_str("</conditionalFormatting>");
    }

    out
}

fn render_cf_rule(rule: &CfRule, priority: u32, local_to_global_dxf: Option<&[u32]>) -> Option<String> {
    let mut attrs = String::new();

    if let Some(id) = rule.id.as_deref() {
        attrs.push_str(r#" id=""#);
        attrs.push_str(&escape_attr(id));
        attrs.push('"');
    }

    attrs.push_str(&format!(r#" priority="{}""#, priority));

    if rule.stop_if_true {
        attrs.push_str(r#" stopIfTrue="1""#);
    }

    // Remap per-sheet `dxf_id` to the workbook-global `dxfs` index table. Best-effort:
    // out-of-bounds indices are emitted as no `dxfId` attribute.
    let global_dxf_id = rule
        .dxf_id
        .and_then(|local| local_to_global_dxf?.get(local as usize).copied());
    if let Some(global) = global_dxf_id {
        attrs.push_str(&format!(r#" dxfId="{}""#, global));
    }

    let (type_attr, body, extra_attrs) = match &rule.kind {
        CfRuleKind::Expression { formula } => (
            "expression",
            {
                let file_formula =
                    crate::formula_text::add_xlfn_prefixes(strip_leading_equals(formula));
                format!(r#"<formula>{}</formula>"#, escape_text(&file_formula))
            },
            String::new(),
        ),
        CfRuleKind::CellIs { operator, formulas } => {
            let op = cell_is_operator_attr(*operator);
            let mut inner = String::new();
            for f in formulas {
                let file_formula = crate::formula_text::add_xlfn_prefixes(strip_leading_equals(f));
                inner.push_str(&format!(
                    r#"<formula>{}</formula>"#,
                    escape_text(&file_formula)
                ));
            }
            ("cellIs", inner, format!(r#" operator="{op}""#))
        }
        // Best-effort: skip rules we can't currently serialize.
        _ => return None,
    };

    Some(format!(
        r#"<cfRule type="{type_attr}"{extra_attrs}{attrs}>{body}</cfRule>"#
    ))
}

fn cell_is_operator_attr(op: CellIsOperator) -> &'static str {
    match op {
        CellIsOperator::GreaterThan => "greaterThan",
        CellIsOperator::GreaterThanOrEqual => "greaterThanOrEqual",
        CellIsOperator::LessThan => "lessThan",
        CellIsOperator::LessThanOrEqual => "lessThanOrEqual",
        CellIsOperator::Equal => "equal",
        CellIsOperator::NotEqual => "notEqual",
        CellIsOperator::Between => "between",
        CellIsOperator::NotBetween => "notBetween",
    }
}

fn render_sheet_protection(protection: &SheetProtection, prefix: Option<&str>) -> String {
    if !protection.enabled {
        return String::new();
    }

    // SpreadsheetML models most flags as allow-list booleans. `objects` and `scenarios` are
    // inverted "is protected" flags.
    let tag = crate::xml::prefixed_tag(prefix, "sheetProtection");
    let mut out = String::new();
    out.push('<');
    out.push_str(&tag);

    // Enable worksheet protection.
    out.push_str(r#" sheet="1""#);

    if let Some(hash) = protection.password_hash {
        out.push_str(&format!(r#" password="{:04X}""#, hash));
    }

    // Emit explicit values for all modeled attributes so round-tripping through our own reader is
    // unambiguous (especially for `objects`/`scenarios`, whose defaults vary across producers).
    out.push_str(&format!(
        r#" selectLockedCells="{}""#,
        if protection.select_locked_cells {
            "1"
        } else {
            "0"
        }
    ));
    out.push_str(&format!(
        r#" selectUnlockedCells="{}""#,
        if protection.select_unlocked_cells {
            "1"
        } else {
            "0"
        }
    ));
    out.push_str(&format!(
        r#" formatCells="{}""#,
        if protection.format_cells { "1" } else { "0" }
    ));
    out.push_str(&format!(
        r#" formatColumns="{}""#,
        if protection.format_columns { "1" } else { "0" }
    ));
    out.push_str(&format!(
        r#" formatRows="{}""#,
        if protection.format_rows { "1" } else { "0" }
    ));
    out.push_str(&format!(
        r#" insertColumns="{}""#,
        if protection.insert_columns { "1" } else { "0" }
    ));
    out.push_str(&format!(
        r#" insertRows="{}""#,
        if protection.insert_rows { "1" } else { "0" }
    ));
    out.push_str(&format!(
        r#" insertHyperlinks="{}""#,
        if protection.insert_hyperlinks {
            "1"
        } else {
            "0"
        }
    ));
    out.push_str(&format!(
        r#" deleteColumns="{}""#,
        if protection.delete_columns { "1" } else { "0" }
    ));
    out.push_str(&format!(
        r#" deleteRows="{}""#,
        if protection.delete_rows { "1" } else { "0" }
    ));
    out.push_str(&format!(
        r#" sort="{}""#,
        if protection.sort { "1" } else { "0" }
    ));
    out.push_str(&format!(
        r#" autoFilter="{}""#,
        if protection.auto_filter { "1" } else { "0" }
    ));
    out.push_str(&format!(
        r#" pivotTables="{}""#,
        if protection.pivot_tables { "1" } else { "0" }
    ));
    // Inverted "protected" flags.
    out.push_str(&format!(
        r#" objects="{}""#,
        if protection.edit_objects { "0" } else { "1" }
    ));
    out.push_str(&format!(
        r#" scenarios="{}""#,
        if protection.edit_scenarios { "0" } else { "1" }
    ));

    out.push_str("/>");
    out
}

fn patch_worksheet_xml(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    sheet: &Worksheet,
    original: &[u8],
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
) -> Result<Vec<u8>, WriteError> {
    let (original_has_dimension, original_used_range) = scan_worksheet_xml(original)?;
    let new_used_range = dimension::worksheet_used_range(sheet);
    let insert_dimension = !original_has_dimension && original_used_range != new_used_range;
    let dimension_range = dimension::worksheet_dimension_range(sheet);
    let dimension_ref = dimension_range.to_string();
    let patched = sheetdata_patch::patch_worksheet_xml(
        doc,
        sheet_meta,
        sheet,
        original,
        shared_lookup,
        style_to_xf,
        cell_meta_sheet_ids,
        changed_formula_cells,
    )?;

    let mut out = patch_worksheet_dimension(&patched, insert_dimension, dimension_range, &dimension_ref)?;

    // If conditional formatting was removed from the workbook model, ensure we strip any
    // corresponding worksheet XML blocks (both Office 2007 `<conditionalFormatting>` and the
    // x14 extension payload under `<ext uri="{78C0D931-...}">`).
    //
    // Note: conditional formatting parsing is best-effort. Only strip conditional formatting when
    // we can successfully parse *rules* from the original worksheet XML; otherwise preserve it for
    // high-fidelity round-trip (including unknown/unmodeled conditional formatting shapes).
    if sheet.conditional_formatting_rules.is_empty() {
        let original_xml = std::str::from_utf8(original).map_err(|e| {
            WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e))
        })?;
        if original_xml.contains("conditionalFormatting") {
            if let Ok(parsed) = crate::parse_worksheet_conditional_formatting_streaming(original_xml)
            {
                if !parsed.rules.is_empty() {
                    out = strip_worksheet_conditional_formatting_blocks(&out)?;
                }
            }
        }
    }

    Ok(out)
}

const X14_CF_EXT_URI: &str = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";

fn strip_worksheet_conditional_formatting_blocks(sheet_xml: &[u8]) -> Result<Vec<u8>, WriteError> {
    let mut reader = Reader::from_reader(sheet_xml);
    reader.config_mut().trim_text(false);

    let mut writer = Writer::new(Vec::with_capacity(sheet_xml.len()));
    let mut buf = Vec::new();

    let mut saw_root = false;
    // Depth of open elements *within* the worksheet root (root excluded).
    let mut depth: usize = 0;
    let mut skip_depth: usize = 0;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            _ if skip_depth > 0 => match event {
                Event::Start(_) => skip_depth += 1,
                Event::End(_) => skip_depth = skip_depth.saturating_sub(1),
                Event::Empty(_) => {}
                _ => {}
            },
            Event::Start(ref e) => {
                let local = local_name(e.name().into_inner());
                if !saw_root && local == b"worksheet" {
                    saw_root = true;
                    writer.write_event(Event::Start(e.to_owned()))?;
                } else if saw_root {
                    if depth == 0 && local == b"conditionalFormatting" {
                        skip_depth = 1;
                    } else if local == b"ext" && ext_uri_matches(e, X14_CF_EXT_URI)? {
                        skip_depth = 1;
                    } else {
                        writer.write_event(Event::Start(e.to_owned()))?;
                        depth = depth.saturating_add(1);
                    }
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
            }
            Event::Empty(ref e) => {
                let local = local_name(e.name().into_inner());
                if !saw_root && local == b"worksheet" {
                    saw_root = true;
                    writer.write_event(Event::Empty(e.to_owned()))?;
                } else if saw_root {
                    if depth == 0 && local == b"conditionalFormatting" {
                        // Skip empty conditional formatting blocks.
                    } else if local == b"ext" && ext_uri_matches(e, X14_CF_EXT_URI)? {
                        // Skip x14 conditional formatting extension blocks.
                    } else {
                        writer.write_event(Event::Empty(e.to_owned()))?;
                    }
                } else {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::End(ref e) => {
                let local = local_name(e.name().into_inner());
                if saw_root {
                    if depth == 0 && local == b"worksheet" {
                        writer.write_event(Event::End(e.to_owned()))?;
                        saw_root = false;
                    } else {
                        depth = depth.saturating_sub(1);
                        writer.write_event(Event::End(e.to_owned()))?;
                    }
                } else {
                    writer.write_event(Event::End(e.to_owned()))?;
                }
            }
            _ => {
                writer.write_event(event.to_owned())?;
            }
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

fn ext_uri_matches(e: &quick_xml::events::BytesStart<'_>, uri: &str) -> Result<bool, WriteError> {
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() != b"uri" {
            continue;
        }
        let value = attr.unescape_value()?.into_owned();
        return Ok(value.trim().eq_ignore_ascii_case(uri));
    }
    Ok(false)
}

fn patch_worksheet_dimension(
    worksheet_xml: &[u8],
    insert_dimension: bool,
    dimension_range: Range,
    dimension_ref: &str,
) -> Result<Vec<u8>, WriteError> {
    let mut reader = Reader::from_reader(worksheet_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut writer = Writer::new(Vec::with_capacity(
        worksheet_xml.len() + dimension_ref.len(),
    ));

    let mut inserted_dimension = false;
    let mut saw_sheet_pr = false;
    let mut in_sheet_pr = false;
    let mut worksheet_prefix: Option<String> = None;
    let mut worksheet_has_default_ns = false;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) if local_name(e.name().as_ref()) == b"worksheet" => {
                if worksheet_prefix.is_none() {
                    worksheet_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                    worksheet_has_default_ns = worksheet_has_default_spreadsheetml_ns(&e)?;
                }
                writer.write_event(Event::Start(e.into_owned()))?;
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"sheetPr" => {
                saw_sheet_pr = true;
                in_sheet_pr = true;
                writer.write_event(Event::Start(e.into_owned()))?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"sheetPr" => {
                saw_sheet_pr = true;
                if insert_dimension && !inserted_dimension {
                    let prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                    let tag = prefixed_tag(prefix.as_deref(), "dimension");
                    writer.write_event(Event::Empty(e.into_owned()))?;
                    insert_dimension_element(&mut writer, tag.as_str(), dimension_ref);
                    inserted_dimension = true;
                } else {
                    writer.write_event(Event::Empty(e.into_owned()))?;
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"sheetPr" => {
                in_sheet_pr = false;
                let prefix = element_prefix(e.name().as_ref())
                    .and_then(|p| std::str::from_utf8(p).ok())
                    .map(|s| s.to_string());
                writer.write_event(Event::End(e.into_owned()))?;
                if insert_dimension && !inserted_dimension {
                    let tag = prefixed_tag(prefix.as_deref(), "dimension");
                    insert_dimension_element(&mut writer, tag.as_str(), dimension_ref);
                    inserted_dimension = true;
                }
            }

            Event::Start(e) if local_name(e.name().as_ref()) == b"dimension" => {
                if dimension_matches(&e, dimension_range)? {
                    writer.write_event(Event::Start(e.into_owned()))?;
                } else {
                    write_dimension_element(&mut writer, &e, dimension_ref, false)?;
                }
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"dimension" => {
                if dimension_matches(&e, dimension_range)? {
                    writer.write_event(Event::Empty(e.into_owned()))?;
                } else {
                    write_dimension_element(&mut writer, &e, dimension_ref, true)?;
                }
            }

            Event::Start(e) if local_name(e.name().as_ref()) == b"sheetData" => {
                if insert_dimension && !inserted_dimension && !saw_sheet_pr && !in_sheet_pr {
                    let prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                    let tag = prefixed_tag(prefix.as_deref(), "dimension");
                    insert_dimension_element(&mut writer, tag.as_str(), dimension_ref);
                    inserted_dimension = true;
                }
                writer.write_event(Event::Start(e.into_owned()))?;
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"sheetData" => {
                if insert_dimension && !inserted_dimension && !saw_sheet_pr && !in_sheet_pr {
                    let prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                    let tag = prefixed_tag(prefix.as_deref(), "dimension");
                    insert_dimension_element(&mut writer, tag.as_str(), dimension_ref);
                    inserted_dimension = true;
                }
                writer.write_event(Event::Empty(e.into_owned()))?;
            }

            Event::Eof => break,
            ev => {
                match &ev {
                    Event::Start(e) | Event::Empty(e)
                        if insert_dimension
                            && !inserted_dimension
                            && !saw_sheet_pr
                            && !in_sheet_pr
                            && local_name(e.name().as_ref()) != b"worksheet" =>
                    {
                        let prefix = element_prefix(e.name().as_ref())
                            .and_then(|p| std::str::from_utf8(p).ok())
                            .map(|s| s.to_string());
                        let tag = prefixed_tag(prefix.as_deref(), "dimension");
                        insert_dimension_element(&mut writer, tag.as_str(), dimension_ref);
                        inserted_dimension = true;
                    }
                    Event::End(e)
                        if insert_dimension
                            && !inserted_dimension
                            && local_name(e.name().as_ref()) == b"worksheet" =>
                    {
                        let prefix = if worksheet_has_default_ns {
                            None
                        } else {
                            worksheet_prefix.as_deref()
                        };
                        let tag = prefixed_tag(prefix, "dimension");
                        insert_dimension_element(&mut writer, tag.as_str(), dimension_ref);
                        inserted_dimension = true;
                    }
                    _ => {}
                }
                writer.write_event(ev.into_owned())?;
            }
        }
        buf.clear();
    }

    Ok(writer.into_inner())
}

#[derive(Clone, Debug, PartialEq)]
struct ColXmlProps {
    width: Option<f32>,
    hidden: bool,
    outline_level: u8,
    collapsed: bool,
    style_xf: Option<u32>,
}

fn cols_xml_props_from_sheet(
    sheet: &Worksheet,
    style_to_xf: &HashMap<u32, u32>,
) -> BTreeMap<u32, ColXmlProps> {
    // OOXML column indices are 1-based. The model stores `col_properties` 0-based, and
    // `outline.cols` 1-based.
    let mut col_xml_props: BTreeMap<u32, ColXmlProps> = BTreeMap::new();
    for (col0, props) in sheet.col_properties.iter() {
        let col_1 = col0.saturating_add(1);
        if col_1 == 0 || col_1 > formula_model::EXCEL_MAX_COLS {
            continue;
        }
        // Preserve `style="0"` when the workbook's xf index 0 maps to a non-default style
        // (some producers place custom xfs at index 0).
        //
        // When xf 0 truly represents the default style, the style_id will be 0 and filtered out
        // above, so we won't emit a redundant `style="0"` in that case.
        let style_xf = props
            .style_id
            .filter(|style_id| *style_id != 0)
            .and_then(|style_id| style_to_xf.get(&style_id).copied());
        col_xml_props.insert(
            col_1,
            ColXmlProps {
                width: props.width,
                hidden: props.hidden,
                style_xf,
                outline_level: 0,
                collapsed: false,
            },
        );
    }

    for (col_1, entry) in sheet.outline.cols.iter() {
        if col_1 == 0 || col_1 > formula_model::EXCEL_MAX_COLS {
            continue;
        }
        if entry.level == 0 && !entry.hidden.is_hidden() && !entry.collapsed {
            continue;
        }
        col_xml_props
            .entry(col_1)
            .and_modify(|props| {
                props.outline_level = entry.level;
                props.collapsed = entry.collapsed;
                props.hidden |= entry.hidden.is_hidden();
            })
            .or_insert_with(|| ColXmlProps {
                width: None,
                hidden: entry.hidden.is_hidden(),
                outline_level: entry.level,
                collapsed: entry.collapsed,
                style_xf: None,
            });
    }

    col_xml_props
}

fn render_cols(sheet: &Worksheet, prefix: Option<&str>, style_to_xf: &HashMap<u32, u32>) -> String {
    let cols_tag = crate::xml::prefixed_tag(prefix, "cols");
    let col_tag = crate::xml::prefixed_tag(prefix, "col");

    let col_xml_props = cols_xml_props_from_sheet(sheet, style_to_xf);

    if col_xml_props.is_empty() {
        return String::new();
    }

    let mut out = String::new();
    out.push('<');
    out.push_str(&cols_tag);
    out.push('>');

    let mut current: Option<(u32, u32, ColXmlProps)> = None;
    for (&col, props) in col_xml_props.iter() {
        let props = props.clone();
        match current.take() {
            None => current = Some((col, col, props)),
            Some((start, end, cur)) if col == end + 1 && props == cur => {
                current = Some((start, col, cur));
            }
            Some((start, end, cur)) => {
                out.push_str(&render_col_range(&col_tag, start, end, &cur));
                current = Some((col, col, props));
            }
        }
    }
    if let Some((start, end, cur)) = current {
        out.push_str(&render_col_range(&col_tag, start, end, &cur));
    }

    out.push_str("</");
    out.push_str(&cols_tag);
    out.push('>');
    out
}

fn render_col_range(
    col_tag: &str,
    start_col_1: u32,
    end_col_1: u32,
    props: &ColXmlProps,
) -> String {
    let mut s = String::new();
    let min = start_col_1;
    let max = end_col_1;
    s.push_str(&format!(r#"<{col_tag} min="{min}" max="{max}""#));
    if let Some(width) = props.width {
        // `f32::to_string()` prints `-0.0` as `-0`; normalize for XML stability.
        let width = if width == 0.0 { 0.0 } else { width };
        s.push_str(&format!(r#" width="{width}""#));
        s.push_str(r#" customWidth="1""#);
    }
    if let Some(style_xf) = props.style_xf {
        s.push_str(&format!(r#" style="{style_xf}""#));
        s.push_str(r#" customFormat="1""#);
    }
    if props.hidden {
        s.push_str(r#" hidden="1""#);
    }
    if props.outline_level > 0 {
        s.push_str(&format!(r#" outlineLevel="{}""#, props.outline_level));
    }
    if props.collapsed {
        s.push_str(r#" collapsed="1""#);
    }
    s.push_str("/>");
    s
}

fn scan_worksheet_xml(original: &[u8]) -> Result<(bool, Option<Range>), WriteError> {
    let mut reader = Reader::from_reader(original);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    let mut in_sheet_data = false;
    let mut min_cell: Option<CellRef> = None;
    let mut max_cell: Option<CellRef> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e) if local_name(e.name().as_ref()) == b"dimension" => {
                // If the worksheet already has a <dimension> element, we don't need to scan the
                // potentially-large <sheetData> section just to decide whether to insert one.
                return Ok((true, None));
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"sheetData" => {
                in_sheet_data = true
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"sheetData" => in_sheet_data = false,
            Event::Empty(e) if local_name(e.name().as_ref()) == b"sheetData" => {
                in_sheet_data = false;
                drop(e);
            }
            Event::Start(e) | Event::Empty(e)
                if in_sheet_data && local_name(e.name().as_ref()) == b"c" =>
            {
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() != b"r" {
                        continue;
                    }
                    let a1 = attr.unescape_value()?.into_owned();
                    let Ok(cell_ref) = CellRef::from_a1(&a1) else {
                        continue;
                    };
                    min_cell = Some(match min_cell {
                        Some(min) => {
                            CellRef::new(min.row.min(cell_ref.row), min.col.min(cell_ref.col))
                        }
                        None => cell_ref,
                    });
                    max_cell = Some(match max_cell {
                        Some(max) => {
                            CellRef::new(max.row.max(cell_ref.row), max.col.max(cell_ref.col))
                        }
                        None => cell_ref,
                    });
                    break;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }

    let used_range = match (min_cell, max_cell) {
        (Some(start), Some(end)) => Some(Range::new(start, end)),
        _ => None,
    };
    Ok((false, used_range))
}

fn insert_dimension_element(
    writer: &mut Writer<Vec<u8>>,
    dimension_tag: &str,
    dimension_ref: &str,
) {
    writer.get_mut().push(b'<');
    writer.get_mut().extend_from_slice(dimension_tag.as_bytes());
    writer.get_mut().extend_from_slice(b" ref=\"");
    writer
        .get_mut()
        .extend_from_slice(escape_attr(dimension_ref).as_bytes());
    writer.get_mut().extend_from_slice(b"\"/>");
}

fn dimension_matches(
    e: &quick_xml::events::BytesStart<'_>,
    expected: Range,
) -> Result<bool, WriteError> {
    let mut ref_value = None;
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"ref" {
            ref_value = Some(attr.unescape_value()?.into_owned());
            break;
        }
    }
    let Some(ref_value) = ref_value else {
        return Ok(false);
    };
    Ok(dimension::parse_dimension_ref(&ref_value) == Some(expected))
}

fn write_dimension_element(
    writer: &mut Writer<Vec<u8>>,
    e: &quick_xml::events::BytesStart<'_>,
    dimension_ref: &str,
    is_empty: bool,
) -> Result<(), WriteError> {
    writer.get_mut().push(b'<');
    writer.get_mut().extend_from_slice(e.name().as_ref());
    let mut wrote_ref = false;
    for attr in e.attributes() {
        let attr = attr?;
        writer.get_mut().push(b' ');
        writer.get_mut().extend_from_slice(attr.key.as_ref());
        writer.get_mut().extend_from_slice(b"=\"");
        if attr.key.as_ref() == b"ref" {
            wrote_ref = true;
            writer
                .get_mut()
                .extend_from_slice(escape_attr(dimension_ref).as_bytes());
        } else {
            writer
                .get_mut()
                .extend_from_slice(escape_attr(&attr.unescape_value()?.into_owned()).as_bytes());
        }
        writer.get_mut().push(b'"');
    }

    if !wrote_ref {
        writer.get_mut().extend_from_slice(b" ref=\"");
        writer
            .get_mut()
            .extend_from_slice(escape_attr(dimension_ref).as_bytes());
        writer.get_mut().push(b'"');
    }

    if is_empty {
        writer.get_mut().extend_from_slice(b"/>");
    } else {
        writer.get_mut().push(b'>');
    }
    Ok(())
}

fn render_sheet_data(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    sheet: &Worksheet,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    outline: Option<&Outline>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
) -> String {
    let shared_formulas = shared_formula_groups(doc, sheet_meta.worksheet_id);

    if let Some((origin, rows, cols)) = sheet.columnar_table_extent() {
        if sheet.columnar_table().is_some() {
            return render_sheet_data_columnar(
                doc,
                sheet_meta,
                sheet,
                shared_lookup,
                style_to_xf,
                cell_meta_sheet_ids,
                outline,
                &shared_formulas,
                changed_formula_cells,
                origin,
                rows,
                cols,
            );
        }
    }

    let mut out = String::new();
    out.push_str("<sheetData>");

    let mut cells: Vec<(CellRef, &formula_model::Cell)> = sheet.iter_cells().collect();
    cells.sort_by_key(|(r, _)| (r.row, r.col));

    let mut rows: BTreeMap<u32, ()> = BTreeMap::new();
    for (cell_ref, _) in &cells {
        rows.insert(cell_ref.row + 1, ());
    }
    for (row, props) in sheet.row_properties.iter() {
        if props.height.is_some() || props.hidden || props.style_id.is_some_and(|id| id != 0) {
            rows.insert(row + 1, ());
        }
    }

    let mut outline_rows: Vec<u32> = Vec::new();
    if let Some(outline) = outline {
        // Preserve outline-only rows (groups, hidden rows, etc) even if they contain no cells.
        // We don't attempt to preserve all row-level metadata yet—only the outline-related attrs.
        for (row, entry) in outline.rows.iter() {
            if entry.level > 0 || entry.hidden.is_hidden() || entry.collapsed {
                outline_rows.push(row);
                rows.insert(row, ());
            }
        }
    }

    let mut cell_idx = 0usize;
    let mut outline_idx = 0usize;

    for row_1_based in rows.keys().copied() {
        // Keep the existing outline-only row tracking, but the actual row list is now the union of:
        // - cell rows
        // - outline rows
        // - modeled row properties
        if outline_rows.get(outline_idx).copied() == Some(row_1_based) {
            outline_idx += 1;
        }

        let outline_entry: OutlineEntry = outline
            .map(|outline| outline.rows.entry(row_1_based))
            .unwrap_or_default();
        let row_props = sheet.row_properties(row_1_based.saturating_sub(1));

        out.push_str(&format!(r#"<row r="{row_1_based}""#));
        if let Some(row_props) = row_props {
            if let Some(height) = row_props.height {
                // `f32::to_string()` prints `-0.0` as `-0`; normalize for XML stability.
                let height = if height == 0.0 { 0.0 } else { height };
                out.push_str(&format!(r#" ht="{height}""#));
                out.push_str(r#" customHeight="1""#);
            }
            if row_props.hidden {
                out.push_str(r#" hidden="1""#);
            }
            if let Some(style_id) = row_props.style_id.filter(|id| *id != 0) {
                // Preserve `s="0"` when the workbook's xf index 0 maps to a non-default style
                // (some producers place custom xfs at index 0).
                //
                // When xf 0 truly represents the default style, the style_id will be 0 and filtered
                // out above, so we won't emit a redundant `s="0"` in that case.
                if let Some(style_xf) = style_to_xf.get(&style_id).copied() {
                    out.push_str(&format!(r#" s="{style_xf}" customFormat="1""#));
                }
            }
        }
        if outline_entry.level > 0 {
            out.push_str(&format!(r#" outlineLevel="{}""#, outline_entry.level));
        }
        if outline_entry.hidden.is_hidden() && !row_props.is_some_and(|p| p.hidden) {
            out.push_str(r#" hidden="1""#);
        }
        if outline_entry.collapsed {
            out.push_str(r#" collapsed="1""#);
        }

        let mut wrote_any_cell = false;

        while let Some((cell_ref, cell)) = cells.get(cell_idx).copied() {
            if cell_ref.row + 1 != row_1_based {
                break;
            }
            if !wrote_any_cell {
                out.push('>');
                wrote_any_cell = true;
            }
            cell_idx += 1;
            append_cell_xml(
                &mut out,
                doc,
                sheet_meta,
                sheet,
                cell_ref,
                cell,
                shared_lookup,
                style_to_xf,
                cell_meta_sheet_ids,
                &shared_formulas,
                changed_formula_cells,
            );
        }

        if wrote_any_cell {
            out.push_str("</row>");
        } else {
            out.push_str("/>");
        }
    }
    out.push_str("</sheetData>");
    out
}

fn render_sheet_data_columnar(
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    sheet: &Worksheet,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    outline: Option<&Outline>,
    shared_formulas: &HashMap<u32, SharedFormulaGroup>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
    origin: CellRef,
    table_rows: usize,
    table_cols: usize,
) -> String {
    let Some(columnar) = sheet.columnar_table() else {
        return String::from("<sheetData></sheetData>");
    };
    let table = columnar.as_ref();

    let mut out = String::new();
    out.push_str("<sheetData>");

    // Sparse overlay cells in row-major order.
    let mut overlay_cells: Vec<(CellRef, &formula_model::Cell)> = sheet.iter_cells().collect();
    overlay_cells.sort_by_key(|(r, _)| (r.row, r.col));
    let mut overlay_idx = 0usize;

    // Rows to emit outside the contiguous table row range (row properties, outline-only rows,
    // and rows containing overlay cells).
    let mut extra_rows: BTreeMap<u32, ()> = BTreeMap::new();
    for (cell_ref, _) in &overlay_cells {
        extra_rows.insert(cell_ref.row + 1, ());
    }
    for (row, props) in sheet.row_properties.iter() {
        if props.height.is_some() || props.hidden || props.style_id.is_some_and(|id| id != 0) {
            extra_rows.insert(row + 1, ());
        }
    }
    if let Some(outline) = outline {
        // Preserve outline-only rows (groups, hidden rows, etc) even if they contain no cells.
        // We don't attempt to preserve all row-level metadata yet—only the outline-related attrs.
        for (row, entry) in outline.rows.iter() {
            if entry.level > 0 || entry.hidden.is_hidden() || entry.collapsed {
                extra_rows.insert(row, ());
            }
        }
    }
    let extra_rows: Vec<u32> = extra_rows.keys().copied().collect();
    let mut extra_idx = 0usize;

    let table_row_start_1 = origin.row.saturating_add(1);
    let table_row_end_1 = origin
        .row
        .saturating_add(table_rows.saturating_sub(1) as u32)
        .saturating_add(1);

    let mut table_row_1 = table_row_start_1;

    loop {
        let next_extra = extra_rows.get(extra_idx).copied();
        let next_table = (table_row_1 <= table_row_end_1).then_some(table_row_1);

        let Some(row_1_based) = (match (next_table, next_extra) {
            (Some(t), Some(e)) => Some(t.min(e)),
            (Some(t), None) => Some(t),
            (None, Some(e)) => Some(e),
            (None, None) => None,
        }) else {
            break;
        };

        if next_extra == Some(row_1_based) {
            extra_idx += 1;
        }
        if next_table == Some(row_1_based) {
            table_row_1 = table_row_1.saturating_add(1);
        }

        let row_zero = row_1_based.saturating_sub(1);

        // Gather overlay cells for this row.
        while overlay_idx < overlay_cells.len()
            && overlay_cells[overlay_idx].0.row.saturating_add(1) < row_1_based
        {
            overlay_idx += 1;
        }
        let overlay_start = overlay_idx;
        while overlay_idx < overlay_cells.len()
            && overlay_cells[overlay_idx].0.row.saturating_add(1) == row_1_based
        {
            overlay_idx += 1;
        }
        let overlay_slice = &overlay_cells[overlay_start..overlay_idx];

        let outline_entry: OutlineEntry = outline
            .map(|outline| outline.rows.entry(row_1_based))
            .unwrap_or_default();
        let row_props = sheet.row_properties(row_zero);

        // Pre-render the row attributes so we can skip truly-empty rows (no attrs and no cells).
        let mut row_attrs = String::new();
        if let Some(row_props) = row_props {
            if let Some(height) = row_props.height {
                // `f32::to_string()` prints `-0.0` as `-0`; normalize for XML stability.
                let height = if height == 0.0 { 0.0 } else { height };
                row_attrs.push_str(&format!(r#" ht="{height}""#));
                row_attrs.push_str(r#" customHeight="1""#);
            }
            if row_props.hidden {
                row_attrs.push_str(r#" hidden="1""#);
            }
            if let Some(style_id) = row_props.style_id.filter(|id| *id != 0) {
                // Preserve `s="0"` when the workbook's xf index 0 maps to a non-default style
                // (some producers place custom xfs at index 0).
                //
                // When xf 0 truly represents the default style, the style_id will be 0 and filtered
                // out above, so we won't emit a redundant `s="0"` in that case.
                if let Some(style_xf) = style_to_xf.get(&style_id).copied() {
                    row_attrs.push_str(&format!(r#" s="{style_xf}" customFormat="1""#));
                }
            }
        }
        if outline_entry.level > 0 {
            row_attrs.push_str(&format!(r#" outlineLevel="{}""#, outline_entry.level));
        }
        if outline_entry.hidden.is_hidden() && !row_props.is_some_and(|p| p.hidden) {
            row_attrs.push_str(r#" hidden="1""#);
        }
        if outline_entry.collapsed {
            row_attrs.push_str(r#" collapsed="1""#);
        }

        let mut cells_xml = String::new();
        let mut wrote_any_cell = false;

        let in_table_row =
            row_zero >= origin.row && row_zero < origin.row.saturating_add(table_rows as u32);

        if in_table_row {
            let row_off = (row_zero - origin.row) as usize;

            let mut overlay_cell_idx = 0usize;

            // Overlay cells left of the table.
            while overlay_cell_idx < overlay_slice.len()
                && overlay_slice[overlay_cell_idx].0.col < origin.col
            {
                let (cell_ref, cell) = overlay_slice[overlay_cell_idx];
                append_cell_xml(
                    &mut cells_xml,
                    doc,
                    sheet_meta,
                    sheet,
                    cell_ref,
                    cell,
                    shared_lookup,
                    style_to_xf,
                    cell_meta_sheet_ids,
                    shared_formulas,
                    changed_formula_cells,
                );
                overlay_cell_idx += 1;
                wrote_any_cell = true;
            }

            // Table columns (overlay overrides).
            for col_off in 0..table_cols {
                let col_idx = origin.col.saturating_add(col_off as u32);
                if overlay_cell_idx < overlay_slice.len()
                    && overlay_slice[overlay_cell_idx].0.col == col_idx
                {
                    let (cell_ref, cell) = overlay_slice[overlay_cell_idx];
                    append_cell_xml(
                        &mut cells_xml,
                        doc,
                        sheet_meta,
                        sheet,
                        cell_ref,
                        cell,
                        shared_lookup,
                        style_to_xf,
                        cell_meta_sheet_ids,
                        shared_formulas,
                        changed_formula_cells,
                    );
                    overlay_cell_idx += 1;
                    wrote_any_cell = true;
                    continue;
                }

                let cell_ref = CellRef::new(row_zero, col_idx);
                if sheet.merged_regions.resolve_cell(cell_ref) != cell_ref {
                    continue;
                }

                let value = table.get_cell(row_off, col_off);
                let col_type = table
                    .schema()
                    .get(col_off)
                    .map(|s| s.column_type)
                    .unwrap_or(ColumnarType::String);
                let cell_value = columnar_to_cell_value(value, col_type);
                if matches!(cell_value, CellValue::Empty) {
                    continue;
                }
                let cell = formula_model::Cell::new(cell_value);
                append_cell_xml(
                    &mut cells_xml,
                    doc,
                    sheet_meta,
                    sheet,
                    cell_ref,
                    &cell,
                    shared_lookup,
                    style_to_xf,
                    cell_meta_sheet_ids,
                    shared_formulas,
                    changed_formula_cells,
                );
                wrote_any_cell = true;
            }

            // Overlay cells right of the table.
            while overlay_cell_idx < overlay_slice.len() {
                let (cell_ref, cell) = overlay_slice[overlay_cell_idx];
                append_cell_xml(
                    &mut cells_xml,
                    doc,
                    sheet_meta,
                    sheet,
                    cell_ref,
                    cell,
                    shared_lookup,
                    style_to_xf,
                    cell_meta_sheet_ids,
                    shared_formulas,
                    changed_formula_cells,
                );
                overlay_cell_idx += 1;
                wrote_any_cell = true;
            }
        } else {
            // Row outside the columnar table; only overlay cells apply.
            for (cell_ref, cell) in overlay_slice {
                append_cell_xml(
                    &mut cells_xml,
                    doc,
                    sheet_meta,
                    sheet,
                    *cell_ref,
                    cell,
                    shared_lookup,
                    style_to_xf,
                    cell_meta_sheet_ids,
                    shared_formulas,
                    changed_formula_cells,
                );
                wrote_any_cell = true;
            }
        }

        if !wrote_any_cell && row_attrs.is_empty() {
            continue;
        }

        out.push_str(&format!(r#"<row r="{row_1_based}""#));
        out.push_str(&row_attrs);
        if wrote_any_cell {
            out.push('>');
            out.push_str(&cells_xml);
            out.push_str("</row>");
        } else {
            out.push_str("/>");
        }
    }

    out.push_str("</sheetData>");
    out
}

fn columnar_to_cell_value(value: ColumnarValue, column_type: ColumnarType) -> CellValue {
    match value {
        ColumnarValue::Null => CellValue::Empty,
        ColumnarValue::Number(v) => CellValue::Number(v),
        ColumnarValue::Boolean(v) => CellValue::Boolean(v),
        ColumnarValue::String(v) => CellValue::String(v.as_ref().to_string()),
        ColumnarValue::DateTime(v) => CellValue::Number(v as f64),
        ColumnarValue::Currency(v) => match column_type {
            ColumnarType::Currency { scale } => {
                let denom = 10f64.powi(scale as i32);
                CellValue::Number(v as f64 / denom)
            }
            _ => CellValue::Number(v as f64),
        },
        ColumnarValue::Percentage(v) => match column_type {
            ColumnarType::Percentage { scale } => {
                let denom = 10f64.powi(scale as i32);
                CellValue::Number(v as f64 / denom)
            }
            _ => CellValue::Number(v as f64),
        },
    }
}

fn append_cell_xml(
    out: &mut String,
    doc: &XlsxDocument,
    sheet_meta: &SheetMeta,
    _sheet: &Worksheet,
    cell_ref: CellRef,
    cell: &formula_model::Cell,
    shared_lookup: &HashMap<SharedStringKey, u32>,
    style_to_xf: &HashMap<u32, u32>,
    cell_meta_sheet_ids: &HashMap<WorksheetId, WorksheetId>,
    shared_formulas: &HashMap<u32, SharedFormulaGroup>,
    changed_formula_cells: &HashSet<(WorksheetId, CellRef)>,
) {
    out.push_str(r#"<c r=""#);
    out.push_str(&cell_ref.to_a1());
    out.push('"');

    if cell.style_id != 0 {
        if let Some(xf_index) = style_to_xf.get(&cell.style_id) {
            out.push_str(&format!(r#" s="{xf_index}""#));
        }
    }

    let meta = lookup_cell_meta(doc, cell_meta_sheet_ids, sheet_meta.worksheet_id, cell_ref);
    let value_kind = effective_value_kind(meta, cell);

    let meta_sheet_id = cell_meta_sheet_ids
        .get(&sheet_meta.worksheet_id)
        .copied()
        .unwrap_or(sheet_meta.worksheet_id);
    let clear_cached_value = cell
        .formula
        .as_deref()
        .is_some_and(|f| !strip_leading_equals(f).is_empty())
        && changed_formula_cells.contains(&(meta_sheet_id, cell_ref));
    let has_value = !clear_cached_value && !matches!(cell.value, CellValue::Empty);

    if has_value {
        match &value_kind {
            CellValueKind::SharedString { .. } => out.push_str(r#" t="s""#),
            CellValueKind::InlineString => out.push_str(r#" t="inlineStr""#),
            CellValueKind::Bool => out.push_str(r#" t="b""#),
            CellValueKind::Error => out.push_str(r#" t="e""#),
            CellValueKind::Str => out.push_str(r#" t="str""#),
            CellValueKind::Number => {}
            CellValueKind::Other { t } => {
                out.push_str(&format!(r#" t="{}""#, escape_attr(t)));
            }
        }
    }
    // SpreadsheetML cell metadata pointers.
    //
    // Excel emits `vm`/`cm` attributes on `<c>` elements to reference value metadata and cell
    // metadata records (used for modern features like linked data types / rich values).
    if let Some(vm) = meta.and_then(|m| m.vm.as_deref()).filter(|s| !s.is_empty()) {
        // `vm="..."` is a SpreadsheetML value-metadata pointer (typically into `xl/metadata*.xml`).
        // If the cell's value changes and we can't update the corresponding metadata records,
        // drop `vm` to avoid leaving a dangling reference.
        let preserve_vm = matches!(cell.value, CellValue::Error(ErrorValue::Value))
            || meta.is_some_and(|m| m.raw_value.is_none())
            || match (&cell.value, meta.and_then(|m| m.raw_value.as_deref())) {
                (CellValue::Number(n), Some(raw)) => raw.parse::<f64>().ok() == Some(*n),
                (CellValue::Boolean(b), Some(raw)) => (raw == "1" && *b) || (raw == "0" && !*b),
                (CellValue::Error(err), Some(raw)) => raw == err.as_str(),
                (CellValue::String(s), Some(raw)) => raw == s,
                (CellValue::RichText(rich), Some(raw)) => raw == rich.text,
                (CellValue::Entity(entity), Some(raw)) => raw == entity.display_value,
                (CellValue::Record(record), Some(raw)) => raw == record.to_string(),
                (CellValue::Image(image), Some(raw)) => image
                    .alt_text
                    .as_deref()
                    .filter(|s| !s.is_empty())
                    .is_some_and(|alt| raw == alt),
                _ => false,
            };
        if preserve_vm {
            out.push_str(&format!(r#" vm="{}""#, escape_attr(vm)));
        }
    }
    if let Some(cm) = meta.and_then(|m| m.cm.as_deref()).filter(|s| !s.is_empty()) {
        out.push_str(&format!(r#" cm="{}""#, escape_attr(cm)));
    }
    out.push('>');

    let model_formula = cell.formula.as_deref();
    let mut preserve_textless_shared = false;
    let mut formula_meta = match (model_formula, meta.and_then(|m| m.formula.clone())) {
        (Some(_), Some(meta)) => Some(meta),
        (Some(formula), None) => Some(crate::FormulaMeta {
            file_text: crate::formula_text::add_xlfn_prefixes(strip_leading_equals(formula)),
            ..Default::default()
        }),
        (None, Some(meta)) => {
            // The in-memory model doesn't currently represent shared formulas for follower
            // cells. Preserve those formulas when the stored SpreadsheetML indicates a formula
            // even if the model omits it.
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
                // Model cleared the formula; don't keep stale formula text from metadata.
                None
            }
        }
        (None, None) => None,
    };

    if let (Some(display), Some(meta)) = (model_formula, formula_meta.as_mut()) {
        if meta.t.as_deref() == Some("shared") && meta.file_text.is_empty() {
            if let Some(si) = meta.shared_index {
                if let Some(expected) = shared_formula_expected(shared_formulas, si, cell_ref) {
                    if expected == strip_leading_equals(display) {
                        preserve_textless_shared = true;
                    } else {
                        // The cell's model formula differs from the shared-formula expansion,
                        // so break sharing and store the explicit formula text.
                        meta.t = None;
                        meta.reference = None;
                        meta.shared_index = None;
                        meta.file_text =
                            crate::formula_text::add_xlfn_prefixes(strip_leading_equals(display));
                    }
                } else {
                    // Without the shared-formula master we can't validate equivalence, so
                    // prefer preserving the model formula over keeping the shared structure.
                    meta.t = None;
                    meta.reference = None;
                    meta.shared_index = None;
                    meta.file_text =
                        crate::formula_text::add_xlfn_prefixes(strip_leading_equals(display));
                }
            } else {
                // Malformed shared-formula follower; fall back to a normal formula.
                meta.t = None;
                meta.reference = None;
                meta.shared_index = None;
                meta.file_text =
                    crate::formula_text::add_xlfn_prefixes(strip_leading_equals(display));
            }
        }
    }

    if let Some(formula_meta) = formula_meta {
        out.push_str("<f");
        if let Some(t) = &formula_meta.t {
            out.push_str(&format!(r#" t="{}""#, escape_attr(t)));
        }
        if let Some(r) = &formula_meta.reference {
            out.push_str(&format!(r#" ref="{}""#, escape_attr(r)));
        }
        if let Some(si) = formula_meta.shared_index {
            out.push_str(&format!(r#" si="{si}""#));
        }
        if let Some(aca) = formula_meta.always_calc {
            out.push_str(&format!(r#" aca="{}""#, if aca { "1" } else { "0" }));
        }

        let file_text = if preserve_textless_shared {
            String::new()
        } else {
            formula_file_text(&formula_meta, model_formula)
        };
        if file_text.is_empty() {
            out.push_str("/>");
        } else {
            out.push('>');
            out.push_str(&escape_text(&file_text));
            out.push_str("</f>");
        }
    }

    if !clear_cached_value {
        match &cell.value {
            CellValue::Empty => {}
            value @ CellValue::String(s) if matches!(&value_kind, CellValueKind::Other { .. }) => {
                out.push_str("<v>");
                out.push_str(&escape_text(&raw_or_other(meta, s)));
                out.push_str("</v>");
            }
            CellValue::Number(n) => {
                out.push_str("<v>");
                out.push_str(&escape_text(&raw_or_number(meta, *n)));
                out.push_str("</v>");
            }
            CellValue::Boolean(b) => {
                out.push_str("<v>");
                out.push_str(raw_or_bool(meta, *b));
                out.push_str("</v>");
            }
            CellValue::Error(err) => {
                out.push_str("<v>");
                out.push_str(&escape_text(&raw_or_error(meta, *err)));
                out.push_str("</v>");
            }
            value @ CellValue::String(s) => match &value_kind {
                CellValueKind::SharedString { .. } => {
                    let idx = shared_string_index(doc, meta, value, shared_lookup);
                    out.push_str("<v>");
                    out.push_str(&idx.to_string());
                    out.push_str("</v>");
                }
                CellValueKind::InlineString => {
                    out.push_str("<is><t");
                    if needs_space_preserve(s) {
                        out.push_str(r#" xml:space="preserve""#);
                    }
                    out.push('>');
                    out.push_str(&escape_text(s));
                    out.push_str("</t>");
                    if let Some(phonetic) = cell.phonetic.as_deref() {
                        let base_len = s.chars().count();
                        out.push_str(&format!(r#"<rPh sb="0" eb="{base_len}"><t"#));
                        if needs_space_preserve(phonetic) {
                            out.push_str(r#" xml:space="preserve""#);
                        }
                        out.push('>');
                        out.push_str(&escape_text(phonetic));
                        out.push_str("</t></rPh>");
                    }
                    out.push_str("</is>");
                }
                CellValueKind::Str => {
                    out.push_str("<v>");
                    out.push_str(&escape_text(&raw_or_str(meta, s)));
                    out.push_str("</v>");
                }
                _ => {
                    // Fallback: treat as shared string.
                    let idx = shared_string_index(doc, meta, value, shared_lookup);
                    out.push_str("<v>");
                    out.push_str(&idx.to_string());
                    out.push_str("</v>");
                }
            },
            value @ CellValue::Entity(entity) => {
                let s = entity.display_value.as_str();
                match &value_kind {
                    CellValueKind::SharedString { .. } => {
                        let idx = shared_string_index(doc, meta, value, shared_lookup);
                        out.push_str("<v>");
                        out.push_str(&idx.to_string());
                        out.push_str("</v>");
                    }
                    CellValueKind::InlineString => {
                        out.push_str("<is><t");
                        if needs_space_preserve(s) {
                            out.push_str(r#" xml:space="preserve""#);
                        }
                        out.push('>');
                        out.push_str(&escape_text(s));
                        out.push_str("</t>");
                        if let Some(phonetic) = cell.phonetic.as_deref() {
                            let base_len = s.chars().count();
                            out.push_str(&format!(r#"<rPh sb="0" eb="{base_len}"><t"#));
                            if needs_space_preserve(phonetic) {
                                out.push_str(r#" xml:space="preserve""#);
                            }
                            out.push('>');
                            out.push_str(&escape_text(phonetic));
                            out.push_str("</t></rPh>");
                        }
                        out.push_str("</is>");
                    }
                    CellValueKind::Str => {
                        out.push_str("<v>");
                        out.push_str(&escape_text(&raw_or_str(meta, s)));
                        out.push_str("</v>");
                    }
                    _ => {
                        // Fallback: treat as shared string.
                        let idx = shared_string_index(doc, meta, value, shared_lookup);
                        out.push_str("<v>");
                        out.push_str(&idx.to_string());
                        out.push_str("</v>");
                    }
                }
            }
            value @ CellValue::Record(record) => {
                let s = record.to_string();
                match &value_kind {
                    CellValueKind::SharedString { .. } => {
                        let idx = shared_string_index(doc, meta, value, shared_lookup);
                        out.push_str("<v>");
                        out.push_str(&idx.to_string());
                        out.push_str("</v>");
                    }
                    CellValueKind::InlineString => {
                        out.push_str("<is><t");
                        if needs_space_preserve(&s) {
                            out.push_str(r#" xml:space="preserve""#);
                        }
                        out.push('>');
                        out.push_str(&escape_text(&s));
                        out.push_str("</t>");
                        if let Some(phonetic) = cell.phonetic.as_deref() {
                            let base_len = s.chars().count();
                            out.push_str(&format!(r#"<rPh sb="0" eb="{base_len}"><t"#));
                            if needs_space_preserve(phonetic) {
                                out.push_str(r#" xml:space="preserve""#);
                            }
                            out.push('>');
                            out.push_str(&escape_text(phonetic));
                            out.push_str("</t></rPh>");
                        }
                        out.push_str("</is>");
                    }
                    CellValueKind::Str => {
                        out.push_str("<v>");
                        out.push_str(&escape_text(&raw_or_str(meta, &s)));
                        out.push_str("</v>");
                    }
                    _ => {
                        // Fallback: treat as shared string.
                        let idx = shared_string_index(doc, meta, value, shared_lookup);
                        out.push_str("<v>");
                        out.push_str(&idx.to_string());
                        out.push_str("</v>");
                    }
                }
            }
            value @ CellValue::Image(image) => {
                if let Some(alt) = image.alt_text.as_deref().filter(|s| !s.is_empty()) {
                    match &value_kind {
                        CellValueKind::SharedString { .. } => {
                            let idx = shared_string_index(doc, meta, value, shared_lookup);
                            out.push_str("<v>");
                            out.push_str(&idx.to_string());
                            out.push_str("</v>");
                        }
                        CellValueKind::InlineString => {
                            out.push_str("<is><t");
                            if needs_space_preserve(alt) {
                                out.push_str(r#" xml:space="preserve""#);
                            }
                            out.push('>');
                            out.push_str(&escape_text(alt));
                            out.push_str("</t>");
                            if let Some(phonetic) = cell.phonetic.as_deref() {
                                let base_len = alt.chars().count();
                                out.push_str(&format!(r#"<rPh sb="0" eb="{base_len}"><t"#));
                                if needs_space_preserve(phonetic) {
                                    out.push_str(r#" xml:space="preserve""#);
                                }
                                out.push('>');
                                out.push_str(&escape_text(phonetic));
                                out.push_str("</t></rPh>");
                            }
                            out.push_str("</is>");
                        }
                        CellValueKind::Str => {
                            out.push_str("<v>");
                            out.push_str(&escape_text(&raw_or_str(meta, alt)));
                            out.push_str("</v>");
                        }
                        _ => {
                            // Fallback: treat as shared string.
                            let idx = shared_string_index(doc, meta, value, shared_lookup);
                            out.push_str("<v>");
                            out.push_str(&idx.to_string());
                            out.push_str("</v>");
                        }
                    }
                }
            }
            value @ CellValue::RichText(rich) => {
                // Rich text is stored in the shared strings table.
                let idx = shared_string_index(doc, meta, value, shared_lookup);
                if idx != 0 || !rich.text.is_empty() {
                    out.push_str("<v>");
                    out.push_str(&idx.to_string());
                    out.push_str("</v>");
                }
            }
            _ => {
                // Array/Spill not yet modeled for writing. Preserve as blank.
            }
        }
    }

    out.push_str("</c>");
}

fn infer_value_kind(cell: &formula_model::Cell) -> CellValueKind {
    match &cell.value {
        CellValue::Boolean(_) => CellValueKind::Bool,
        CellValue::Error(_) => CellValueKind::Error,
        CellValue::Number(_) => CellValueKind::Number,
        CellValue::String(_) => CellValueKind::SharedString { index: 0 },
        CellValue::RichText(_) => CellValueKind::SharedString { index: 0 },
        CellValue::Entity(_) | CellValue::Record(_) => CellValueKind::SharedString { index: 0 },
        CellValue::Image(image) if image.alt_text.as_deref().is_some_and(|s| !s.is_empty()) => {
            CellValueKind::SharedString { index: 0 }
        }
        CellValue::Empty => CellValueKind::Number,
        _ => CellValueKind::Number,
    }
}

fn effective_value_kind(
    meta: Option<&crate::CellMeta>,
    cell: &formula_model::Cell,
) -> CellValueKind {
    if let Some(meta) = meta {
        if let Some(kind) = meta.value_kind.clone() {
            // Cells with phonetic guide metadata are tricky: the phonetic runs may be stored either
            // inline (`<c t="inlineStr"><is><rPh>…`) or inside the referenced shared-string table
            // entry (`sharedStrings.xml <si><rPh>…`).
            //
            // If the original file used a shared string, prefer preserving that representation so
            // we keep the original shared string index (and avoid collapsing duplicate visible text
            // entries that differ only in phonetic metadata).
            if cell.phonetic.is_some() && matches!(&cell.value, CellValue::String(_)) {
                if matches!(&kind, CellValueKind::SharedString { .. })
                    && value_kind_compatible(&kind, &cell.value)
                {
                    return kind;
                }
                return CellValueKind::InlineString;
            }

            // Cells with less-common or unknown `t=` attributes require the original `<v>` payload
            // to round-trip safely. If we don't have it, fall back to the inferred kind so we emit
            // a valid SpreadsheetML representation.
            if matches!(&kind, CellValueKind::Other { .. }) {
                if meta.raw_value.is_some() && matches!(&cell.value, CellValue::String(_)) {
                    return kind;
                }
            } else if value_kind_compatible(&kind, &cell.value) {
                if cell.phonetic.is_some() && matches!(&cell.value, CellValue::String(_)) {
                    // If the cell has phonetic guide metadata, we generally need an inline string
                    // so we can emit SpreadsheetML `<rPh>` runs in the `<is>` payload.
                    //
                    // However, when round-tripping an existing workbook we may already have a
                    // specific shared string index (`t="s"`, `<v>idx</v>`) whose corresponding
                    // `<si>` entry (including phonetic subtrees) we can preserve byte-for-byte. In
                    // that case, keep the shared-string representation to avoid switching between
                    // duplicate `<si>` entries that differ only in phonetic metadata.
                    if let CellValueKind::SharedString { index } = &kind {
                        let raw_matches = meta
                            .raw_value
                            .as_deref()
                            .and_then(|raw| raw.trim().parse::<u32>().ok())
                            .is_some_and(|raw| raw == *index);
                        if raw_matches {
                            return kind;
                        }
                    } else {
                        return kind;
                    }
                } else {
                    return kind;
                }
            }
        }
    }

    // If the cell has phonetic guide metadata, we prefer an inline string so we can emit
    // SpreadsheetML `<rPh>` runs in the `<is>` payload.
    //
    // Note: this override is applied *after* honoring any compatible `CellMeta` value kind so
    // that round-tripping a workbook can preserve existing shared string indices (and their
    // associated phonetic/extension subtrees) when the visible text is unchanged.
    if cell.phonetic.is_some() && matches!(&cell.value, CellValue::String(_)) {
        return CellValueKind::InlineString;
    }

    infer_value_kind(cell)
}

fn value_kind_compatible(kind: &CellValueKind, value: &CellValue) -> bool {
    match (kind, value) {
        (_, CellValue::Empty) => true,
        (CellValueKind::Number, CellValue::Number(_)) => true,
        (CellValueKind::Bool, CellValue::Boolean(_)) => true,
        (CellValueKind::Error, CellValue::Error(_)) => true,
        (
            CellValueKind::SharedString { .. },
            CellValue::String(_)
            | CellValue::RichText(_)
            | CellValue::Entity(_)
            | CellValue::Record(_),
        ) => true,
        (CellValueKind::SharedString { .. }, CellValue::Image(image))
            if image.alt_text.as_deref().is_some_and(|s| !s.is_empty()) =>
        {
            true
        }
        (
            CellValueKind::InlineString,
            CellValue::String(_) | CellValue::Entity(_) | CellValue::Record(_),
        ) => true,
        (CellValueKind::InlineString, CellValue::Image(image))
            if image.alt_text.as_deref().is_some_and(|s| !s.is_empty()) =>
        {
            true
        }
        (
            CellValueKind::Str,
            CellValue::String(_) | CellValue::Entity(_) | CellValue::Record(_),
        ) => true,
        (CellValueKind::Str, CellValue::Image(image))
            if image.alt_text.as_deref().is_some_and(|s| !s.is_empty()) =>
        {
            true
        }
        _ => false,
    }
}

#[derive(Debug, Clone)]
struct SharedFormulaGroup {
    range: Range,
    ast: formula_engine::Ast,
}

fn shared_formula_groups(
    doc: &XlsxDocument,
    sheet_id: WorksheetId,
) -> HashMap<u32, SharedFormulaGroup> {
    let mut groups = HashMap::new();

    for ((ws_id, cell_ref), meta) in &doc.meta.cell_meta {
        if *ws_id != sheet_id {
            continue;
        }
        let Some(formula_meta) = meta.formula.as_ref() else {
            continue;
        };

        let is_shared_master = formula_meta.t.as_deref() == Some("shared")
            && formula_meta.reference.is_some()
            && formula_meta.shared_index.is_some()
            && !formula_meta.file_text.is_empty();
        if !is_shared_master {
            continue;
        }

        let Some(reference) = formula_meta.reference.as_deref() else {
            continue;
        };
        let Some(shared_index) = formula_meta.shared_index else {
            continue;
        };

        let range = match Range::from_a1(reference) {
            Ok(range) => range,
            Err(_) => continue,
        };

        let master_display = crate::formula_text::strip_xlfn_prefixes(&formula_meta.file_text);
        let mut opts = ParseOptions::default();
        opts.normalize_relative_to = Some(CellAddr::new(cell_ref.row, cell_ref.col));
        let ast = match parse_formula(&master_display, opts) {
            Ok(ast) => ast,
            Err(_) => continue,
        };

        groups.insert(shared_index, SharedFormulaGroup { range, ast });
    }

    groups
}

fn shared_formula_expected(
    shared_formulas: &HashMap<u32, SharedFormulaGroup>,
    shared_index: u32,
    cell_ref: CellRef,
) -> Option<String> {
    let group = shared_formulas.get(&shared_index)?;
    if !group.range.contains(cell_ref) {
        return None;
    }

    let mut ser = SerializeOptions::default();
    ser.origin = Some(CellAddr::new(cell_ref.row, cell_ref.col));
    ser.omit_equals = true;
    group.ast.to_string(ser).ok()
}

fn formula_file_text(meta: &crate::FormulaMeta, display: Option<&str>) -> String {
    let Some(display) = display else {
        return strip_leading_equals(&meta.file_text).to_string();
    };

    let display = strip_leading_equals(display);

    // Preserve stored file text if the model's display text matches.
    if !meta.file_text.is_empty()
        && crate::formula_text::strip_xlfn_prefixes(&meta.file_text) == display
    {
        return strip_leading_equals(&meta.file_text).to_string();
    }

    crate::formula_text::add_xlfn_prefixes(display)
}

fn strip_leading_equals(s: &str) -> &str {
    let trimmed = s.trim();
    let stripped = trimmed.strip_prefix('=').unwrap_or(trimmed);
    stripped.trim()
}

fn raw_or_number(meta: Option<&crate::CellMeta>, n: f64) -> String {
    if let Some(meta) = meta {
        if let Some(raw) = &meta.raw_value {
            if raw.parse::<f64>().ok() == Some(n) {
                return raw.clone();
            }
        }
    }
    // Default formatting is fine for deterministic output; raw_value is used to preserve
    // round-trip fidelity where available.
    n.to_string()
}

fn raw_or_bool(meta: Option<&crate::CellMeta>, b: bool) -> &'static str {
    if let Some(meta) = meta {
        if let Some(raw) = meta.raw_value.as_deref() {
            if (raw == "1" && b) || (raw == "0" && !b) {
                return if b { "1" } else { "0" };
            }
        }
    }
    if b {
        "1"
    } else {
        "0"
    }
}

fn raw_or_error(meta: Option<&crate::CellMeta>, err: ErrorValue) -> String {
    if let Some(meta) = meta {
        if let Some(raw) = &meta.raw_value {
            if raw == err.as_str() {
                return raw.clone();
            }
        }
    }
    err.as_str().to_string()
}

fn raw_or_str(meta: Option<&crate::CellMeta>, s: &str) -> String {
    if let Some(meta) = meta {
        if let Some(raw) = &meta.raw_value {
            if raw == s {
                return raw.clone();
            }
        }
    }
    s.to_string()
}

fn raw_or_other(meta: Option<&crate::CellMeta>, s: &str) -> String {
    // Unknown/less-common `t=` types store their payload as text; preserve the original `<v>`
    // content when it still matches the in-memory value.
    raw_or_str(meta, s)
}

fn shared_string_index(
    doc: &XlsxDocument,
    meta: Option<&crate::CellMeta>,
    value: &CellValue,
    shared_lookup: &HashMap<SharedStringKey, u32>,
) -> u32 {
    match value {
        CellValue::String(text) => {
            if let Some(meta) = meta {
                if let Some(CellValueKind::SharedString { index }) = &meta.value_kind {
                    if doc
                        .shared_strings
                        .get(*index as usize)
                        .map(|rt| rt.text.as_str())
                        == Some(text.as_str())
                    {
                        return *index;
                    }
                }
            }
            shared_lookup
                .get(&SharedStringKey::plain(text))
                .copied()
                .unwrap_or(0)
        }
        CellValue::Image(image) => match image.alt_text.as_deref().filter(|s| !s.is_empty()) {
            Some(text) => {
                if let Some(meta) = meta {
                    if let Some(CellValueKind::SharedString { index }) = &meta.value_kind {
                        if doc
                            .shared_strings
                            .get(*index as usize)
                            .map(|rt| rt.text.as_str())
                            == Some(text)
                        {
                            return *index;
                        }
                    }
                }
                shared_lookup
                    .get(&SharedStringKey::plain(text))
                    .copied()
                    .unwrap_or(0)
            }
            None => 0,
        },
        CellValue::Entity(entity) => {
            let text = entity.display_value.as_str();
            if let Some(meta) = meta {
                if let Some(CellValueKind::SharedString { index }) = &meta.value_kind {
                    if doc
                        .shared_strings
                        .get(*index as usize)
                        .map(|rt| rt.text.as_str())
                        == Some(text)
                    {
                        return *index;
                    }
                }
            }
            shared_lookup
                .get(&SharedStringKey::plain(text))
                .copied()
                .unwrap_or(0)
        }
        CellValue::Record(record) => {
            let text_owned = record.to_string();
            let text = text_owned.as_str();
            if let Some(meta) = meta {
                if let Some(CellValueKind::SharedString { index }) = &meta.value_kind {
                    if doc
                        .shared_strings
                        .get(*index as usize)
                        .map(|rt| rt.text.as_str())
                        == Some(text)
                    {
                        return *index;
                    }
                }
            }
            shared_lookup
                .get(&SharedStringKey::plain(text))
                .copied()
                .unwrap_or(0)
        }
        CellValue::RichText(rich) => {
            if let Some(meta) = meta {
                if let Some(CellValueKind::SharedString { index }) = &meta.value_kind {
                    if doc
                        .shared_strings
                        .get(*index as usize)
                        .map(|rt| rt == rich)
                        .unwrap_or(false)
                    {
                        return *index;
                    }
                }
            }
            shared_lookup
                .get(&SharedStringKey::from_rich_text(rich))
                .copied()
                .unwrap_or(0)
        }
        _ => 0,
    }
}

fn generate_minimal_package(
    sheets: &[SheetMeta],
    workbook_kind: WorkbookKind,
) -> Result<BTreeMap<String, Vec<u8>>, WriteError> {
    let mut parts = BTreeMap::new();

    parts.insert(
        "_rels/.rels".to_string(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#
        .to_vec(),
    );

    // Minimal workbook relationships; existing packages preserve the original bytes.
    parts.insert(
        "xl/_rels/workbook.xml.rels".to_string(),
        minimal_workbook_rels_xml(sheets).into_bytes(),
    );

    parts.insert(
        "[Content_Types].xml".to_string(),
        minimal_content_types_xml(sheets, workbook_kind).into_bytes(),
    );

    Ok(parts)
}

fn minimal_workbook_rels_xml(sheets: &[SheetMeta]) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );

    for sheet_meta in sheets {
        let target = relationship_target_from_workbook(&sheet_meta.path);
        xml.push_str(r#"<Relationship Id=""#);
        xml.push_str(&escape_attr(&sheet_meta.relationship_id));
        xml.push_str(r#"" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target=""#);
        xml.push_str(&escape_attr(&target));
        xml.push_str(r#""/>"#);
    }

    let next = next_relationship_id(sheets.iter().map(|s| s.relationship_id.as_str()));
    xml.push_str(&format!(
        r#"<Relationship Id="rId{next}" Type="{REL_TYPE_STYLES}" Target="styles.xml"/>"#
    ));
    let next2 = next + 1;
    xml.push_str(&format!(
        r#"<Relationship Id="rId{next2}" Type="{REL_TYPE_SHARED_STRINGS}" Target="sharedStrings.xml"/>"#
    ));
    xml.push_str("</Relationships>");
    xml
}

fn relationship_target_from_workbook(part_name: &str) -> String {
    let base_dir = WORKBOOK_PART
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");
    relative_target(base_dir, part_name)
}

fn relative_target(base_dir: &str, part_name: &str) -> String {
    let base_parts: Vec<&str> = base_dir.split('/').filter(|p| !p.is_empty()).collect();
    let target_parts: Vec<&str> = part_name.split('/').filter(|p| !p.is_empty()).collect();

    let mut common = 0usize;
    while common < base_parts.len()
        && common < target_parts.len()
        && base_parts[common] == target_parts[common]
    {
        common += 1;
    }

    let mut out: Vec<&str> = Vec::new();
    for _ in common..base_parts.len() {
        out.push("..");
    }
    out.extend_from_slice(&target_parts[common..]);

    if out.is_empty() {
        ".".to_string()
    } else {
        out.join("/")
    }
}

fn minimal_content_types_xml(sheets: &[SheetMeta], workbook_kind: WorkbookKind) -> String {
    let mut xml = String::new();
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#);
    xml.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
    xml.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);
    xml.push_str(r#"<Override PartName="/xl/workbook.xml" ContentType=""#);
    xml.push_str(workbook_kind.workbook_content_type());
    xml.push_str(r#""/>"#);
    for sheet_meta in sheets {
        xml.push_str(r#"<Override PartName="/"#);
        xml.push_str(&escape_attr(&sheet_meta.path));
        xml.push_str(r#"" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#);
    }
    xml.push_str(r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#);
    xml.push_str(r#"<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"#);
    xml.push_str("</Types>");
    xml
}

fn ensure_content_types_override(
    parts: &mut BTreeMap<String, Vec<u8>>,
    part_name: &str,
    content_type: &str,
) -> Result<(), WriteError> {
    let Some(existing) = parts.get("[Content_Types].xml").cloned() else {
        // Avoid synthesizing a full file for existing packages.
        return Ok(());
    };

    // Fast path: if the part already exists, avoid rewriting the file to keep roundtrip
    // output stable.
    let xml = std::str::from_utf8(&existing)
        .map_err(|e| WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))?;
    if xml.contains(&format!(r#"PartName="{part_name}""#)) {
        return Ok(());
    }

    let mut reader = Reader::from_reader(existing.as_slice());
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(existing.len() + 128));
    let mut buf = Vec::new();

    let mut saw_part = false;
    let mut inserted = false;

    let mut types_prefix: Option<String> = None;
    let mut has_default_ns = false;
    let mut override_prefix: Option<String> = None;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"Types" => {
                if types_prefix.is_none() {
                    types_prefix = element_prefix(e.name().as_ref())
                        .map(|p| String::from_utf8_lossy(p).into_owned());
                }
                if !has_default_ns {
                    has_default_ns = content_types_has_default_ns(e)?;
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"Types" => {
                // Some producers emit an empty content-types part as `<Types .../>`.
                // Expand it so we can insert the required Override.
                if types_prefix.is_none() {
                    types_prefix = element_prefix(e.name().as_ref())
                        .map(|p| String::from_utf8_lossy(p).into_owned());
                }
                if !has_default_ns {
                    has_default_ns = content_types_has_default_ns(e)?;
                }

                if saw_part {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                } else {
                    let prefix = override_prefix.as_deref().or_else(|| {
                        if !has_default_ns {
                            types_prefix.as_deref()
                        } else {
                            None
                        }
                    });
                    let tag = prefixed_tag(prefix, "Override");
                    let mut override_el = quick_xml::events::BytesStart::new(tag.as_str());
                    override_el.push_attribute(("PartName", part_name));
                    override_el.push_attribute(("ContentType", content_type));

                    writer.write_event(Event::Start(e.to_owned()))?;
                    writer.write_event(Event::Empty(override_el))?;

                    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.get_mut().extend_from_slice(b"</");
                    writer.get_mut().extend_from_slice(tag.as_bytes());
                    writer.get_mut().extend_from_slice(b">");

                    inserted = true;
                }
            }
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"Override" => {
                if override_prefix.is_none() {
                    override_prefix = element_prefix(e.name().as_ref())
                        .map(|p| String::from_utf8_lossy(p).into_owned());
                }
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"PartName"
                        && attr.unescape_value()?.as_ref() == part_name
                    {
                        saw_part = true;
                        break;
                    }
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"Override" => {
                if override_prefix.is_none() {
                    override_prefix = element_prefix(e.name().as_ref())
                        .map(|p| String::from_utf8_lossy(p).into_owned());
                }
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"PartName"
                        && attr.unescape_value()?.as_ref() == part_name
                    {
                        saw_part = true;
                        break;
                    }
                }
                writer.write_event(Event::Empty(e.to_owned()))?;
            }
            Event::End(ref e) if local_name(e.name().as_ref()) == b"Types" => {
                if !saw_part {
                    let prefix = override_prefix.as_deref().or_else(|| {
                        if !has_default_ns {
                            types_prefix.as_deref()
                        } else {
                            None
                        }
                    });
                    let tag = prefixed_tag(prefix, "Override");
                    let mut override_el = quick_xml::events::BytesStart::new(tag.as_str());
                    override_el.push_attribute(("PartName", part_name));
                    override_el.push_attribute(("ContentType", content_type));
                    writer.write_event(Event::Empty(override_el))?;
                    inserted = true;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            Event::Eof => break,
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    if inserted {
        parts.insert("[Content_Types].xml".to_string(), writer.into_inner());
    }
    Ok(())
}

#[allow(dead_code)]
fn ensure_content_types_default(
    parts: &mut BTreeMap<String, Vec<u8>>,
    ext: &str,
    content_type: &str,
) -> Result<(), WriteError> {
    let Some(existing) = parts.get("[Content_Types].xml").cloned() else {
        // Avoid synthesizing a full file for existing packages.
        return Ok(());
    };
    let ext = ext.trim();

    let mut reader = Reader::from_reader(existing.as_slice());
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(existing.len() + 128));
    let mut buf = Vec::new();

    let mut saw_ext = false;
    let mut inserted = false;

    let mut types_prefix: Option<String> = None;
    let mut has_default_ns = false;
    let mut default_prefix: Option<String> = None;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"Types" => {
                if types_prefix.is_none() {
                    types_prefix = element_prefix(e.name().as_ref())
                        .map(|p| String::from_utf8_lossy(p).into_owned());
                }
                if !has_default_ns {
                    has_default_ns = content_types_has_default_ns(e)?;
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"Types" => {
                // Some producers emit an empty content-types part as `<Types .../>`.
                // Expand it so we can insert the required Default.
                if types_prefix.is_none() {
                    types_prefix = element_prefix(e.name().as_ref())
                        .map(|p| String::from_utf8_lossy(p).into_owned());
                }
                if !has_default_ns {
                    has_default_ns = content_types_has_default_ns(e)?;
                }

                if saw_ext {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                } else {
                    let prefix = default_prefix.as_deref().or_else(|| {
                        if !has_default_ns {
                            types_prefix.as_deref()
                        } else {
                            None
                        }
                    });
                    let tag = prefixed_tag(prefix, "Default");
                    let mut default_el = quick_xml::events::BytesStart::new(tag.as_str());
                    default_el.push_attribute(("Extension", ext));
                    default_el.push_attribute(("ContentType", content_type));

                    writer.write_event(Event::Start(e.to_owned()))?;
                    writer.write_event(Event::Empty(default_el))?;

                    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.get_mut().extend_from_slice(b"</");
                    writer.get_mut().extend_from_slice(tag.as_bytes());
                    writer.get_mut().extend_from_slice(b">");

                    inserted = true;
                }
            }
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"Default" => {
                if default_prefix.is_none() {
                    default_prefix = element_prefix(e.name().as_ref())
                        .map(|p| String::from_utf8_lossy(p).into_owned());
                }
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"Extension"
                        && attr
                            .unescape_value()?
                            .as_ref()
                            .trim()
                            .eq_ignore_ascii_case(ext)
                    {
                        saw_ext = true;
                        break;
                    }
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"Default" => {
                if default_prefix.is_none() {
                    default_prefix = element_prefix(e.name().as_ref())
                        .map(|p| String::from_utf8_lossy(p).into_owned());
                }
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"Extension"
                        && attr
                            .unescape_value()?
                            .as_ref()
                            .trim()
                            .eq_ignore_ascii_case(ext)
                    {
                        saw_ext = true;
                        break;
                    }
                }
                writer.write_event(Event::Empty(e.to_owned()))?;
            }
            Event::End(ref e) if local_name(e.name().as_ref()) == b"Types" => {
                if !saw_ext {
                    let prefix = default_prefix.as_deref().or_else(|| {
                        if !has_default_ns {
                            types_prefix.as_deref()
                        } else {
                            None
                        }
                    });
                    let tag = prefixed_tag(prefix, "Default");
                    let mut default_el = quick_xml::events::BytesStart::new(tag.as_str());
                    default_el.push_attribute(("Extension", ext));
                    default_el.push_attribute(("ContentType", content_type));
                    writer.write_event(Event::Empty(default_el))?;
                    inserted = true;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            Event::Eof => break,
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    if inserted {
        parts.insert("[Content_Types].xml".to_string(), writer.into_inner());
    }

    Ok(())
}

fn relationship_targets_by_type(rels_xml: &[u8], rel_type: &str) -> Result<Vec<String>, WriteError> {
    let mut reader = Reader::from_reader(rels_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    let mut out = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                let mut type_ = None;
                let mut target = None;
                let mut target_mode = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    let key = local_name(attr.key.as_ref());
                    if key.eq_ignore_ascii_case(b"Type") {
                        type_ = Some(attr.unescape_value()?.into_owned());
                    } else if key.eq_ignore_ascii_case(b"Target") {
                        target = Some(attr.unescape_value()?.into_owned());
                    } else if key.eq_ignore_ascii_case(b"TargetMode") {
                        target_mode = Some(attr.unescape_value()?.into_owned());
                    }
                }

                if target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                {
                    continue;
                }

                if type_.as_deref() == Some(rel_type) {
                    if let Some(target) = target {
                        out.push(target);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(out)
}

fn ensure_drawing_part_content_types(
    parts: &mut BTreeMap<String, Vec<u8>>,
    drawing_part_path: &str,
) -> Result<(), WriteError> {
    if !parts.contains_key(drawing_part_path) {
        return Ok(());
    }

    ensure_content_types_override(
        parts,
        &format!("/{drawing_part_path}"),
        DRAWING_CONTENT_TYPE,
    )?;

    let drawing_rels_path = crate::drawings::DrawingPart::rels_path_for(drawing_part_path);
    let Some(drawing_rels_bytes) = parts.get(&drawing_rels_path) else {
        return Ok(());
    };

    let image_targets =
        relationship_targets_by_type(drawing_rels_bytes, crate::drawings::REL_TYPE_IMAGE)?;
    let chart_targets = relationship_targets_by_type(drawing_rels_bytes, CHART_REL_TYPE)?;

    // Ensure image media extensions referenced from this drawing have Default content types.
    for target in image_targets {
        let media_part = resolve_target(drawing_part_path, &target);
        if !media_part.starts_with("xl/media/") {
            continue;
        }
        let Some((_, ext)) = media_part.rsplit_once('.') else {
            continue;
        };
        let ext = ext.trim().to_ascii_lowercase();
        if ext.is_empty() {
            continue;
        }
        let content_type = crate::drawings::content_type_for_extension(&ext);
        if content_type == "application/octet-stream" {
            continue;
        }
        ensure_content_types_default(parts, &ext, content_type)?;
    }

    // Ensure chart parts referenced from this drawing have Overrides.
    for target in chart_targets {
        let chart_part = resolve_target(drawing_part_path, &target);
        if !parts.contains_key(&chart_part) {
            continue;
        }
        ensure_content_types_override(
            parts,
            &format!("/{chart_part}"),
            CHART_CONTENT_TYPE,
        )?;
    }

    Ok(())
}

fn relationship_target_by_type(
    rels_xml: &[u8],
    rel_type: &str,
) -> Result<Option<String>, WriteError> {
    let mut reader = Reader::from_reader(rels_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                let mut type_ = None;
                let mut target = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    let key = local_name(attr.key.as_ref());
                    if key.eq_ignore_ascii_case(b"Type") {
                        type_ = Some(attr.unescape_value()?.into_owned());
                    } else if key.eq_ignore_ascii_case(b"Target") {
                        target = Some(attr.unescape_value()?.into_owned());
                    }
                }
                if type_.as_deref() == Some(rel_type) {
                    return Ok(target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(None)
}

fn relationship_target_by_id(rels_xml: &[u8], rel_id: &str) -> Result<Option<String>, WriteError> {
    let mut reader = Reader::from_reader(rels_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                let mut id = None;
                let mut target = None;
                let mut target_mode = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    let key = local_name(attr.key.as_ref());
                    if key.eq_ignore_ascii_case(b"Id") {
                        id = Some(attr.unescape_value()?.into_owned());
                    } else if key.eq_ignore_ascii_case(b"Target") {
                        target = Some(attr.unescape_value()?.into_owned());
                    } else if key.eq_ignore_ascii_case(b"TargetMode") {
                        target_mode = Some(attr.unescape_value()?.into_owned());
                    }
                }

                if id.as_deref() == Some(rel_id) {
                    if target_mode
                        .as_deref()
                        .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
                    {
                        return Ok(None);
                    }
                    return Ok(target);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(None)
}

fn relationships_root_is_prefix_only(rels_xml: &[u8]) -> Result<bool, WriteError> {
    let mut reader = Reader::from_reader(rels_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(ref e) | Event::Empty(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") =>
            {
                // Root is prefix-only when it has a tag prefix (e.g. `rel:Relationships`) *and*
                // does not declare the relationship namespace as the default xmlns (meaning
                // `<Relationships xmlns="...">`).
                let has_prefix = element_prefix(e.name().as_ref()).is_some();
                if !has_prefix {
                    return Ok(false);
                }

                let mut has_default_ns = false;
                for attr in e.attributes().with_checks(false) {
                    let attr = attr?;
                    if attr.key.as_ref() == b"xmlns"
                        && attr.value.as_ref() == crate::relationships::PACKAGE_REL_NS.as_bytes()
                    {
                        has_default_ns = true;
                        break;
                    }
                }
                return Ok(has_prefix && !has_default_ns);
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(false)
}
fn relationship_id_and_target_by_type(
    rels_xml: &[u8],
    rel_type: &str,
) -> Result<Option<(String, String)>, WriteError> {
    let mut reader = Reader::from_reader(rels_xml);
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Start(e) | Event::Empty(e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                let mut id = None;
                let mut type_ = None;
                let mut target = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    let key = local_name(attr.key.as_ref());
                    if key.eq_ignore_ascii_case(b"Id") {
                        id = Some(attr.unescape_value()?.into_owned());
                    } else if key.eq_ignore_ascii_case(b"Type") {
                        type_ = Some(attr.unescape_value()?.into_owned());
                    } else if key.eq_ignore_ascii_case(b"Target") {
                        target = Some(attr.unescape_value()?.into_owned());
                    }
                }
                if type_.as_deref() == Some(rel_type) {
                    if let (Some(id), Some(target)) = (id, target) {
                        return Ok(Some((id, target)));
                    }
                    return Ok(None);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    Ok(None)
}

fn ensure_workbook_rels_has_relationship(
    parts: &mut BTreeMap<String, Vec<u8>>,
    rel_type: &str,
    target: &str,
) -> Result<(), WriteError> {
    let rels_name = WORKBOOK_RELS_PART;
    let Some(existing) = parts.get(rels_name).cloned() else {
        return Ok(());
    };
    if relationship_target_by_type(&existing, rel_type)?.is_some() {
        return Ok(());
    }

    let xml = String::from_utf8(existing.clone())
        .map_err(|e| WriteError::Io(std::io::Error::new(std::io::ErrorKind::InvalidData, e)))?;
    let next = next_relationship_id_in_xml(&xml);
    let id = format!("rId{next}");

    let mut reader = Reader::from_reader(existing.as_slice());
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(existing.len() + 128));
    let mut buf = Vec::new();

    let mut root_prefix: Option<String> = None;
    let mut root_has_default_ns = false;
    let mut root_declared_prefixes: HashSet<String> = HashSet::new();
    let mut relationship_prefix: Option<String> = None;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") =>
            {
                if root_prefix.is_none() {
                    root_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                if !root_has_default_ns || root_declared_prefixes.is_empty() {
                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = attr.key.as_ref();
                        if key == b"xmlns"
                            && attr.value.as_ref()
                                == crate::relationships::PACKAGE_REL_NS.as_bytes()
                        {
                            root_has_default_ns = true;
                        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                            if attr.value.as_ref()
                                == crate::relationships::PACKAGE_REL_NS.as_bytes()
                            {
                                if let Ok(prefix) = std::str::from_utf8(prefix) {
                                    root_declared_prefixes.insert(prefix.to_string());
                                }
                            }
                        }
                    }
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") =>
            {
                // Some producers emit an empty relationships part as `<Relationships .../>`.
                // Expand it so we can insert the required Relationship.
                if root_prefix.is_none() {
                    root_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                if !root_has_default_ns {
                    for attr in e.attributes() {
                        let attr = attr?;
                        if attr.key.as_ref() == b"xmlns"
                            && attr.value.as_ref()
                                == crate::relationships::PACKAGE_REL_NS.as_bytes()
                        {
                            root_has_default_ns = true;
                            break;
                        }
                    }
                }

                let prefix = relationship_prefix.as_deref().or_else(|| {
                    if root_has_default_ns {
                        None
                    } else {
                        root_prefix.as_deref()
                    }
                });
                let relationship_tag = prefixed_tag(prefix, "Relationship");
                let mut rel = quick_xml::events::BytesStart::new(relationship_tag.as_str());
                rel.push_attribute(("Id", id.as_str()));
                rel.push_attribute(("Type", rel_type));
                rel.push_attribute(("Target", target));

                writer.write_event(Event::Start(e.to_owned()))?;
                writer.write_event(Event::Empty(rel))?;

                let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                writer.get_mut().extend_from_slice(b"</");
                writer.get_mut().extend_from_slice(tag.as_bytes());
                writer.get_mut().extend_from_slice(b">");
            }
            Event::Start(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                if relationship_prefix.is_none() {
                    relationship_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                if relationship_prefix.is_none() {
                    relationship_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                writer.write_event(Event::Empty(e.to_owned()))?;
            }
            Event::End(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") =>
            {
                let prefix = relationship_prefix
                    .as_deref()
                    // Only reuse the existing Relationship element prefix if it is declared on the
                    // root element. Otherwise, we could emit a new sibling with an out-of-scope
                    // prefix (invalid XML), e.g. if the input declares `xmlns:pr` on each
                    // `<pr:Relationship>` element instead of the root.
                    .filter(|p| root_declared_prefixes.contains(*p))
                    .or_else(|| {
                        if root_has_default_ns {
                            None
                        } else {
                            root_prefix.as_deref()
                        }
                    });
                let relationship_tag = prefixed_tag(prefix, "Relationship");
                let mut rel = quick_xml::events::BytesStart::new(relationship_tag.as_str());
                rel.push_attribute(("Id", id.as_str()));
                rel.push_attribute(("Type", rel_type));
                rel.push_attribute(("Target", target));
                writer.write_event(Event::Empty(rel))?;

                writer.write_event(Event::End(e.to_owned()))?;
            }
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    parts.insert(rels_name.to_string(), writer.into_inner());
    Ok(())
}

fn patch_workbook_rels_for_sheet_edits(
    parts: &mut BTreeMap<String, Vec<u8>>,
    removed: &[SheetMeta],
    added: &[SheetMeta],
) -> Result<(), WriteError> {
    let rels_name = "xl/_rels/workbook.xml.rels";
    let Some(existing) = parts.get(rels_name).cloned() else {
        return Ok(());
    };

    let remove_ids: HashSet<&str> = removed.iter().map(|m| m.relationship_id.as_str()).collect();

    let mut reader = Reader::from_reader(existing.as_slice());
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(existing.len() + added.len() * 128));
    let mut buf = Vec::new();

    let mut root_prefix: Option<String> = None;
    let mut root_has_default_ns = false;
    let mut root_declared_prefixes: HashSet<String> = HashSet::new();
    let mut relationship_prefix: Option<String> = None;

    let mut skipping = false;
    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") =>
            {
                if root_prefix.is_none() {
                    root_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                if !root_has_default_ns || root_declared_prefixes.is_empty() {
                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = attr.key.as_ref();
                        if key == b"xmlns"
                            && attr.value.as_ref()
                                == crate::relationships::PACKAGE_REL_NS.as_bytes()
                        {
                            root_has_default_ns = true;
                        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                            if attr.value.as_ref()
                                == crate::relationships::PACKAGE_REL_NS.as_bytes()
                            {
                                if let Ok(prefix) = std::str::from_utf8(prefix) {
                                    root_declared_prefixes.insert(prefix.to_string());
                                }
                            }
                        }
                    }
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") =>
            {
                // Some producers emit an empty relationships part as `<Relationships .../>`.
                // If we're adding new sheets, expand that root element so we can insert children.
                if root_prefix.is_none() {
                    root_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                if !root_has_default_ns || root_declared_prefixes.is_empty() {
                    for attr in e.attributes() {
                        let attr = attr?;
                        let key = attr.key.as_ref();
                        if key == b"xmlns"
                            && attr.value.as_ref()
                                == crate::relationships::PACKAGE_REL_NS.as_bytes()
                        {
                            root_has_default_ns = true;
                        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
                            if attr.value.as_ref()
                                == crate::relationships::PACKAGE_REL_NS.as_bytes()
                            {
                                if let Ok(prefix) = std::str::from_utf8(prefix) {
                                    root_declared_prefixes.insert(prefix.to_string());
                                }
                            }
                        }
                    }
                }

                if added.is_empty() {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;

                    let prefix = relationship_prefix
                        .as_deref()
                        .filter(|p| root_declared_prefixes.contains(*p))
                        .or_else(|| {
                            if root_has_default_ns {
                                None
                            } else {
                                root_prefix.as_deref()
                            }
                        });
                    let relationship_tag = prefixed_tag(prefix, "Relationship");
                    for sheet in added {
                        let target = relationship_target_from_workbook(&sheet.path);
                        let mut rel = quick_xml::events::BytesStart::new(relationship_tag.as_str());
                        rel.push_attribute(("Id", sheet.relationship_id.as_str()));
                        rel.push_attribute(("Type", WORKSHEET_REL_TYPE));
                        rel.push_attribute(("Target", target.as_str()));
                        writer.write_event(Event::Empty(rel))?;
                    }

                    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.get_mut().extend_from_slice(b"</");
                    writer.get_mut().extend_from_slice(tag.as_bytes());
                    writer.get_mut().extend_from_slice(b">");
                }
            }
            Event::Start(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                if relationship_prefix.is_none() {
                    relationship_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                let mut id = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Id") {
                        id = Some(attr.unescape_value()?.into_owned());
                    }
                }
                if id.as_deref().is_some_and(|id| remove_ids.contains(id)) {
                    skipping = true;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;
                }
            }
            Event::Empty(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                if relationship_prefix.is_none() {
                    relationship_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                let mut id = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Id") {
                        id = Some(attr.unescape_value()?.into_owned());
                    }
                }
                if !id.as_deref().is_some_and(|id| remove_ids.contains(id)) {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                }
            }
            Event::End(ref e)
                if skipping
                    && local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                skipping = false;
            }
            Event::End(ref e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationships") =>
            {
                let prefix = relationship_prefix
                    .as_deref()
                    .filter(|p| root_declared_prefixes.contains(*p))
                    .or_else(|| {
                        if root_has_default_ns {
                            None
                        } else {
                            root_prefix.as_deref()
                        }
                    });
                let relationship_tag = prefixed_tag(prefix, "Relationship");
                for sheet in added {
                    let target = relationship_target_from_workbook(&sheet.path);
                    let mut rel = quick_xml::events::BytesStart::new(relationship_tag.as_str());
                    rel.push_attribute(("Id", sheet.relationship_id.as_str()));
                    rel.push_attribute(("Type", WORKSHEET_REL_TYPE));
                    rel.push_attribute(("Target", target.as_str()));
                    writer.write_event(Event::Empty(rel))?;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            ev if skipping => drop(ev),
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    parts.insert(rels_name.to_string(), writer.into_inner());
    Ok(())
}

fn patch_content_types_for_sheet_edits(
    parts: &mut BTreeMap<String, Vec<u8>>,
    removed: &[SheetMeta],
    added: &[SheetMeta],
) -> Result<(), WriteError> {
    let ct_name = "[Content_Types].xml";
    let Some(existing) = parts.get(ct_name).cloned() else {
        return Ok(());
    };

    let removed_parts: HashSet<String> = removed
        .iter()
        .map(|m| {
            if m.path.starts_with('/') {
                m.path.clone()
            } else {
                format!("/{}", m.path)
            }
        })
        .collect();

    let mut reader = Reader::from_reader(existing.as_slice());
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(existing.len() + added.len() * 128));
    let mut buf = Vec::new();

    let mut existing_overrides: HashSet<String> = HashSet::new();
    let mut skipping = false;

    let mut types_prefix: Option<String> = None;
    let mut has_default_ns = false;
    let mut override_prefix: Option<String> = None;

    loop {
        let event = reader.read_event_into(&mut buf)?;
        match event {
            Event::Eof => break,
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"Types" => {
                if types_prefix.is_none() {
                    types_prefix = element_prefix(e.name().as_ref())
                        .map(|p| String::from_utf8_lossy(p).into_owned());
                }
                if !has_default_ns {
                    has_default_ns = content_types_has_default_ns(e)?;
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"Types" => {
                // Some producers emit an empty content-types part as `<Types .../>`.
                // If we're adding new sheets, expand that root element so we can insert children.
                if types_prefix.is_none() {
                    types_prefix = element_prefix(e.name().as_ref())
                        .map(|p| String::from_utf8_lossy(p).into_owned());
                }
                if !has_default_ns {
                    has_default_ns = content_types_has_default_ns(e)?;
                }

                if added.is_empty() {
                    writer.write_event(Event::Empty(e.to_owned()))?;
                } else {
                    writer.write_event(Event::Start(e.to_owned()))?;

                    let prefix = override_prefix.as_deref().or_else(|| {
                        if !has_default_ns {
                            types_prefix.as_deref()
                        } else {
                            None
                        }
                    });
                    for sheet in added {
                        let part_name = if sheet.path.starts_with('/') {
                            sheet.path.clone()
                        } else {
                            format!("/{}", sheet.path)
                        };
                        if existing_overrides.contains(&part_name) {
                            continue;
                        }
                        let tag = prefixed_tag(prefix, "Override");
                        let mut override_el = quick_xml::events::BytesStart::new(tag.as_str());
                        override_el.push_attribute(("PartName", part_name.as_str()));
                        override_el.push_attribute(("ContentType", WORKSHEET_CONTENT_TYPE));
                        writer.write_event(Event::Empty(override_el))?;
                    }

                    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    writer.get_mut().extend_from_slice(b"</");
                    writer.get_mut().extend_from_slice(tag.as_bytes());
                    writer.get_mut().extend_from_slice(b">");
                }
            }
            Event::Start(ref e) if local_name(e.name().as_ref()) == b"Override" => {
                if override_prefix.is_none() {
                    override_prefix = element_prefix(e.name().as_ref())
                        .map(|p| String::from_utf8_lossy(p).into_owned());
                }
                let mut part_name = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"PartName" {
                        part_name = Some(attr.unescape_value()?.into_owned());
                    }
                }
                if let Some(name) = &part_name {
                    existing_overrides.insert(name.clone());
                    if removed_parts.contains(name) {
                        skipping = true;
                        continue;
                    }
                }
                writer.write_event(Event::Start(e.to_owned()))?;
            }
            Event::Empty(ref e) if local_name(e.name().as_ref()) == b"Override" => {
                if override_prefix.is_none() {
                    override_prefix = element_prefix(e.name().as_ref())
                        .map(|p| String::from_utf8_lossy(p).into_owned());
                }
                let mut part_name = None;
                for attr in e.attributes() {
                    let attr = attr?;
                    if attr.key.as_ref() == b"PartName" {
                        part_name = Some(attr.unescape_value()?.into_owned());
                    }
                }
                if let Some(name) = &part_name {
                    existing_overrides.insert(name.clone());
                    if removed_parts.contains(name) {
                        continue;
                    }
                }
                writer.write_event(Event::Empty(e.to_owned()))?;
            }
            Event::End(ref e) if skipping && local_name(e.name().as_ref()) == b"Override" => {
                skipping = false;
            }
            Event::End(ref e) if local_name(e.name().as_ref()) == b"Types" => {
                let prefix = override_prefix.as_deref().or_else(|| {
                    if !has_default_ns {
                        types_prefix.as_deref()
                    } else {
                        None
                    }
                });
                for sheet in added {
                    let part_name = if sheet.path.starts_with('/') {
                        sheet.path.clone()
                    } else {
                        format!("/{}", sheet.path)
                    };
                    if existing_overrides.contains(&part_name) {
                        continue;
                    }
                    let tag = prefixed_tag(prefix, "Override");
                    let mut override_el = quick_xml::events::BytesStart::new(tag.as_str());
                    override_el.push_attribute(("PartName", part_name.as_str()));
                    override_el.push_attribute(("ContentType", WORKSHEET_CONTENT_TYPE));
                    writer.write_event(Event::Empty(override_el))?;
                }
                writer.write_event(Event::End(e.to_owned()))?;
            }
            ev if skipping => drop(ev),
            ev => writer.write_event(ev.into_owned())?,
        }
        buf.clear();
    }

    parts.insert(ct_name.to_string(), writer.into_inner());
    Ok(())
}

fn next_relationship_id<'a>(ids: impl Iterator<Item = &'a str>) -> u32 {
    let mut max_id = 0u32;
    for id in ids {
        if let Some(rest) = id.strip_prefix("rId") {
            if let Ok(n) = rest.parse::<u32>() {
                max_id = max_id.max(n);
            }
        }
    }
    max_id + 1
}

fn next_relationship_id_in_xml(xml: &str) -> u32 {
    // Fast-ish path: parse the XML and extract `Relationship/@Id` values.
    // This is more robust than substring search because valid OPC producers may
    // add whitespace around `=` or use different attribute ordering.
    let mut max_id = 0u32;

    let mut reader = Reader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) | Ok(Event::Empty(e))
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                for attr in e.attributes().with_checks(false) {
                    let attr = match attr {
                        Ok(attr) => attr,
                        Err(_) => continue,
                    };
                    if !local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Id") {
                        continue;
                    }
                    let value = match attr.unescape_value() {
                        Ok(v) => v.into_owned(),
                        Err(_) => continue,
                    };

                    let value_bytes = value.as_bytes();
                    if value_bytes.len() < 3 || !value_bytes[..3].eq_ignore_ascii_case(b"rId") {
                        continue;
                    }

                    let mut n = 0u32;
                    let mut saw_digit = false;
                    for &b in &value_bytes[3..] {
                        if b.is_ascii_digit() {
                            saw_digit = true;
                            n = n.saturating_mul(10).saturating_add((b - b'0') as u32);
                        } else {
                            break;
                        }
                    }
                    if saw_digit {
                        max_id = max_id.max(n);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(_) => {
                // Fallback: if parsing fails (malformed XML), fall back to a simple substring scan
                // so we can still make best-effort progress.
                let mut rest = xml;
                while let Some(idx) = rest.find("Id=\"rId") {
                    let after = &rest[idx + "Id=\"rId".len()..];
                    let mut digits = String::new();
                    for ch in after.chars() {
                        if ch.is_ascii_digit() {
                            digits.push(ch);
                        } else {
                            break;
                        }
                    }
                    if let Ok(n) = digits.parse::<u32>() {
                        max_id = max_id.max(n);
                    }
                    rest = &after[digits.len()..];
                }
                break;
            }
        }

        buf.clear();
    }

    max_id + 1
}

#[cfg(test)]
mod tests;
