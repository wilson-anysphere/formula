//! Part-preserving cell edit model + patch application.
//!
//! This module provides a small edit DSL (`WorkbookCellPatches`) that can be
//! applied to an existing [`crate::XlsxPackage`] without regenerating the whole
//! workbook. The implementation focuses on preserving every unrelated part
//! (charts, pivots, customXml, VBA, etc.) while rewriting only the affected
//! worksheet XML parts (plus `sharedStrings.xml` / `workbook.xml` when needed).

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};

use formula_model::rich_text::{RichText, RichTextRun, RichTextRunStyle, Underline};
use formula_model::{CellRef, CellValue, ColProperties, ErrorValue, StyleTable};
use formula_model::Color;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer};

use crate::openxml::{parse_relationships, rels_part_name, resolve_relationship_target};
use crate::path::resolve_target;
use crate::recalc_policy::apply_recalc_policy_to_parts;
use crate::shared_strings::preserve::SharedStringsEditor;
use crate::styles::XlsxStylesEditor;
use crate::{RecalcPolicy, WorkbookSheetInfo, XlsxError, XlsxPackage};

const WORKBOOK_PART: &str = "xl/workbook.xml";
const REL_TYPE_SHARED_STRINGS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

const REL_TYPE_STYLES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

const SPREADSHEETML_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// An owned set of cell edits to apply to an existing workbook package.
///
/// Patches are keyed by a worksheet selector, then by cell address.
///
/// Supported worksheet selectors:
/// - Worksheet (tab) name (case-insensitive, as in Excel)
/// - Worksheet part name (any key containing `/`, e.g. `xl/worksheets/sheet2.xml`)
/// - Workbook relationship id (e.g. `rId2`) when no sheet name matches
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
    /// Optional patch for the worksheet `<cols>` section (column metadata).
    ///
    /// The payload is a sparse map of 0-based column indices to overrides, matching
    /// [`formula_model::Worksheet::col_properties`]. When applied, only the `width` and `hidden`
    /// attributes are updated to match the provided map; any other existing `<col>` attributes
    /// (e.g. `outlineLevel`, `collapsed`, `style`) are preserved when present.
    ///
    /// - `None`: preserve the existing `<cols>` section.
    /// - `Some(map)`: update the existing `<cols>` section so `width`/`hidden` match `map`, and
    ///   remove `<cols>` only if it becomes empty after applying these updates.
    col_properties: Option<BTreeMap<u32, ColProperties>>,
}

impl WorksheetCellPatches {
    /// Returns `true` if there are no pending edits.
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty() && self.col_properties.is_none()
    }

    /// Insert/replace a patch for a single cell.
    pub fn set_cell(&mut self, cell: CellRef, patch: CellPatch) {
        self.cells.insert((cell.row, cell.col), patch);
    }

    /// Patch the worksheet `<cols>` section using the provided `col_properties` map.
    ///
    /// Column indices are 0-based (matching `formula_model`); `width` values are expressed in
    /// Excel "character" units (OOXML `col/@width`).
    pub fn set_col_properties(&mut self, col_properties: BTreeMap<u32, ColProperties>) {
        self.col_properties = Some(col_properties);
    }

    /// Clear all `width`/`hidden` column overrides.
    ///
    /// This removes the `width`/`customWidth` and `hidden` attributes from any existing `<col>`
    /// elements. Other `<col>` attributes (e.g. outline metadata) are preserved.
    pub fn clear_col_properties(&mut self) {
        self.col_properties = Some(BTreeMap::new());
    }

    pub(crate) fn col_properties(&self) -> Option<&BTreeMap<u32, ColProperties>> {
        self.col_properties.as_ref()
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

/// A cell style reference used by patch APIs.
///
/// Excel stores cell formatting as `xf` indices (`c/@s`) referencing `<cellXfs>` in `styles.xml`.
/// `formula_model` cells instead refer to a `style_id` in a [`StyleTable`].
///
/// Both representations use `0` as the default style; patchers treat `0` as a signal to **remove**
/// the `s` attribute (equivalent to setting it to `0`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellStyleRef {
    /// A SpreadsheetML `xf` index (`c/@s`).
    XfIndex(u32),
    /// A `formula_model` `style_id` (resolved to an `xf` index via [`XlsxStylesEditor`]).
    StyleId(u32),
}

/// A single cell edit.
#[derive(Debug, Clone, PartialEq)]
pub enum CellPatch {
    /// Clear cell contents (formula + value). Formatting is preserved unless
    /// `style` overrides it.
    Clear {
        /// Optional style override.
        style: Option<CellStyleRef>,
        /// Optional cell `vm` attribute override.
        ///
        /// SpreadsheetML uses `c/@vm` for RichData-backed cell content (e.g. images-in-cell).
        ///
        /// - `None`: preserve the existing attribute when patching an existing cell (and omit it
        ///   when inserting a new cell).
        /// - `Some(Some(n))`: set/overwrite `vm="n"`.
        /// - `Some(None)`: remove the attribute.
        vm: Option<Option<u32>>,
        /// Optional cell `cm` attribute override.
        ///
        /// Some RichData-backed cell content also requires `c/@cm`.
        ///
        /// - `None`: preserve the existing attribute when patching an existing cell (and omit it
        ///   when inserting a new cell).
        /// - `Some(Some(n))`: set/overwrite `cm="n"`.
        /// - `Some(None)`: remove the attribute.
        cm: Option<Option<u32>>,
    },
    /// Set a cell value (and optionally a formula).
    Set {
        value: CellValue,
        /// If provided, writes an `<f>` element (leading `=` is accepted).
        formula: Option<String>,
        /// Optional style override.
        style: Option<CellStyleRef>,
        /// Optional cell `vm` attribute override.
        ///
        /// See [`CellPatch::Clear::vm`] for semantics.
        vm: Option<Option<u32>>,
        /// Optional cell `cm` attribute override.
        ///
        /// See [`CellPatch::Clear::cm`] for semantics.
        cm: Option<Option<u32>>,
    },
}

impl CellPatch {
    pub fn clear() -> Self {
        Self::Clear {
            style: None,
            vm: None,
            cm: None,
        }
    }

    pub fn clear_with_style(style_index: u32) -> Self {
        Self::Clear {
            style: Some(CellStyleRef::XfIndex(style_index)),
            vm: None,
            cm: None,
        }
    }

    pub fn clear_with_style_id(style_id: u32) -> Self {
        Self::Clear {
            style: Some(CellStyleRef::StyleId(style_id)),
            vm: None,
            cm: None,
        }
    }

    pub fn set_value(value: CellValue) -> Self {
        Self::Set {
            value,
            formula: None,
            style: None,
            vm: None,
            cm: None,
        }
    }

    pub fn set_value_with_formula(value: CellValue, formula: impl Into<String>) -> Self {
        Self::Set {
            value,
            formula: Some(formula.into()),
            style: None,
            vm: None,
            cm: None,
        }
    }

    pub fn set_value_with_style(value: CellValue, style_index: u32) -> Self {
        Self::Set {
            value,
            formula: None,
            style: Some(CellStyleRef::XfIndex(style_index)),
            vm: None,
            cm: None,
        }
    }

    pub fn set_value_with_style_id(value: CellValue, style_id: u32) -> Self {
        Self::Set {
            value,
            formula: None,
            style: Some(CellStyleRef::StyleId(style_id)),
            vm: None,
            cm: None,
        }
    }

    pub fn set_value_with_formula_and_style(
        value: CellValue,
        formula: impl Into<String>,
        style_index: u32,
    ) -> Self {
        Self::Set {
            value,
            formula: Some(formula.into()),
            style: Some(CellStyleRef::XfIndex(style_index)),
            vm: None,
            cm: None,
        }
    }

    pub fn set_value_with_formula_and_style_id(
        value: CellValue,
        formula: impl Into<String>,
        style_id: u32,
    ) -> Self {
        Self::Set {
            value,
            formula: Some(formula.into()),
            style: Some(CellStyleRef::StyleId(style_id)),
            vm: None,
            cm: None,
        }
    }

    pub fn with_style_ref(self, style: CellStyleRef) -> Self {
        match self {
            CellPatch::Clear { vm, cm, .. } => CellPatch::Clear {
                style: Some(style),
                vm,
                cm,
            },
            CellPatch::Set {
                value,
                formula,
                vm,
                cm,
                ..
            } => CellPatch::Set {
                value,
                formula,
                style: Some(style),
                vm,
                cm,
            },
        }
    }

    pub fn set_value_with_vm(value: CellValue, vm: u32) -> Self {
        Self::set_value(value).with_vm(vm)
    }

    pub fn with_vm(self, vm: u32) -> Self {
        self.with_vm_override(Some(Some(vm)))
    }

    pub fn clear_vm(self) -> Self {
        self.with_vm_override(Some(None))
    }

    pub fn with_vm_override(self, vm: Option<Option<u32>>) -> Self {
        match self {
            CellPatch::Clear { style, cm, .. } => CellPatch::Clear { style, vm, cm },
            CellPatch::Set {
                value,
                formula,
                style,
                cm,
                ..
            } => CellPatch::Set {
                value,
                formula,
                style,
                vm,
                cm,
            },
        }
    }

    pub fn vm_override(&self) -> Option<Option<u32>> {
        match self {
            CellPatch::Clear { vm, .. } | CellPatch::Set { vm, .. } => *vm,
        }
    }

    pub fn with_cm(self, cm: u32) -> Self {
        self.with_cm_override(Some(Some(cm)))
    }

    pub fn clear_cm(self) -> Self {
        self.with_cm_override(Some(None))
    }

    pub fn with_cm_override(self, cm: Option<Option<u32>>) -> Self {
        match self {
            CellPatch::Clear { style, vm, .. } => CellPatch::Clear { style, vm, cm },
            CellPatch::Set {
                value,
                formula,
                style,
                vm,
                ..
            } => CellPatch::Set {
                value,
                formula,
                style,
                vm,
                cm,
            },
        }
    }

    pub fn cm_override(&self) -> Option<Option<u32>> {
        match self {
            CellPatch::Clear { cm, .. } | CellPatch::Set { cm, .. } => *cm,
        }
    }

    pub fn with_style_id(self, style_id: u32) -> Self {
        self.with_style_ref(CellStyleRef::StyleId(style_id))
    }

    pub fn with_style_index(self, style_index: u32) -> Self {
        self.with_style_ref(CellStyleRef::XfIndex(style_index))
    }

    pub fn style_ref(&self) -> Option<CellStyleRef> {
        match self {
            CellPatch::Clear { style, .. } | CellPatch::Set { style, .. } => *style,
        }
    }

    pub fn style_id(&self) -> Option<u32> {
        match self.style_ref()? {
            CellStyleRef::StyleId(style_id) => Some(style_id),
            _ => None,
        }
    }

    pub fn style_index(&self) -> Option<u32> {
        match self.style_ref()? {
            CellStyleRef::XfIndex(xf_index) => Some(xf_index),
            // Clearing by style_id doesn't require an `xf` mapping.
            CellStyleRef::StyleId(0) => Some(0),
            CellStyleRef::StyleId(_) => None,
        }
    }

    fn style_index_override(
        &self,
        style_id_to_xf: Option<&HashMap<u32, u32>>,
    ) -> Result<Option<u32>, XlsxError> {
        let Some(style) = self.style_ref() else {
            return Ok(None);
        };

        match style {
            CellStyleRef::XfIndex(xf_index) => Ok(Some(xf_index)),
            CellStyleRef::StyleId(0) => Ok(Some(0)),
            CellStyleRef::StyleId(style_id) => {
                let style_id_to_xf = style_id_to_xf.ok_or_else(|| {
                    XlsxError::Invalid(
                        "style_id patches require apply_cell_patches_with_styles".to_string(),
                    )
                })?;
                let xf_index = style_id_to_xf.get(&style_id).copied().ok_or_else(|| {
                    XlsxError::Invalid(format!("unknown style_id {style_id} (missing xf mapping)"))
                })?;
                Ok(Some(xf_index))
            }
        }
    }
}

#[derive(Debug)]
struct SharedStringsState {
    editor: SharedStringsEditor,
    // Best-effort shared-string reference count delta from cell patches.
    count_delta: i32,
}

impl SharedStringsState {
    fn from_part(bytes: &[u8]) -> Result<Self, XlsxError> {
        let editor = SharedStringsEditor::parse(bytes)
            .map_err(|e| XlsxError::Invalid(format!("sharedStrings.xml parse error: {e}")))?;
        Ok(Self {
            editor,
            count_delta: 0,
        })
    }

    fn get_or_insert_plain(&mut self, text: &str) -> u32 {
        self.editor.get_or_insert_plain(text)
    }

    fn get_or_insert_rich(&mut self, rich: &RichText) -> u32 {
        self.editor.get_or_insert_rich(rich)
    }

    fn rich_at(&self, idx: u32) -> Option<&RichText> {
        self.editor.rich_at(idx)
    }

    fn note_shared_string_ref_delta(&mut self, old_uses_shared: bool, new_uses_shared: bool) {
        match (old_uses_shared, new_uses_shared) {
            (true, false) => self.count_delta -= 1,
            (false, true) => self.count_delta += 1,
            _ => {}
        }
    }

    fn write_if_dirty(&self) -> Result<Option<Vec<u8>>, XlsxError> {
        if !self.editor.is_dirty() {
            return Ok(None);
        }

        let count_hint = self
            .editor
            .original_count()
            .map(|base| base.saturating_add_signed(self.count_delta));
        let updated = self
            .editor
            .to_xml_bytes(count_hint)
            .map_err(|e| XlsxError::Invalid(format!("sharedStrings.xml write error: {e}")))?;
        Ok(Some(updated))
    }
}

pub(crate) fn apply_cell_patches_to_package(
    pkg: &mut XlsxPackage,
    patches: &WorkbookCellPatches,
    recalc_policy: RecalcPolicy,
) -> Result<(), XlsxError> {
    if patches.is_empty() {
        return Ok(());
    }

    let style_ids = collect_style_id_overrides(patches);
    if !style_ids.is_empty() {
        return Err(XlsxError::Invalid(
            "style_id patches require apply_cell_patches_with_styles".to_string(),
        ));
    }

    apply_cell_patches_to_package_inner(pkg, patches, None, recalc_policy)
}

pub(crate) fn apply_cell_patches_to_package_with_styles(
    pkg: &mut XlsxPackage,
    patches: &WorkbookCellPatches,
    style_table: &StyleTable,
    recalc_policy: RecalcPolicy,
) -> Result<(), XlsxError> {
    if patches.is_empty() {
        return Ok(());
    }

    let style_ids = collect_style_id_overrides(patches);
    if style_ids.is_empty() {
        return apply_cell_patches_to_package(pkg, patches, recalc_policy);
    }

    let styles_part_name = resolve_styles_part(pkg)?;
    let styles_bytes = pkg
        .part(&styles_part_name)
        .ok_or_else(|| XlsxError::MissingPart(styles_part_name.clone()))?;

    let mut style_table = style_table.clone();
    let mut styles_editor =
        XlsxStylesEditor::parse_or_default(Some(styles_bytes), &mut style_table)
            .map_err(|e| XlsxError::Invalid(format!("styles.xml error: {e}")))?;

    let before_xfs = styles_editor.styles_part().cell_xfs_count();
    let style_id_to_xf = styles_editor
        .ensure_styles_for_style_ids(style_ids, &style_table)
        .map_err(|e| XlsxError::Invalid(format!("styles.xml error: {e}")))?;
    let after_xfs = styles_editor.styles_part().cell_xfs_count();

    // Avoid rewriting styles.xml unless we actually appended new xfs; preserving the original
    // bytes keeps unrelated diffs smaller for high-fidelity edit workflows.
    if before_xfs != after_xfs {
        pkg.set_part(styles_part_name, styles_editor.to_styles_xml_bytes());
    }

    apply_cell_patches_to_package_inner(pkg, patches, Some(&style_id_to_xf), recalc_policy)
}

fn apply_cell_patches_to_package_inner(
    pkg: &mut XlsxPackage,
    patches: &WorkbookCellPatches,
    style_id_to_xf: Option<&HashMap<u32, u32>>,
    recalc_policy: RecalcPolicy,
) -> Result<(), XlsxError> {
    let workbook_sheets = pkg.workbook_sheets()?;

    let shared_strings_part_name = resolve_shared_strings_part_name(pkg)?.or_else(|| {
        pkg.part("xl/sharedStrings.xml")
            .map(|_| "xl/sharedStrings.xml".to_string())
    });
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

        let worksheet_part =
            resolve_worksheet_part_for_selector(pkg, &workbook_sheets, sheet_name)?;
        let original = pkg
            .part(&worksheet_part)
            .ok_or_else(|| XlsxError::MissingPart(worksheet_part.clone()))?;

        let (updated, formula_changed) = patch_worksheet_xml(
            original,
            sheet_patches,
            shared_strings.as_mut(),
            style_id_to_xf,
            recalc_policy.clear_cached_values_on_formula_change,
        )?;
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
        apply_recalc_policy_to_parts(pkg.parts_map_mut(), recalc_policy)?;
    }

    Ok(())
}

fn resolve_worksheet_part_for_selector(
    pkg: &XlsxPackage,
    workbook_sheets: &[WorkbookSheetInfo],
    selector: &str,
) -> Result<String, XlsxError> {
    let selector = selector.strip_prefix('/').unwrap_or(selector);

    // Worksheet part selector: any string containing `/` cannot be an Excel sheet name, and is
    // treated as an explicit part name.
    if selector.contains('/') {
        return Ok(selector.to_string());
    }

    // Sheet name selector (case-insensitive, matching Excel).
    if let Some(sheet) = workbook_sheets
        .iter()
        .find(|s| formula_model::sheet_name_eq_case_insensitive(&s.name, selector))
    {
        return resolve_worksheet_part(pkg, sheet);
    }

    // RelId selector: if no sheet name matches, treat the key as a workbook relationship Id.
    if let Some(part) = resolve_relationship_target(pkg, WORKBOOK_PART, selector)? {
        return Ok(part);
    }

    let rels_name = rels_part_name(WORKBOOK_PART);
    let rels_hint = if pkg.part(&rels_name).is_some() {
        rels_name
    } else {
        format!("{rels_name} (missing)")
    };

    Err(XlsxError::Invalid(format!(
        "unknown sheet selector: {selector} (tried sheet name match against {WORKBOOK_PART} and relId lookup in {rels_hint}; worksheet part selectors contain '/' e.g. xl/worksheets/sheet2.xml)"
    )))
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
        .find(|rel| {
            rel.type_uri == REL_TYPE_SHARED_STRINGS
                && !rel
                    .target_mode
                    .as_deref()
                    .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
        })
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

fn resolve_styles_part(pkg: &XlsxPackage) -> Result<String, XlsxError> {
    let rels_name = rels_part_name(WORKBOOK_PART);
    let rels_bytes = match pkg.part(&rels_name) {
        Some(bytes) => bytes,
        None => {
            // Fallback: common path when rels are missing but the part exists (best-effort).
            if pkg.part("xl/styles.xml").is_some() {
                return Ok("xl/styles.xml".to_string());
            }
            return Err(XlsxError::Invalid(
                "workbook.xml.rels missing styles relationship".to_string(),
            ));
        }
    };
    let rels = parse_relationships(rels_bytes)?;

    if let Some(rel) = rels.iter().find(|rel| {
        rel.type_uri == REL_TYPE_STYLES
            && !rel
                .target_mode
                .as_deref()
                .is_some_and(|mode| mode.trim().eq_ignore_ascii_case("External"))
    }) {
        return Ok(resolve_target(WORKBOOK_PART, &rel.target));
    }

    // Fallback: common path when rels are missing but the part exists (best-effort).
    if pkg.part("xl/styles.xml").is_some() {
        return Ok("xl/styles.xml".to_string());
    }

    Err(XlsxError::Invalid(
        "workbook.xml.rels missing styles relationship".to_string(),
    ))
}

fn collect_style_id_overrides(patches: &WorkbookCellPatches) -> Vec<u32> {
    let mut out = Vec::new();
    for (_, sheet_patches) in patches.sheets() {
        for (_, patch) in sheet_patches.iter() {
            if let Some(style_id) = patch.style_id().filter(|id| *id != 0) {
                out.push(style_id);
            }
        }
    }
    out
}

#[derive(Debug, Default, Clone, Copy)]
struct WorksheetXmlScan {
    has_dimension: bool,
    has_sheet_pr: bool,
    sheet_uses_row_spans: bool,
    /// Best-effort used-range bounds derived from `c/@r` inside `<sheetData>` and any merged
    /// ranges declared in `<mergeCells>`.
    existing_used_range: Option<(u32, u32, u32, u32)>,
}

fn scan_worksheet_xml(
    original: &[u8],
    target_cells: Option<&HashSet<(u32, u32)>>,
) -> Result<(WorksheetXmlScan, HashSet<(u32, u32)>), XlsxError> {
    let mut reader = Reader::from_reader(original);
    reader.config_mut().trim_text(true);

    let mut scan = WorksheetXmlScan::default();
    let mut in_sheet_data = false;
    let mut buf = Vec::new();
    let mut found_target_cells: HashSet<(u32, u32)> =
        HashSet::with_capacity(target_cells.map_or(0, HashSet::len));

    let mut min_row = u32::MAX;
    let mut min_col = u32::MAX;
    let mut max_row = 0u32;
    let mut max_col = 0u32;

    loop {
        match reader.read_event_into(&mut buf)? {
            Event::Eof => break,
            Event::Start(e) => match local_name(e.name().as_ref()) {
                b"dimension" => scan.has_dimension = true,
                b"sheetPr" => scan.has_sheet_pr = true,
                b"sheetData" => in_sheet_data = true,
                b"mergeCell" => {
                    for attr in e.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()) != b"ref" {
                            continue;
                        }
                        let a1 = attr.unescape_value()?.into_owned();
                        if let Some((r1, c1, r2, c2)) = parse_dimension_ref(&a1) {
                            min_row = min_row.min(r1.min(r2));
                            min_col = min_col.min(c1.min(c2));
                            max_row = max_row.max(r1.max(r2));
                            max_col = max_col.max(c1.max(c2));
                        }
                        break;
                    }
                }
                b"row" if in_sheet_data => {
                    for attr in e.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()) == b"spans" {
                            scan.sheet_uses_row_spans = true;
                            break;
                        }
                    }
                }
                b"c" if in_sheet_data => {
                    for attr in e.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()) != b"r" {
                            continue;
                        }
                        let a1 = attr.unescape_value()?.into_owned();
                        if let Ok(cell_ref) = CellRef::from_a1(&a1) {
                            let row_1 = cell_ref.row + 1;
                            let col_1 = cell_ref.col + 1;
                            min_row = min_row.min(row_1);
                            min_col = min_col.min(col_1);
                            max_row = max_row.max(row_1);
                            max_col = max_col.max(col_1);
                            if target_cells.is_some_and(|targets| {
                                targets.contains(&(cell_ref.row, cell_ref.col))
                            }) {
                                found_target_cells.insert((cell_ref.row, cell_ref.col));
                            }
                        }
                        break;
                    }
                }
                _ => {}
            },
            Event::Empty(e) => match local_name(e.name().as_ref()) {
                b"dimension" => scan.has_dimension = true,
                b"sheetPr" => scan.has_sheet_pr = true,
                b"mergeCell" => {
                    for attr in e.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()) != b"ref" {
                            continue;
                        }
                        let a1 = attr.unescape_value()?.into_owned();
                        if let Some((r1, c1, r2, c2)) = parse_dimension_ref(&a1) {
                            min_row = min_row.min(r1.min(r2));
                            min_col = min_col.min(c1.min(c2));
                            max_row = max_row.max(r1.max(r2));
                            max_col = max_col.max(c1.max(c2));
                        }
                        break;
                    }
                }
                // `<sheetData/>` has no children.
                b"row" if in_sheet_data => {
                    for attr in e.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()) == b"spans" {
                            scan.sheet_uses_row_spans = true;
                            break;
                        }
                    }
                }
                b"c" if in_sheet_data => {
                    for attr in e.attributes() {
                        let attr = attr?;
                        if local_name(attr.key.as_ref()) != b"r" {
                            continue;
                        }
                        let a1 = attr.unescape_value()?.into_owned();
                        if let Ok(cell_ref) = CellRef::from_a1(&a1) {
                            let row_1 = cell_ref.row + 1;
                            let col_1 = cell_ref.col + 1;
                            min_row = min_row.min(row_1);
                            min_col = min_col.min(col_1);
                            max_row = max_row.max(row_1);
                            max_col = max_col.max(col_1);
                            if target_cells.is_some_and(|targets| {
                                targets.contains(&(cell_ref.row, cell_ref.col))
                            }) {
                                found_target_cells.insert((cell_ref.row, cell_ref.col));
                            }
                        }
                        break;
                    }
                }
                _ => {}
            },
            Event::End(e) if local_name(e.name().as_ref()) == b"sheetData" => {
                in_sheet_data = false;
            }
            _ => {}
        }
        buf.clear();
    }

    if min_row != u32::MAX {
        scan.existing_used_range = Some((min_row, min_col, max_row, max_col));
    }

    Ok((scan, found_target_cells))
}

fn patch_worksheet_xml(
    original: &[u8],
    patches: &WorksheetCellPatches,
    mut shared_strings: Option<&mut SharedStringsState>,
    style_id_to_xf: Option<&HashMap<u32, u32>>,
    clear_cached_values_on_formula_change: bool,
) -> Result<(Vec<u8>, bool), XlsxError> {
    // Column metadata patches do not require scanning sheetData (we only rewrite the `<cols>`
    // section), so avoid the potentially-large worksheet scan when there are no cell patches.
    let mut scan = WorksheetXmlScan::default();
    let mut existing_non_material_cells: HashSet<(u32, u32)> = HashSet::new();
    if !patches.cells.is_empty() {
        let mut non_material_targets: HashSet<(u32, u32)> = HashSet::new();
        for (cell_ref, patch) in patches.iter() {
            if !cell_patch_is_material_for_insertion(patch, style_id_to_xf)? {
                non_material_targets.insert((cell_ref.row, cell_ref.col));
            }
        }

        let (scan0, existing0) = scan_worksheet_xml(
            original,
            (!non_material_targets.is_empty()).then_some(&non_material_targets),
        )?;
        scan = scan0;
        existing_non_material_cells = existing0;
    }

    // Drop patches that are guaranteed to be a no-op:
    // a non-material patch (clear with no value/formula/style) targeting a missing cell cannot
    // reference an existing `<c>` element, and since it would not insert a new cell, it cannot
    // change the worksheet XML.
    let mut effective_patches = WorksheetCellPatches::default();
    effective_patches.col_properties = patches.col_properties.clone();
    if !patches.cells.is_empty() {
        for (cell_ref, patch) in patches.iter() {
            if cell_patch_is_material_for_insertion(patch, style_id_to_xf)?
                || existing_non_material_cells.contains(&(cell_ref.row, cell_ref.col))
            {
                effective_patches
                    .cells
                    .insert((cell_ref.row, cell_ref.col), patch.clone());
                continue;
            }
        }
    }

    if effective_patches.is_empty() {
        return Ok((original.to_vec(), false));
    }
    let patches = &effective_patches;

    // Track whether any formulas actually changed (added/removed/updated) so we can apply the
    // workbook recalculation policy. This is computed while patching so no-op patches don't churn
    // calc state.
    let mut formula_changed = false;

    // Track the bounds of *material* patches (cells that would require a `<c>` element to be
    // inserted, e.g. value/formula/style writes) so we can expand the worksheet
    // `<dimension ref="..."/>` if needed.
    //
    // We don't shrink dimensions (clears), mirroring Excel's typical behavior.
    let has_cell_patches = !patches.cells.is_empty();

    let patch_bounds = if has_cell_patches {
        patch_bounds(patches, style_id_to_xf)?
    } else {
        None
    };

    // Only insert/patch `<dimension>` when cell patches are present; column metadata edits should
    // not introduce unrelated structural changes.
    let dimension_ref_to_insert = (has_cell_patches && !scan.has_dimension).then(|| {
        let merged = match (scan.existing_used_range, patch_bounds) {
            (Some((min_r, min_c, max_r, max_c)), Some((p_min_r, p_min_c, p_max_r, p_max_c))) => {
                Some((
                    min_r.min(p_min_r),
                    min_c.min(p_min_c),
                    max_r.max(p_max_r),
                    max_c.max(p_max_c),
                ))
            }
            (Some(existing), None) => Some(existing),
            (None, Some(patch)) => Some(patch),
            (None, None) => None,
        };
        merged
            .map(|(min_r, min_c, max_r, max_c)| format_dimension(min_r, min_c, max_r, max_c))
            .unwrap_or_else(|| "A1".to_string())
    });

    let row_patches = patches.by_row();
    let mut remaining_patch_rows: Vec<u32> = row_patches.keys().copied().collect();
    let mut patch_row_idx = 0usize;

    let mut reader = Reader::from_reader(original);
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(
        original.len() + patches.cells.len() * 64,
    ));

    let mut buf = Vec::new();
    let mut worksheet_prefix: Option<String> = None;
    let mut worksheet_has_default_ns = false;
    let mut saw_sheet_data = false;
    let mut inserted_dimension = false;
    let col_properties = patches.col_properties.as_ref();
    let mut cols_written = false;
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

                if !inserted_dimension && !scan.has_sheet_pr {
                    if let Some(ref_str) = dimension_ref_to_insert.as_deref() {
                        let dim_prefix = if worksheet_has_default_ns {
                            None
                        } else {
                            worksheet_prefix.as_deref()
                        };
                        let dim_tag = prefixed_tag(dim_prefix, "dimension");
                        let mut dim = BytesStart::new(dim_tag.as_str());
                        dim.push_attribute(("ref", ref_str));
                        writer.write_event(Event::Empty(dim))?;
                        inserted_dimension = true;
                    }
                }
            }
            Event::Start(e)
                if col_properties.is_some() && local_name(e.name().as_ref()) == b"cols" =>
            {
                let name = e.name();
                let prefix =
                    element_prefix(name.as_ref()).and_then(|p| std::str::from_utf8(p).ok());
                let mut attrs_by_col = {
                    let mut cols_buf = Vec::new();
                    parse_cols_attribute_map_from_reader(&mut reader, &mut cols_buf)?
                };
                merge_col_properties_into_attrs_by_col(
                    &mut attrs_by_col,
                    col_properties.expect("checked is_some above"),
                );
                if !cols_written {
                    let cols_xml = render_cols_xml_from_attrs_by_col(prefix, &attrs_by_col);
                    if !cols_xml.is_empty() {
                        writer.get_mut().extend_from_slice(cols_xml.as_bytes());
                    }
                    cols_written = true;
                }
            }
            Event::Empty(e)
                if col_properties.is_some() && local_name(e.name().as_ref()) == b"cols" =>
            {
                let name = e.name();
                let prefix =
                    element_prefix(name.as_ref()).and_then(|p| std::str::from_utf8(p).ok());
                let mut attrs_by_col = BTreeMap::new();
                merge_col_properties_into_attrs_by_col(
                    &mut attrs_by_col,
                    col_properties.expect("checked is_some above"),
                );
                if !cols_written {
                    let cols_xml = render_cols_xml_from_attrs_by_col(prefix, &attrs_by_col);
                    if !cols_xml.is_empty() {
                        writer.get_mut().extend_from_slice(cols_xml.as_bytes());
                    }
                    cols_written = true;
                }
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"sheetPr" => {
                writer.write_event(Event::Empty(e.into_owned()))?;
                if !inserted_dimension && scan.has_sheet_pr {
                    if let Some(ref_str) = dimension_ref_to_insert.as_deref() {
                        let dim_prefix = if worksheet_has_default_ns {
                            None
                        } else {
                            worksheet_prefix.as_deref()
                        };
                        let dim_tag = prefixed_tag(dim_prefix, "dimension");
                        let mut dim = BytesStart::new(dim_tag.as_str());
                        dim.push_attribute(("ref", ref_str));
                        writer.write_event(Event::Empty(dim))?;
                        inserted_dimension = true;
                    }
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"sheetPr" => {
                writer.write_event(Event::End(e.into_owned()))?;
                if !inserted_dimension && scan.has_sheet_pr {
                    if let Some(ref_str) = dimension_ref_to_insert.as_deref() {
                        let dim_prefix = if worksheet_has_default_ns {
                            None
                        } else {
                            worksheet_prefix.as_deref()
                        };
                        let dim_tag = prefixed_tag(dim_prefix, "dimension");
                        let mut dim = BytesStart::new(dim_tag.as_str());
                        dim.push_attribute(("ref", ref_str));
                        writer.write_event(Event::Empty(dim))?;
                        inserted_dimension = true;
                    }
                }
            }
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
                if let Some(col_properties) = col_properties {
                    if !cols_written {
                        let name = e.name();
                        let sheet_prefix =
                            element_prefix(name.as_ref()).and_then(|p| std::str::from_utf8(p).ok());
                        let cols_xml = render_cols_xml(col_properties, sheet_prefix);
                        if !cols_xml.is_empty() {
                            writer.get_mut().extend_from_slice(cols_xml.as_bytes());
                            cols_written = true;
                        }
                    }
                }
                saw_sheet_data = true;
                let sheet_prefix = element_prefix(e.name().as_ref())
                    .and_then(|p| std::str::from_utf8(p).ok())
                    .map(|s| s.to_string());
                writer.write_event(Event::Start(e.into_owned()))?;
                let changed = patch_sheet_data(
                    &mut reader,
                    &mut writer,
                    &row_patches,
                    &mut remaining_patch_rows,
                    &mut patch_row_idx,
                    scan.sheet_uses_row_spans,
                    &mut shared_strings,
                    style_id_to_xf,
                    sheet_prefix.as_deref(),
                    clear_cached_values_on_formula_change,
                )?;
                formula_changed |= changed;
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"sheetData" => {
                if let Some(col_properties) = col_properties {
                    if !cols_written {
                        let name = e.name();
                        let sheet_prefix =
                            element_prefix(name.as_ref()).and_then(|p| std::str::from_utf8(p).ok());
                        let cols_xml = render_cols_xml(col_properties, sheet_prefix);
                        if !cols_xml.is_empty() {
                            writer.get_mut().extend_from_slice(cols_xml.as_bytes());
                            cols_written = true;
                        }
                    }
                }
                saw_sheet_data = true;
                if row_patches.is_empty() || patch_bounds.is_none() {
                    writer.write_event(Event::Empty(e.into_owned()))?;
                } else {
                    // Convert `<sheetData/>` into `<sheetData>...</sheetData>`.
                    let sheet_data_tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();
                    let sheet_prefix = element_prefix(e.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                    writer.write_event(Event::Start(e.into_owned()))?;
                    for row in remaining_patch_rows.iter().skip(patch_row_idx).copied() {
                        let cells = row_patches.get(&row).map(Vec::as_slice).unwrap_or_default();
                        formula_changed |= cells.iter().any(|(_, patch)| patch_has_formula(patch));
                        write_new_row(
                            &mut writer,
                            row,
                            cells,
                            scan.sheet_uses_row_spans,
                            &mut shared_strings,
                            style_id_to_xf,
                            sheet_prefix.as_deref(),
                            clear_cached_values_on_formula_change,
                        )?;
                    }
                    patch_row_idx = remaining_patch_rows.len();
                    writer.write_event(Event::End(BytesEnd::new(sheet_data_tag.as_str())))?;
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"worksheet" => {
                if !cols_written {
                    if let Some(col_properties) = col_properties {
                        let prefix = if worksheet_has_default_ns {
                            None
                        } else {
                            worksheet_prefix.as_deref()
                        };
                        let cols_xml = render_cols_xml(col_properties, prefix);
                        if !cols_xml.is_empty() {
                            writer.get_mut().extend_from_slice(cols_xml.as_bytes());
                            cols_written = true;
                        }
                    }
                }
                if !saw_sheet_data && !row_patches.is_empty() {
                    // Insert missing <sheetData> just before </worksheet>.
                    let sheet_prefix = if worksheet_has_default_ns {
                        None
                    } else {
                        worksheet_prefix.as_deref()
                    };
                    let sheet_data_tag = prefixed_tag(sheet_prefix, "sheetData");
                    writer.write_event(Event::Start(BytesStart::new(sheet_data_tag.as_str())))?;
                    for row in remaining_patch_rows.iter().skip(patch_row_idx).copied() {
                        let cells = row_patches.get(&row).map(Vec::as_slice).unwrap_or_default();
                        formula_changed |= cells.iter().any(|(_, patch)| patch_has_formula(patch));
                        write_new_row(
                            &mut writer,
                            row,
                            cells,
                            scan.sheet_uses_row_spans,
                            &mut shared_strings,
                            style_id_to_xf,
                            sheet_prefix,
                            clear_cached_values_on_formula_change,
                        )?;
                    }
                    patch_row_idx = remaining_patch_rows.len();
                    writer.write_event(Event::End(BytesEnd::new(sheet_data_tag.as_str())))?;
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

#[derive(Clone, Debug, PartialEq)]
struct ColXmlProps {
    width: Option<f32>,
    hidden: bool,
}

fn escape_xml_attr(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

pub(crate) fn render_cols_xml_from_attrs_by_col(
    prefix: Option<&str>,
    attrs_by_col: &BTreeMap<u32, BTreeMap<String, String>>,
) -> String {
    if attrs_by_col.is_empty() {
        return String::new();
    }

    let cols_tag = prefixed_tag(prefix, "cols");
    let col_tag = prefixed_tag(prefix, "col");

    let mut out = String::new();
    out.push('<');
    out.push_str(&cols_tag);
    out.push('>');

    let mut current: Option<(u32, u32, BTreeMap<String, String>)> = None;
    for (&col_1, attrs) in attrs_by_col {
        let attrs = attrs.clone();
        match current.take() {
            None => current = Some((col_1, col_1, attrs)),
            Some((start, end, cur)) if col_1 == end + 1 && attrs == cur => {
                current = Some((start, col_1, cur));
            }
            Some((start, end, cur)) => {
                out.push_str(&render_col_range_from_attrs(&col_tag, start, end, &cur));
                current = Some((col_1, col_1, attrs));
            }
        }
    }
    if let Some((start, end, cur)) = current {
        out.push_str(&render_col_range_from_attrs(&col_tag, start, end, &cur));
    }

    out.push_str("</");
    out.push_str(&cols_tag);
    out.push('>');
    out
}

fn render_col_range_from_attrs(
    col_tag: &str,
    start_col_1: u32,
    end_col_1: u32,
    attrs: &BTreeMap<String, String>,
) -> String {
    let mut s = String::new();
    s.push_str(&format!(
        r#"<{col_tag} min="{start_col_1}" max="{end_col_1}""#
    ));
    for (key, value) in attrs {
        s.push(' ');
        s.push_str(key);
        s.push_str("=\"");
        s.push_str(&escape_xml_attr(value));
        s.push('"');
    }
    s.push_str("/>");
    s
}

pub(crate) fn merge_col_properties_into_attrs_by_col(
    attrs_by_col: &mut BTreeMap<u32, BTreeMap<String, String>>,
    col_properties: &BTreeMap<u32, ColProperties>,
) {
    let mut touched_cols: BTreeSet<u32> = attrs_by_col.keys().copied().collect();
    touched_cols.extend(col_properties.keys().copied().map(|c0| c0.saturating_add(1)));

    for col_1 in touched_cols {
        if col_1 == 0 || col_1 > formula_model::EXCEL_MAX_COLS {
            continue;
        }
        let col_0 = col_1 - 1;
        let desired_props = col_properties.get(&col_0);
        let desired_width = desired_props.and_then(|p| p.width);
        let desired_hidden = desired_props.map(|p| p.hidden).unwrap_or(false);

        let mut entry = attrs_by_col.remove(&col_1).unwrap_or_default();

        match desired_width {
            Some(width) => {
                entry.insert("width".to_string(), width.to_string());
                entry.insert("customWidth".to_string(), "1".to_string());
            }
            None => {
                entry.remove("width");
                entry.remove("customWidth");
            }
        }

        if desired_hidden {
            entry.insert("hidden".to_string(), "1".to_string());
        } else {
            entry.remove("hidden");
        }

        if !entry.is_empty() {
            attrs_by_col.insert(col_1, entry);
        }
    }
}

fn parse_cols_attribute_map_from_reader<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    buf: &mut Vec<u8>,
) -> Result<BTreeMap<u32, BTreeMap<String, String>>, XlsxError> {
    let mut attrs_by_col: BTreeMap<u32, BTreeMap<String, String>> = BTreeMap::new();

    // We're entering the `<cols>` section after consuming its start tag; consume events until we
    // hit `</cols>`.
    let mut depth: usize = 1;
    loop {
        match reader.read_event_into(buf)? {
            Event::Eof => {
                return Err(XlsxError::Invalid(
                    "unexpected EOF while parsing <cols> section".to_string(),
                ))
            }
            Event::Start(e) => {
                if local_name(e.name().as_ref()) == b"col" {
                    parse_col_element_attrs(&e, &mut attrs_by_col)?;
                }
                depth += 1;
            }
            Event::Empty(e) => {
                if local_name(e.name().as_ref()) == b"col" {
                    parse_col_element_attrs(&e, &mut attrs_by_col)?;
                }
            }
            Event::End(e) => {
                depth = depth.saturating_sub(1);
                if depth == 0 && local_name(e.name().as_ref()) == b"cols" {
                    break;
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(attrs_by_col)
}

fn parse_col_element_attrs(
    e: &BytesStart<'_>,
    attrs_by_col: &mut BTreeMap<u32, BTreeMap<String, String>>,
) -> Result<(), XlsxError> {
    let mut min: Option<u32> = None;
    let mut max: Option<u32> = None;
    let mut element_attrs: Vec<(String, String)> = Vec::new();

    for attr in e.attributes() {
        let attr = attr?;
        let key_bytes = attr.key.as_ref();
        let key = match std::str::from_utf8(key_bytes) {
            Ok(s) => s.to_string(),
            Err(_) => continue,
        };
        let val = attr.unescape_value()?.into_owned();
        match key_bytes {
            b"min" => min = val.parse().ok(),
            b"max" => max = val.parse().ok(),
            _ => element_attrs.push((key, val)),
        }
    }

    let Some(min_1) = min else {
        return Ok(());
    };
    let max_1 = max.unwrap_or(min_1).min(formula_model::EXCEL_MAX_COLS);
    if min_1 == 0 || max_1 == 0 || min_1 > formula_model::EXCEL_MAX_COLS {
        return Ok(());
    }

    for col_1 in min_1..=max_1 {
        let entry = attrs_by_col.entry(col_1).or_default();
        for (k, v) in &element_attrs {
            entry.insert(k.clone(), v.clone());
        }
    }
    Ok(())
}

pub(crate) fn render_cols_xml(
    col_properties: &BTreeMap<u32, ColProperties>,
    prefix: Option<&str>,
) -> String {
    let cols_tag = prefixed_tag(prefix, "cols");
    let col_tag = prefixed_tag(prefix, "col");

    // `col_properties` keys are 0-based; OOXML uses 1-based column indices.
    let mut col_xml_props: BTreeMap<u32, ColXmlProps> = BTreeMap::new();
    for (&col0, props) in col_properties {
        let col_1 = col0.saturating_add(1);
        if col_1 == 0 || col_1 > formula_model::EXCEL_MAX_COLS {
            continue;
        }
        if props.width.is_none() && !props.hidden {
            continue;
        }
        col_xml_props.insert(
            col_1,
            ColXmlProps {
                width: props.width,
                hidden: props.hidden,
            },
        );
    }

    if col_xml_props.is_empty() {
        return String::new();
    }

    let mut out = String::new();
    out.push('<');
    out.push_str(&cols_tag);
    out.push('>');

    let mut current: Option<(u32, u32, ColXmlProps)> = None;
    for (&col_1, props) in col_xml_props.iter() {
        let props = props.clone();
        match current.take() {
            None => current = Some((col_1, col_1, props)),
            Some((start, end, cur)) if col_1 == end + 1 && props == cur => {
                current = Some((start, col_1, cur));
            }
            Some((start, end, cur)) => {
                out.push_str(&render_col_range(&col_tag, start, end, &cur));
                current = Some((col_1, col_1, props));
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

fn render_col_range(col_tag: &str, start_col_1: u32, end_col_1: u32, props: &ColXmlProps) -> String {
    let mut s = String::new();
    s.push_str(&format!(
        r#"<{col_tag} min="{start_col_1}" max="{end_col_1}""#
    ));
    if let Some(width) = props.width {
        s.push_str(&format!(r#" width="{width}""#));
        s.push_str(r#" customWidth="1""#);
    }
    if props.hidden {
        s.push_str(r#" hidden="1""#);
    }
    s.push_str("/>");
    s
}

fn patch_sheet_data<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    writer: &mut Writer<Vec<u8>>,
    row_patches: &BTreeMap<u32, Vec<(u32, &CellPatch)>>,
    remaining_patch_rows: &mut [u32],
    patch_row_idx: &mut usize,
    sheet_uses_row_spans: bool,
    shared_strings: &mut Option<&mut SharedStringsState>,
    style_id_to_xf: Option<&HashMap<u32, u32>>,
    sheet_prefix: Option<&str>,
    clear_cached_values_on_formula_change: bool,
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
                    formula_changed |= cells.iter().any(|(_, patch)| patch_has_formula(patch));
                    write_new_row(
                        writer,
                        row,
                        cells,
                        sheet_uses_row_spans,
                        shared_strings,
                        style_id_to_xf,
                        sheet_prefix,
                        clear_cached_values_on_formula_change,
                    )?;
                    *patch_row_idx += 1;
                }

                if let Some(cells) = row_patches.get(&row_num) {
                    // Consume this patch row if it matches.
                    if *patch_row_idx < remaining_patch_rows.len()
                        && remaining_patch_rows[*patch_row_idx] == row_num
                    {
                        *patch_row_idx += 1;
                    }

                    let row_prefix_owned = element_prefix(row_start.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                    let row_prefix = row_prefix_owned.as_deref().or(sheet_prefix);

                    writer.write_event(Event::Start(rewrite_row_spans(
                        &row_start,
                        cells,
                        style_id_to_xf,
                    )?))?;
                    let changed = patch_row(
                        reader,
                        writer,
                        row_num,
                        cells,
                        shared_strings,
                        style_id_to_xf,
                        row_prefix,
                        clear_cached_values_on_formula_change,
                    )?;
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
                    formula_changed |= cells.iter().any(|(_, patch)| patch_has_formula(patch));
                    write_new_row(
                        writer,
                        row,
                        cells,
                        sheet_uses_row_spans,
                        shared_strings,
                        style_id_to_xf,
                        sheet_prefix,
                        clear_cached_values_on_formula_change,
                    )?;
                    *patch_row_idx += 1;
                }

                if let Some(cells) = row_patches.get(&row_num) {
                    if *patch_row_idx < remaining_patch_rows.len()
                        && remaining_patch_rows[*patch_row_idx] == row_num
                    {
                        *patch_row_idx += 1;
                    }

                    // Convert `<row/>` into `<row>...</row>`.
                    let row_tag = String::from_utf8_lossy(row_empty.name().as_ref()).into_owned();
                    let row_prefix_owned = element_prefix(row_empty.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                    let row_prefix = row_prefix_owned.as_deref().or(sheet_prefix);
                    let mut wrote_any = false;
                    for (col, patch) in cells {
                        if !cell_patch_is_material_for_insertion(patch, style_id_to_xf)? {
                            continue;
                        }
                        if !wrote_any {
                            // Convert `<row/>` into `<row>...</row>`.
                            writer.write_event(Event::Start(rewrite_row_spans(
                                &row_empty,
                                cells,
                                style_id_to_xf,
                            )?))?;
                            wrote_any = true;
                        }
                        formula_changed |= patch_has_formula(patch);
                        write_cell_patch(
                            writer,
                            row_num,
                            *col,
                            patch,
                            None,
                            None,
                            shared_strings,
                            style_id_to_xf,
                            row_prefix,
                            clear_cached_values_on_formula_change,
                        )?;
                    }

                    if wrote_any {
                        writer.write_event(Event::End(BytesEnd::new(row_tag.as_str())))?;
                    } else {
                        // Nothing to insert; preserve the empty row unchanged.
                        writer.write_event(Event::Empty(row_empty))?;
                    }
                } else {
                    writer.write_event(Event::Empty(row_empty))?;
                }
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"sheetData" => {
                // Insert remaining patch rows before closing </sheetData>.
                while *patch_row_idx < remaining_patch_rows.len() {
                    let row = remaining_patch_rows[*patch_row_idx];
                    let cells = row_patches.get(&row).map(Vec::as_slice).unwrap_or_default();
                    formula_changed |= cells.iter().any(|(_, patch)| patch_has_formula(patch));
                    write_new_row(
                        writer,
                        row,
                        cells,
                        sheet_uses_row_spans,
                        shared_strings,
                        style_id_to_xf,
                        sheet_prefix,
                        clear_cached_values_on_formula_change,
                    )?;
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

fn patch_cols_bounds(
    patches: &[(u32, &CellPatch)],
    style_id_to_xf: Option<&HashMap<u32, u32>>,
) -> Result<Option<(u32, u32)>, XlsxError> {
    let mut min_c = u32::MAX;
    let mut max_c = 0u32;
    for (col_0, patch) in patches {
        if !cell_patch_is_material_for_insertion(patch, style_id_to_xf)? {
            continue;
        }
        let col_1 = col_0.saturating_add(1);
        min_c = min_c.min(col_1);
        max_c = max_c.max(col_1);
    }
    Ok((min_c != u32::MAX).then_some((min_c, max_c)))
}

fn parse_row_spans(spans: &str) -> Option<(u32, u32)> {
    let s = spans.trim();
    let (a, b) = s.split_once(':').unwrap_or((s, s));
    let min = a.parse::<u32>().ok()?;
    let max = b.parse::<u32>().ok()?;
    Some((min, max))
}

fn rewrite_row_spans(
    row: &BytesStart<'_>,
    patches: &[(u32, &CellPatch)],
    style_id_to_xf: Option<&HashMap<u32, u32>>,
) -> Result<BytesStart<'static>, XlsxError> {
    let Some((p_min_c, p_max_c)) = patch_cols_bounds(patches, style_id_to_xf)? else {
        return Ok(row.to_owned());
    };

    let mut existing_spans: Option<(u32, u32)> = None;
    for attr in row.attributes() {
        let attr = attr?;
        if local_name(attr.key.as_ref()) == b"spans" {
            existing_spans = parse_row_spans(&attr.unescape_value()?.into_owned());
            break;
        }
    }
    let Some((min_c, max_c)) = existing_spans else {
        return Ok(row.to_owned());
    };

    let new_min = min_c.min(p_min_c);
    let new_max = max_c.max(p_max_c);
    if new_min == min_c && new_max == max_c {
        return Ok(row.to_owned());
    }

    let new_spans = format!("{new_min}:{new_max}");
    let name = row.name();
    let name = std::str::from_utf8(name.as_ref()).unwrap_or("row");
    let mut updated = BytesStart::new(name);
    for attr in row.attributes() {
        let attr = attr?;
        if local_name(attr.key.as_ref()) == b"spans" {
            updated.push_attribute((attr.key.as_ref(), new_spans.as_bytes()));
        } else {
            updated.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
        }
    }
    Ok(updated.into_owned())
}

fn patch_row<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    writer: &mut Writer<Vec<u8>>,
    row_num: u32,
    patches: &[(u32, &CellPatch)],
    shared_strings: &mut Option<&mut SharedStringsState>,
    style_id_to_xf: Option<&HashMap<u32, u32>>,
    default_prefix: Option<&str>,
    clear_cached_values_on_formula_change: bool,
) -> Result<bool, XlsxError> {
    let mut buf = Vec::new();
    let mut patch_idx = 0usize;
    let mut formula_changed = false;
    let mut cell_depth = 0usize;
    let mut cell_prefix: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf)? {
            ev if cell_depth > 0 => {
                match &ev {
                    Event::Start(_) => cell_depth += 1,
                    Event::End(_) => cell_depth = cell_depth.saturating_sub(1),
                    _ => {}
                }
                writer.write_event(ev.into_owned())?;
            }
            Event::Start(e) if local_name(e.name().as_ref()) == b"c" => {
                let cell_start = e.into_owned();
                if cell_prefix.is_none() {
                    cell_prefix = element_prefix(cell_start.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                let insert_prefix = cell_prefix.as_deref().or(default_prefix);
                let Some((cell_ref, existing_t, existing_s)) =
                    parse_cell_addr_and_attrs(&cell_start)?
                else {
                    writer.write_event(Event::Start(cell_start))?;
                    cell_depth = 1;
                    continue;
                };

                if cell_ref.row + 1 != row_num {
                    // Defensive: mismatched cell refs are preserved unchanged.
                    writer.write_event(Event::Start(cell_start))?;
                    cell_depth = 1;
                    continue;
                }

                let col = cell_ref.col;
                while patch_idx < patches.len() && patches[patch_idx].0 < col {
                    let (patch_col, patch) = patches[patch_idx];
                    if cell_patch_is_material_for_insertion(patch, style_id_to_xf)? {
                        formula_changed |= patch_has_formula(patch);
                        write_cell_patch(
                            writer,
                            row_num,
                            patch_col,
                            patch,
                            None,
                            None,
                            shared_strings,
                            style_id_to_xf,
                            insert_prefix,
                            clear_cached_values_on_formula_change,
                        )?;
                    }
                    patch_idx += 1;
                }

                if patch_idx < patches.len() && patches[patch_idx].0 == col {
                    let patch = patches[patch_idx].1;
                    patch_idx += 1;

                    let mut inner_events = Vec::new();
                    let mut depth = 1usize;
                    let cell_end = loop {
                        match reader.read_event_into(&mut buf)? {
                            Event::Start(inner) => {
                                depth += 1;
                                inner_events.push(Event::Start(inner.into_owned()));
                            }
                            Event::Empty(inner) => {
                                inner_events.push(Event::Empty(inner.into_owned()));
                            }
                            Event::End(inner) => {
                                depth = depth.saturating_sub(1);
                                if depth == 0 && local_name(inner.name().as_ref()) == b"c" {
                                    break inner.into_owned();
                                }
                                inner_events.push(Event::End(inner.into_owned()));
                            }
                            Event::Eof => {
                                return Err(XlsxError::Invalid(
                                    "unexpected EOF while skipping patched cell".to_string(),
                                ))
                            }
                            ev => inner_events.push(ev.into_owned()),
                        }
                        buf.clear();
                    };

                    let cell_formula_changed = patch_cell_element(
                        writer,
                        CellRef::new(row_num - 1, col),
                        patch,
                        cell_start,
                        Some(cell_end),
                        inner_events,
                        existing_t,
                        existing_s,
                        shared_strings,
                        style_id_to_xf,
                        false,
                        clear_cached_values_on_formula_change,
                    )?;
                    formula_changed |= cell_formula_changed;
                } else {
                    writer.write_event(Event::Start(cell_start))?;
                    cell_depth = 1;
                }
            }
            Event::Empty(e) if local_name(e.name().as_ref()) == b"c" => {
                let cell_empty = e.into_owned();
                if cell_prefix.is_none() {
                    cell_prefix = element_prefix(cell_empty.name().as_ref())
                        .and_then(|p| std::str::from_utf8(p).ok())
                        .map(|s| s.to_string());
                }
                let insert_prefix = cell_prefix.as_deref().or(default_prefix);
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
                    if cell_patch_is_material_for_insertion(patch, style_id_to_xf)? {
                        formula_changed |= patch_has_formula(patch);
                        write_cell_patch(
                            writer,
                            row_num,
                            patch_col,
                            patch,
                            None,
                            None,
                            shared_strings,
                            style_id_to_xf,
                            insert_prefix,
                            clear_cached_values_on_formula_change,
                        )?;
                    }
                    patch_idx += 1;
                }

                if patch_idx < patches.len() && patches[patch_idx].0 == col {
                    let patch = patches[patch_idx].1;
                    patch_idx += 1;
                    let cell_formula_changed = patch_cell_element(
                        writer,
                        CellRef::new(row_num - 1, col),
                        patch,
                        cell_empty,
                        None,
                        Vec::new(),
                        existing_t,
                        existing_s,
                        shared_strings,
                        style_id_to_xf,
                        true,
                        clear_cached_values_on_formula_change,
                    )?;
                    formula_changed |= cell_formula_changed;
                } else {
                    writer.write_event(Event::Empty(cell_empty))?;
                }
            }
            Event::Start(e) => {
                // Non-cell element inside the row (eg extLst). Ensure any remaining cell patches are
                // emitted before it so cells stay grouped at the start of the row.
                let insert_prefix = cell_prefix.as_deref().or(default_prefix);
                if patch_idx < patches.len() && local_name(e.name().as_ref()) != b"c" {
                    while patch_idx < patches.len() {
                        let (col, patch) = patches[patch_idx];
                        if cell_patch_is_material_for_insertion(patch, style_id_to_xf)? {
                            formula_changed |= patch_has_formula(patch);
                            write_cell_patch(
                                writer,
                                row_num,
                                col,
                                patch,
                                None,
                                None,
                                shared_strings,
                                style_id_to_xf,
                                insert_prefix,
                                clear_cached_values_on_formula_change,
                            )?;
                        }
                        patch_idx += 1;
                    }
                }
                writer.write_event(Event::Start(e.into_owned()))?;
            }
            Event::Empty(e) => {
                let insert_prefix = cell_prefix.as_deref().or(default_prefix);
                if patch_idx < patches.len() && local_name(e.name().as_ref()) != b"c" {
                    while patch_idx < patches.len() {
                        let (col, patch) = patches[patch_idx];
                        if cell_patch_is_material_for_insertion(patch, style_id_to_xf)? {
                            formula_changed |= patch_has_formula(patch);
                            write_cell_patch(
                                writer,
                                row_num,
                                col,
                                patch,
                                None,
                                None,
                                shared_strings,
                                style_id_to_xf,
                                insert_prefix,
                                clear_cached_values_on_formula_change,
                            )?;
                        }
                        patch_idx += 1;
                    }
                }
                writer.write_event(Event::Empty(e.into_owned()))?;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"row" => {
                let insert_prefix = cell_prefix.as_deref().or(default_prefix);
                while patch_idx < patches.len() {
                    let (col, patch) = patches[patch_idx];
                    if cell_patch_is_material_for_insertion(patch, style_id_to_xf)? {
                        formula_changed |= patch_has_formula(patch);
                        write_cell_patch(
                            writer,
                            row_num,
                            col,
                            patch,
                            None,
                            None,
                            shared_strings,
                            style_id_to_xf,
                            insert_prefix,
                            clear_cached_values_on_formula_change,
                        )?;
                    }
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
    sheet_uses_row_spans: bool,
    shared_strings: &mut Option<&mut SharedStringsState>,
    style_id_to_xf: Option<&HashMap<u32, u32>>,
    prefix: Option<&str>,
    clear_cached_values_on_formula_change: bool,
) -> Result<(), XlsxError> {
    let Some((min_c, max_c)) = patch_cols_bounds(patches, style_id_to_xf)? else {
        return Ok(());
    };

    let row_tag = prefixed_tag(prefix, "row");
    let mut row = BytesStart::new(row_tag.as_str());
    let row_num_str = row_num.to_string();
    row.push_attribute(("r", row_num_str.as_str()));

    if sheet_uses_row_spans {
        let spans_str = format!("{min_c}:{max_c}");
        row.push_attribute(("spans", spans_str.as_str()));
    }

    writer.write_event(Event::Start(row))?;
    for (col, patch) in patches {
        if !cell_patch_is_material_for_insertion(patch, style_id_to_xf)? {
            continue;
        }
        write_cell_patch(
            writer,
            row_num,
            *col,
            patch,
            None,
            None,
            shared_strings,
            style_id_to_xf,
            prefix,
            clear_cached_values_on_formula_change,
        )?;
    }

    writer.write_event(Event::End(BytesEnd::new(row_tag.as_str())))?;
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
    style_id_to_xf: Option<&HashMap<u32, u32>>,
    prefix: Option<&str>,
    clear_cached_values_on_formula_change: bool,
) -> Result<bool, XlsxError> {
    let cell_ref = CellRef::new(row_num - 1, col);
    let mut a1 = String::new();
    formula_model::push_a1_cell_ref(cell_ref.row, cell_ref.col, false, false, &mut a1);

    let vm_override = patch.vm_override();
    let cm_override = patch.cm_override();

    // Style: explicit override wins, otherwise preserve existing s=... if present.
    let style_index = patch
        .style_index_override(style_id_to_xf)?
        .or_else(|| existing_s.and_then(|s| s.parse::<u32>().ok()));

    let (value, formula) = match patch {
        CellPatch::Clear { .. } => (None, None),
        CellPatch::Set { value, formula, .. } => (Some(value), formula.as_deref()),
    };
    let formula = match formula {
        Some(formula) if formula_is_material(Some(formula)) => Some(formula),
        _ => None,
    };

    let clear_value = clear_cached_values_on_formula_change && formula.is_some();
    let value = if clear_value { None } else { value };

    let (new_t, body_kind) =
        cell_representation_for_patch(value, formula, existing_t, shared_strings)?;

    if let Some(shared_strings) = shared_strings.as_deref_mut() {
        let old_uses_shared = existing_t == Some("s");
        let new_uses_shared = new_t.as_deref() == Some("s");
        shared_strings.note_shared_string_ref_delta(old_uses_shared, new_uses_shared);
    }

    let cell_tag = prefixed_tag(prefix, "c");
    let formula_tag = prefixed_tag(prefix, "f");
    let v_tag = prefixed_tag(prefix, "v");
    let is_tag = prefixed_tag(prefix, "is");
    let t_tag = prefixed_tag(prefix, "t");

    let mut c = BytesStart::new(cell_tag.as_str());
    c.push_attribute(("r", a1.as_str()));

    let style_value = style_index.filter(|s| *s != 0).map(|s| s.to_string());
    if let Some(style) = style_value.as_deref() {
        c.push_attribute(("s", style));
    }

    if let Some(t) = new_t.as_deref() {
        c.push_attribute(("t", t));
    }

    let vm_value = vm_override.flatten().map(|vm| vm.to_string());
    if let Some(vm) = vm_value.as_deref() {
        c.push_attribute(("vm", vm));
    }
    let cm_value = cm_override.flatten().map(|cm| cm.to_string());
    if let Some(cm) = cm_value.as_deref() {
        c.push_attribute(("cm", cm));
    }

    let has_children = formula.is_some() || !matches!(body_kind, CellBodyKind::None);
    if !has_children {
        writer.write_event(Event::Empty(c))?;
        return Ok(true);
    }

    writer.write_event(Event::Start(c))?;

    if let Some(formula) = formula {
        write_formula_element(writer, None, formula, false, &formula_tag)?;
    }

    if !(clear_cached_values_on_formula_change && formula.is_some()) {
        write_value_element(writer, &body_kind, &v_tag, &is_tag, &t_tag)?;
    }

    writer.write_event(Event::End(BytesEnd::new(cell_tag.as_str())))?;
    Ok(true)
}

fn patch_has_formula(patch: &CellPatch) -> bool {
    match patch {
        CellPatch::Set { formula, .. } => formula_is_material(formula.as_deref()),
        _ => false,
    }
}

#[derive(Debug, Clone)]
struct ExistingCellSemantics {
    formula: Option<String>,
    value: ExistingCellValue,
}

#[derive(Debug, Clone)]
enum ExistingCellValue {
    None,
    Number(f64),
    Boolean(bool),
    Error(String),
    String(String),
    SharedString(RichText),
}

#[derive(Debug, Clone)]
enum CellBodyKind {
    None,
    V(String),
    InlineStr(String),
    InlineRich(RichText),
}

fn patch_cell_element(
    writer: &mut Writer<Vec<u8>>,
    cell_ref: CellRef,
    patch: &CellPatch,
    original_start: BytesStart<'static>,
    original_end: Option<BytesEnd<'static>>,
    inner_events: Vec<Event<'static>>,
    existing_t: Option<String>,
    existing_s: Option<String>,
    shared_strings: &mut Option<&mut SharedStringsState>,
    style_id_to_xf: Option<&HashMap<u32, u32>>,
    original_was_empty: bool,
    clear_cached_values_on_formula_change: bool,
) -> Result<bool, XlsxError> {
    let (patch_value, patch_formula) = match patch {
        CellPatch::Clear { .. } => (None, None),
        CellPatch::Set { value, formula, .. } => (Some(value), formula.as_deref()),
    };
    let patch_formula = match patch_formula {
        Some(formula) if formula_is_material(Some(formula)) => Some(formula),
        _ => None,
    };

    let cell_prefix = element_prefix(original_start.name().as_ref()).map(|p| p.to_vec());
    let existing = parse_existing_cell_semantics(
        existing_t.as_deref(),
        &inner_events,
        shared_strings.as_deref(),
        cell_prefix.as_deref(),
    )?;

    let patch_file_formula = patch_formula.map(formula_to_file_text);
    let formula_eq = match (existing.formula.as_deref(), patch_file_formula.as_deref()) {
        (None, None) => true,
        (Some(a), Some(b)) => a.trim() == b,
        _ => false,
    };
    let value_eq = value_semantics_eq(&existing.value, patch_value);

    let style_override = patch.style_index_override(style_id_to_xf)?;
    let style_change = match style_override {
        None => false,
        Some(0) => existing_s.is_some(),
        Some(xf) => existing_s.as_deref().and_then(|s| s.parse::<u32>().ok()) != Some(xf),
    };

    let vm_override = patch.vm_override();
    let cm_override = patch.cm_override();
    let mut existing_vm: Option<String> = None;
    let mut existing_cm: Option<String> = None;
    if vm_override.is_some() || cm_override.is_some() {
        for attr in original_start.attributes() {
            let attr = attr?;
            match attr.key.as_ref() {
                b"vm" => existing_vm = Some(attr.unescape_value()?.into_owned()),
                b"cm" => existing_cm = Some(attr.unescape_value()?.into_owned()),
                _ => {}
            }
        }
    }
    let vm_change = match vm_override {
        None => false,
        Some(None) => existing_vm.is_some(),
        Some(Some(vm)) => existing_vm
            .as_deref()
            .and_then(|s| s.trim().parse::<u32>().ok())
            != Some(vm),
    };
    let cm_change = match cm_override {
        None => false,
        Some(None) => existing_cm.is_some(),
        Some(Some(cm)) => existing_cm
            .as_deref()
            .and_then(|s| s.trim().parse::<u32>().ok())
            != Some(cm),
    };

    let update_formula = !formula_eq;
    let clear_value =
        clear_cached_values_on_formula_change && update_formula && patch_formula.is_some();
    let update_value = !value_eq || clear_value;
    let any_change = style_change || update_formula || update_value || vm_change || cm_change;

    let clear_cached_value =
        clear_cached_values_on_formula_change && update_formula && patch_formula.is_some();

    if !any_change {
        if original_was_empty {
            writer.write_event(Event::Empty(original_start))?;
        } else {
            writer.write_event(Event::Start(original_start))?;
            for ev in inner_events {
                writer.write_event(ev)?;
            }
            writer.write_event(Event::End(
                original_end.expect("non-empty cell must have end tag"),
            ))?;
        }
        return Ok(false);
    }

    let cell_tag = std::str::from_utf8(original_start.name().as_ref())
        .unwrap_or("c")
        .to_string();
    let prefix_str = cell_tag.rsplit_once(':').map(|(p, _)| p);
    let formula_tag = prefixed_tag(prefix_str, "f");
    let v_tag = prefixed_tag(prefix_str, "v");
    let is_tag = prefixed_tag(prefix_str, "is");
    let t_tag = prefixed_tag(prefix_str, "t");

    let patch_value = if clear_value { None } else { patch_value };
    let (new_t, body_kind) = if update_value {
        cell_representation_for_patch(
            patch_value,
            patch_formula,
            existing_t.as_deref(),
            shared_strings,
        )?
    } else {
        (None, CellBodyKind::None)
    };

    if update_value {
        if let Some(shared_strings) = shared_strings.as_deref_mut() {
            let old_uses_shared = existing_t.as_deref() == Some("s");
            let new_uses_shared = new_t.as_deref() == Some("s");
            shared_strings.note_shared_string_ref_delta(old_uses_shared, new_uses_shared);
        }
    }

    // `vm="..."` points into `xl/metadata.xml` value metadata (rich values / images-in-cell).
    //
    // We preserve `vm` for most rich-data cells even when patching the cached value.
    // The one case where we drop it is when the **original** cell was an embedded-image placeholder
    // represented as an error cell with a `#VALUE!` cached value.
    let existing_is_rich_value_placeholder = matches!(
        &existing.value,
        ExistingCellValue::Error(e) if e.trim() == ErrorValue::Value.as_str()
    );
    let patch_is_rich_value_placeholder =
        matches!(patch_value, Some(CellValue::Error(ErrorValue::Value)));
    let drop_vm =
        update_value && existing_is_rich_value_placeholder && !patch_is_rich_value_placeholder;

    let mut c = BytesStart::new(cell_tag.as_str());
    let mut has_r = false;
    for attr in original_start.attributes() {
        let attr = attr?;
        match attr.key.as_ref() {
            b"s" if style_override.is_some() => continue,
            b"t" if update_value => continue,
            // `vm` points into `xl/metadata.xml` richData/value metadata.
            //
            // Only drop it when editing away from the embedded-image placeholder representation
            // (`t="e"` + `#VALUE!`).
            b"vm" if drop_vm || vm_override.is_some() => continue,
            b"cm" if cm_override.is_some() => continue,
            b"r" => has_r = true,
            _ => {}
        }
        c.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
    }
    if !has_r {
        let mut a1 = String::new();
        formula_model::push_a1_cell_ref(cell_ref.row, cell_ref.col, false, false, &mut a1);
        c.push_attribute(("r", a1.as_str()));
    }

    if let Some(xf) = style_override {
        if xf != 0 {
            let xf_str = xf.to_string();
            c.push_attribute(("s", xf_str.as_str()));
        }
    }

    if update_value {
        if let Some(t) = new_t.as_deref() {
            c.push_attribute(("t", t));
        }
    }

    let vm_value = vm_override.flatten().map(|vm| vm.to_string());
    if let Some(vm) = vm_value.as_deref() {
        c.push_attribute(("vm", vm));
    }
    let cm_value = cm_override.flatten().map(|cm| cm.to_string());
    if let Some(cm) = cm_value.as_deref() {
        c.push_attribute(("cm", cm));
    }

    writer.write_event(Event::Start(c.into_owned()))?;
    write_patched_cell_children(
        writer,
        &inner_events,
        cell_prefix.as_deref(),
        update_formula,
        patch_formula,
        &formula_tag,
        update_value,
        &body_kind,
        clear_cached_value,
        &v_tag,
        &is_tag,
        &t_tag,
    )?;
    writer.write_event(Event::End(BytesEnd::new(cell_tag.as_str())))?;

    Ok(update_formula)
}

fn parse_existing_cell_semantics(
    cell_t: Option<&str>,
    inner_events: &[Event<'static>],
    shared_strings: Option<&SharedStringsState>,
    cell_prefix: Option<&[u8]>,
) -> Result<ExistingCellSemantics, XlsxError> {
    let mut formula: Option<String> = None;
    let mut v_text: Option<String> = None;
    let mut is_value: Option<ExistingCellValue> = None;

    let mut depth = 0usize;
    let mut idx = 0usize;
    while idx < inner_events.len() {
        match &inner_events[idx] {
            Event::Start(e) => {
                if depth == 0 {
                    if formula.is_none() && is_element_named(e.name().as_ref(), cell_prefix, b"f") {
                        let (text, next_idx) = extract_element_text(inner_events, idx)?;
                        formula = Some(text);
                        idx = next_idx;
                        continue;
                    }
                    if v_text.is_none() && is_element_named(e.name().as_ref(), cell_prefix, b"v") {
                        let (text, next_idx) = extract_element_text(inner_events, idx)?;
                        v_text = Some(text);
                        idx = next_idx;
                        continue;
                    }
                    if is_value.is_none() && is_element_named(e.name().as_ref(), cell_prefix, b"is")
                    {
                        let (value, next_idx) =
                            extract_inline_string_value(inner_events, idx, cell_prefix)?;
                        is_value = Some(value);
                        idx = next_idx;
                        continue;
                    }
                }
                depth += 1;
            }
            Event::Empty(e) => {
                if depth == 0 {
                    if formula.is_none() && is_element_named(e.name().as_ref(), cell_prefix, b"f") {
                        formula = Some(String::new());
                    } else if v_text.is_none()
                        && is_element_named(e.name().as_ref(), cell_prefix, b"v")
                    {
                        v_text = Some(String::new());
                    } else if is_value.is_none()
                        && is_element_named(e.name().as_ref(), cell_prefix, b"is")
                    {
                        is_value = Some(ExistingCellValue::String(String::new()));
                    }
                }
            }
            Event::End(_) => {
                depth = depth.saturating_sub(1);
            }
            _ => {}
        }
        idx += 1;
    }

    let value = if let Some(value) = is_value {
        value
    } else if let Some(v) = v_text {
        match cell_t.unwrap_or("n") {
            "b" => ExistingCellValue::Boolean(v.trim() == "1"),
            "e" => ExistingCellValue::Error(v),
            "s" => {
                if let (Some(ss), Ok(idx)) = (shared_strings, v.trim().parse::<u32>()) {
                    if let Some(item) = ss.rich_at(idx) {
                        ExistingCellValue::SharedString(item.clone())
                    } else {
                        ExistingCellValue::String(v)
                    }
                } else {
                    ExistingCellValue::String(v)
                }
            }
            "str" | "inlineStr" => ExistingCellValue::String(v),
            other if should_preserve_unknown_t(other) => ExistingCellValue::String(v),
            _ => v
                .trim()
                .parse::<f64>()
                .map(ExistingCellValue::Number)
                .unwrap_or_else(|_| ExistingCellValue::String(v)),
        }
    } else {
        ExistingCellValue::None
    };

    Ok(ExistingCellSemantics { formula, value })
}

fn extract_element_text(
    events: &[Event<'static>],
    start_idx: usize,
) -> Result<(String, usize), XlsxError> {
    let mut out = String::new();
    let mut idx = start_idx + 1;
    let mut depth = 1usize;
    while idx < events.len() {
        match &events[idx] {
            Event::Start(_) => depth += 1,
            Event::Empty(_) => {}
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return Ok((out, idx + 1));
                }
            }
            Event::Text(t) => out.push_str(&t.unescape()?.into_owned()),
            Event::CData(t) => out.push_str(&String::from_utf8_lossy(t.as_ref())),
            _ => {}
        }
        idx += 1;
    }
    Ok((out, events.len()))
}

fn extract_inline_string_value(
    events: &[Event<'static>],
    start_idx: usize,
    cell_prefix: Option<&[u8]>,
) -> Result<(ExistingCellValue, usize), XlsxError> {
    let mut segments: Vec<(String, RichTextRunStyle)> = Vec::new();
    let mut idx = start_idx + 1;
    let mut depth = 1usize;
    while idx < events.len() {
        match &events[idx] {
            Event::Start(e) => {
                if depth == 1 {
                    match local_name(e.name().as_ref()) {
                        b"t" => {
                            let (text, next_idx) = extract_element_text(events, idx)?;
                            segments.push((text, RichTextRunStyle::default()));
                            idx = next_idx;
                            continue;
                        }
                        b"r" => {
                            let (segment, next_idx) =
                                extract_inline_string_run(events, idx, cell_prefix)?;
                            segments.push(segment);
                            idx = next_idx;
                            continue;
                        }
                        _ => {
                            // Skip non-visible subtrees (phonetic/ruby annotations, extensions, etc.)
                            idx = skip_owned_subtree(events, idx);
                            continue;
                        }
                    }
                }
                depth += 1;
            }
            Event::Empty(e) => {
                if depth == 1 {
                    match local_name(e.name().as_ref()) {
                        b"t" => segments.push((String::new(), RichTextRunStyle::default())),
                        b"r" => segments.push((String::new(), RichTextRunStyle::default())),
                        _ => {}
                    }
                }
            }
            Event::End(e) => {
                depth = depth.saturating_sub(1);
                if depth == 0 && is_element_named(e.name().as_ref(), cell_prefix, b"is") {
                    let plain = segments
                        .iter()
                        .map(|(t, _)| t.as_str())
                        .collect::<String>();
                    if segments.iter().all(|(_, style)| style.is_empty()) {
                        return Ok((ExistingCellValue::String(plain), idx + 1));
                    }
                    return Ok((
                        ExistingCellValue::SharedString(RichText::from_segments(segments)),
                        idx + 1,
                    ));
                }
            }
            _ => {}
        }
        idx += 1;
    }

    Ok((ExistingCellValue::String(String::new()), events.len()))
}

fn extract_inline_string_run(
    events: &[Event<'static>],
    start_idx: usize,
    cell_prefix: Option<&[u8]>,
) -> Result<((String, RichTextRunStyle), usize), XlsxError> {
    let mut style = RichTextRunStyle::default();
    let mut text = String::new();
    let mut idx = start_idx + 1;
    let mut depth = 1usize;
    while idx < events.len() {
        match &events[idx] {
            Event::Start(e) => {
                if depth == 1 {
                    match local_name(e.name().as_ref()) {
                        b"rPr" => {
                            let (parsed, next_idx) =
                                extract_inline_string_rpr(events, idx, cell_prefix)?;
                            style = parsed;
                            idx = next_idx;
                            continue;
                        }
                        b"t" => {
                            let (t, next_idx) = extract_element_text(events, idx)?;
                            text.push_str(&t);
                            idx = next_idx;
                            continue;
                        }
                        _ => {
                            idx = skip_owned_subtree(events, idx);
                            continue;
                        }
                    }
                }
                depth += 1;
            }
            Event::End(e) => {
                depth = depth.saturating_sub(1);
                if depth == 0 && is_element_named(e.name().as_ref(), cell_prefix, b"r") {
                    return Ok(((text, style), idx + 1));
                }
            }
            Event::Empty(e) => {
                if depth == 1 && local_name(e.name().as_ref()) == b"t" {
                    // `<t/>` is legal (empty text run).
                }
            }
            _ => {}
        }
        idx += 1;
    }

    Ok(((text, style), events.len()))
}

fn extract_inline_string_rpr(
    events: &[Event<'static>],
    start_idx: usize,
    _cell_prefix: Option<&[u8]>,
) -> Result<(RichTextRunStyle, usize), XlsxError> {
    let mut style = RichTextRunStyle::default();
    let mut idx = start_idx + 1;
    while idx < events.len() {
        match &events[idx] {
            Event::Empty(e) => {
                parse_inline_string_rpr_tag(e, &mut style)?;
            }
            Event::Start(e) => {
                parse_inline_string_rpr_tag(e, &mut style)?;
                idx = skip_owned_subtree(events, idx);
                continue;
            }
            Event::End(e) if local_name(e.name().as_ref()) == b"rPr" => {
                return Ok((style, idx + 1));
            }
            _ => {}
        }
        idx += 1;
    }
    Ok((style, events.len()))
}

fn parse_inline_string_rpr_tag(
    e: &BytesStart<'_>,
    style: &mut RichTextRunStyle,
) -> Result<(), XlsxError> {
    match local_name(e.name().as_ref()) {
        b"b" => style.bold = Some(parse_bool_val(e)?),
        b"i" => style.italic = Some(parse_bool_val(e)?),
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

fn attr_value(e: &BytesStart<'_>, key: &[u8]) -> Result<Option<String>, XlsxError> {
    for attr in e.attributes().with_checks(false) {
        let attr = attr?;
        if local_name(attr.key.as_ref()) == key {
            return Ok(Some(attr.unescape_value()?.into_owned()));
        }
    }
    Ok(None)
}

fn parse_bool_val(e: &BytesStart<'_>) -> Result<bool, XlsxError> {
    let Some(val) = attr_value(e, b"val")? else {
        return Ok(true);
    };
    Ok(!(val == "0" || val.eq_ignore_ascii_case("false")))
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

fn formula_is_material(formula: Option<&str>) -> bool {
    let Some(formula) = formula else {
        return false;
    };
    !crate::formula_text::normalize_display_formula(formula).is_empty()
}

fn formula_to_file_text(formula: &str) -> String {
    let display = crate::formula_text::normalize_display_formula(formula);
    crate::formula_text::add_xlfn_prefixes(&display)
}

fn value_semantics_eq(existing: &ExistingCellValue, patch_value: Option<&CellValue>) -> bool {
    let Some(patch_value) = patch_value else {
        return matches!(existing, ExistingCellValue::None);
    };

    match patch_value {
        CellValue::Empty => matches!(existing, ExistingCellValue::None),
        CellValue::Number(n) => matches!(existing, ExistingCellValue::Number(m) if m == n),
        CellValue::Boolean(b) => matches!(existing, ExistingCellValue::Boolean(m) if m == b),
        CellValue::Error(err) => {
            matches!(existing, ExistingCellValue::Error(e) if e == err.as_str())
        }
        CellValue::String(s) => match existing {
            ExistingCellValue::String(v) => v == s,
            ExistingCellValue::SharedString(rich) => &rich.text == s,
            _ => false,
        },
        CellValue::Entity(entity) => {
            let display = entity.display_value.as_str();
            match existing {
                ExistingCellValue::String(v) => v == display,
                ExistingCellValue::SharedString(rich) => rich.text == display,
                _ => false,
            }
        },
        CellValue::Record(record) => {
            let display = record.to_string();
            let display = display.as_str();
            match existing {
                ExistingCellValue::String(v) => v == display,
                ExistingCellValue::SharedString(rich) => rich.text == display,
                _ => false,
            }
        },
        CellValue::Image(image) => match image.alt_text.as_deref().filter(|s| !s.is_empty()) {
            Some(alt) => match existing {
                ExistingCellValue::String(v) => v == alt,
                ExistingCellValue::SharedString(rich) => rich.text == alt,
                _ => false,
            },
            None => matches!(existing, ExistingCellValue::None),
        },
        CellValue::RichText(rich) => match existing {
            ExistingCellValue::SharedString(existing_rich) => {
                rich_text_semantics_eq(existing_rich, rich)
            }
            ExistingCellValue::String(v) => &rich.text == v && rich_text_has_no_formatting(rich),
            _ => false,
        },
        CellValue::Array(_) | CellValue::Spill(_) => false,
    }
}

fn rich_text_has_no_formatting(rich: &RichText) -> bool {
    // `RichText::is_plain()` only checks for an empty `runs` array. For patch semantics we treat
    // runs with empty styles (and zero-length runs) as plain too, since they carry no formatting.
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
    // This allows callers to represent rich text as a sparse set of style overrides (only
    // including non-empty runs), while still comparing equal to the fully segmented
    // `RichText::from_segments` representation used by XLSX parsers.
    let mut runs: Vec<RichTextRun> = rich
        .runs
        .iter()
        .filter(|run| run.start < run.end && !run.style.is_empty())
        .cloned()
        .collect();
    runs.sort_by_key(|run| (run.start, run.end));

    let mut merged: Vec<RichTextRun> = Vec::with_capacity(runs.len());
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

fn cell_representation_for_patch(
    value: Option<&CellValue>,
    formula: Option<&str>,
    existing_t: Option<&str>,
    shared_strings: &mut Option<&mut SharedStringsState>,
) -> Result<(Option<String>, CellBodyKind), XlsxError> {
    let formula = match formula {
        Some(formula) if formula_is_material(Some(formula)) => Some(formula),
        _ => None,
    };
    let Some(value) = value else {
        return Ok((None, CellBodyKind::None));
    };

    match value {
        CellValue::Empty => Ok((None, CellBodyKind::None)),
        CellValue::Number(n) => Ok((None, CellBodyKind::V(n.to_string()))),
        CellValue::Boolean(b) => Ok((
            Some("b".to_string()),
            CellBodyKind::V(if *b { "1" } else { "0" }.to_string()),
        )),
        CellValue::Error(e) => Ok((
            Some("e".to_string()),
            CellBodyKind::V(e.as_str().to_string()),
        )),
        CellValue::String(s) => {
            if let Some(existing_t) = existing_t {
                if should_preserve_unknown_t(existing_t) {
                    return Ok((Some(existing_t.to_string()), CellBodyKind::V(s.clone())));
                }
                match existing_t {
                    "inlineStr" => {
                        return Ok((
                            Some("inlineStr".to_string()),
                            CellBodyKind::InlineStr(s.clone()),
                        ));
                    }
                    "str" => return Ok((Some("str".to_string()), CellBodyKind::V(s.clone()))),
                    _ => {}
                }
            }

            if shared_strings.is_some() {
                let wants_shared = match existing_t {
                    Some("s") => true,
                    // Match streaming patcher behavior: prefer `t="str"` for formula string results
                    // unless the original cell already used shared strings.
                    _ => formula.is_none(),
                };
                if wants_shared {
                    let idx = shared_strings
                        .as_deref_mut()
                        .map(|ss| ss.get_or_insert_plain(s))
                        .unwrap_or(0);
                    return Ok((Some("s".to_string()), CellBodyKind::V(idx.to_string())));
                }
            }

            if formula.is_some() {
                Ok((Some("str".to_string()), CellBodyKind::V(s.clone())))
            } else {
                Ok((
                    Some("inlineStr".to_string()),
                    CellBodyKind::InlineStr(s.clone()),
                ))
            }
        }
        CellValue::Entity(entity) => {
            let degraded = CellValue::String(entity.display_value.clone());
            cell_representation_for_patch(Some(&degraded), formula, existing_t, shared_strings)
        },
        CellValue::Record(record) => {
            let degraded = CellValue::String(record.to_string());
            cell_representation_for_patch(Some(&degraded), formula, existing_t, shared_strings)
        },
        CellValue::Image(image) => {
            if let Some(alt) = image.alt_text.as_deref().filter(|s| !s.is_empty()) {
                let degraded = CellValue::String(alt.to_string());
                cell_representation_for_patch(Some(&degraded), formula, existing_t, shared_strings)
            } else {
                Ok((None, CellBodyKind::None))
            }
        },
        CellValue::RichText(rich) => {
            if let Some(existing_t) = existing_t {
                if should_preserve_unknown_t(existing_t) {
                    return Ok((
                        Some(existing_t.to_string()),
                        CellBodyKind::V(rich.text.clone()),
                    ));
                }
                if existing_t == "inlineStr" {
                    return Ok((
                        Some("inlineStr".to_string()),
                        CellBodyKind::InlineRich(rich.clone()),
                    ));
                }
            }

            let prefer_shared = shared_strings.is_some() && existing_t != Some("inlineStr");
            if prefer_shared {
                let idx = shared_strings
                    .as_deref_mut()
                    .map(|ss| ss.get_or_insert_rich(rich))
                    .unwrap_or(0);
                Ok((Some("s".to_string()), CellBodyKind::V(idx.to_string())))
            } else {
                Ok((
                    Some("inlineStr".to_string()),
                    CellBodyKind::InlineRich(rich.clone()),
                ))
            }
        }
        CellValue::Array(_) | CellValue::Spill(_) => Err(XlsxError::Invalid(format!(
            "unsupported cell value type for patch: {value:?}"
        ))),
    }
}

fn write_patched_cell_children(
    writer: &mut Writer<Vec<u8>>,
    inner_events: &[Event<'static>],
    cell_prefix: Option<&[u8]>,
    update_formula: bool,
    patch_formula: Option<&str>,
    formula_tag: &str,
    update_value: bool,
    body_kind: &CellBodyKind,
    clear_cached_value: bool,
    v_tag: &str,
    is_tag: &str,
    t_tag: &str,
) -> Result<(), XlsxError> {
    // If we need to write a new/updated `<f>` element before we've encountered the original `<f>`
    // node, preserve the original formula attributes when possible.
    //
    // This matters for worksheets where the `<f>` element appears *after* `<v>/<is>` (unusual, but
    // legal) since the patcher inserts `<f>` before the value for stability.
    let formula_template = if update_formula && patch_formula.is_some() {
        let mut depth = 0usize;
        let mut template = None;
        for ev in inner_events {
            match ev {
                Event::Start(e) => {
                    if depth == 0 && is_element_named(e.name().as_ref(), cell_prefix, b"f") {
                        template = Some(e.clone());
                        break;
                    }
                    depth += 1;
                }
                Event::Empty(e) => {
                    if depth == 0 && is_element_named(e.name().as_ref(), cell_prefix, b"f") {
                        template = Some(e.clone());
                        break;
                    }
                }
                Event::End(_) => depth = depth.saturating_sub(1),
                _ => {}
            }
        }
        template
    } else {
        None
    };

    let mut formula_written = !update_formula || patch_formula.is_none();
    let mut value_written =
        !update_value || matches!(body_kind, CellBodyKind::None) || clear_cached_value;
    let mut saw_formula = false;
    let mut saw_value = false;

    let mut idx = 0usize;
    while idx < inner_events.len() {
        match &inner_events[idx] {
            Event::Start(e) if is_element_named(e.name().as_ref(), cell_prefix, b"f") => {
                saw_formula = true;
                if update_formula {
                    if !formula_written {
                        if let Some(formula) = patch_formula {
                            let detach_shared = should_detach_shared_formula(e, formula);
                            write_formula_element(
                                writer,
                                Some(e),
                                formula,
                                detach_shared,
                                formula_tag,
                            )?;
                            formula_written = true;
                        }
                    }
                    idx = skip_owned_subtree(inner_events, idx);
                    continue;
                }

                idx = write_owned_subtree(writer, inner_events, idx)?;
                continue;
            }
            Event::Empty(e) if is_element_named(e.name().as_ref(), cell_prefix, b"f") => {
                saw_formula = true;
                if update_formula {
                    if !formula_written {
                        if let Some(formula) = patch_formula {
                            let detach_shared = should_detach_shared_formula(e, formula);
                            write_formula_element(
                                writer,
                                Some(e),
                                formula,
                                detach_shared,
                                formula_tag,
                            )?;
                            formula_written = true;
                        }
                    }
                } else {
                    writer.write_event(Event::Empty(e.clone()))?;
                }
                idx += 1;
                continue;
            }
            Event::Start(e)
                if is_element_named(e.name().as_ref(), cell_prefix, b"v")
                    || is_element_named(e.name().as_ref(), cell_prefix, b"is") =>
            {
                saw_value = true;

                if update_formula && !formula_written {
                    if let Some(formula) = patch_formula {
                        let detach_shared = formula_template
                            .as_ref()
                            .is_some_and(|f| should_detach_shared_formula(f, formula));
                        write_formula_element(
                            writer,
                            formula_template.as_ref(),
                            formula,
                            detach_shared,
                            formula_tag,
                        )?;
                        formula_written = true;
                    }
                }
                if update_value && !value_written {
                    write_value_element(writer, body_kind, v_tag, is_tag, t_tag)?;
                    value_written = true;
                }

                if update_value || clear_cached_value {
                    idx = skip_owned_subtree(inner_events, idx);
                } else {
                    idx = write_owned_subtree(writer, inner_events, idx)?;
                }
                continue;
            }
            Event::Empty(e)
                if is_element_named(e.name().as_ref(), cell_prefix, b"v")
                    || is_element_named(e.name().as_ref(), cell_prefix, b"is") =>
            {
                saw_value = true;

                if update_formula && !formula_written {
                    if let Some(formula) = patch_formula {
                        let detach_shared = formula_template
                            .as_ref()
                            .is_some_and(|f| should_detach_shared_formula(f, formula));
                        write_formula_element(
                            writer,
                            formula_template.as_ref(),
                            formula,
                            detach_shared,
                            formula_tag,
                        )?;
                        formula_written = true;
                    }
                }
                if update_value && !value_written {
                    write_value_element(writer, body_kind, v_tag, is_tag, t_tag)?;
                    value_written = true;
                }

                if !update_value && !clear_cached_value {
                    writer.write_event(Event::Empty(e.clone()))?;
                }

                idx += 1;
                continue;
            }
            ev => {
                if update_formula && !formula_written && !saw_formula {
                    if let Some(formula) = patch_formula {
                        let detach_shared = formula_template
                            .as_ref()
                            .is_some_and(|f| should_detach_shared_formula(f, formula));
                        write_formula_element(
                            writer,
                            formula_template.as_ref(),
                            formula,
                            detach_shared,
                            formula_tag,
                        )?;
                        formula_written = true;
                    }
                }
                if update_value && !value_written && !saw_value {
                    write_value_element(writer, body_kind, v_tag, is_tag, t_tag)?;
                    value_written = true;
                }

                match ev {
                    Event::Start(_) => {
                        idx = write_owned_subtree(writer, inner_events, idx)?;
                        continue;
                    }
                    _ => writer.write_event(ev.clone())?,
                }
            }
        }
        idx += 1;
    }

    if update_formula && !formula_written {
        if let Some(formula) = patch_formula {
            let detach_shared = formula_template
                .as_ref()
                .is_some_and(|f| should_detach_shared_formula(f, formula));
            write_formula_element(
                writer,
                formula_template.as_ref(),
                formula,
                detach_shared,
                formula_tag,
            )?;
        }
    }
    if update_value && !value_written {
        write_value_element(writer, body_kind, v_tag, is_tag, t_tag)?;
    }

    Ok(())
}

fn write_owned_subtree(
    writer: &mut Writer<Vec<u8>>,
    events: &[Event<'static>],
    mut idx: usize,
) -> Result<usize, XlsxError> {
    match &events[idx] {
        Event::Start(_) => {
            let mut depth = 0usize;
            while idx < events.len() {
                let ev = events[idx].clone();
                match &ev {
                    Event::Start(_) => depth += 1,
                    Event::End(_) => {
                        depth = depth.saturating_sub(1);
                    }
                    _ => {}
                }
                writer.write_event(ev)?;
                idx += 1;
                if depth == 0 {
                    break;
                }
            }
            Ok(idx)
        }
        _ => {
            writer.write_event(events[idx].clone())?;
            Ok(idx + 1)
        }
    }
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
                    _ => {}
                }
                idx += 1;
            }
            idx
        }
        _ => idx + 1,
    }
}

fn write_formula_element(
    writer: &mut Writer<Vec<u8>>,
    original: Option<&BytesStart<'_>>,
    formula: &str,
    detach_shared: bool,
    tag_name: &str,
) -> Result<(), XlsxError> {
    let file_formula = formula_to_file_text(formula);

    let mut f = BytesStart::new(tag_name);
    if let Some(orig) = original {
        for attr in orig.attributes() {
            let attr = attr?;
            if detach_shared && matches!(attr.key.as_ref(), b"t" | b"ref" | b"si") {
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
        writer.write_event(Event::End(BytesEnd::new(tag_name)))?;
    }

    Ok(())
}

fn write_value_element(
    writer: &mut Writer<Vec<u8>>,
    body_kind: &CellBodyKind,
    v_tag: &str,
    is_tag: &str,
    t_tag: &str,
) -> Result<(), XlsxError> {
    match body_kind {
        CellBodyKind::V(text) => {
            writer.write_event(Event::Start(BytesStart::new(v_tag)))?;
            writer.write_event(Event::Text(BytesText::new(text)))?;
            writer.write_event(Event::End(BytesEnd::new(v_tag)))?;
        }
        CellBodyKind::InlineStr(text) => {
            writer.write_event(Event::Start(BytesStart::new(is_tag)))?;
            write_rich_text_t(writer, t_tag, text)?;
            writer.write_event(Event::End(BytesEnd::new(is_tag)))?;
        }
        CellBodyKind::InlineRich(rich) => {
            let prefix = t_tag.rsplit_once(':').map(|(p, _)| p);
            let r_tag = prefixed_tag(prefix, "r");
            let rpr_tag = prefixed_tag(prefix, "rPr");

            writer.write_event(Event::Start(BytesStart::new(is_tag)))?;
            if rich.runs.is_empty() {
                write_rich_text_t(writer, t_tag, &rich.text)?;
            } else {
                for run in &rich.runs {
                    writer.write_event(Event::Start(BytesStart::new(r_tag.as_str())))?;
                    if !run.style.is_empty() {
                        writer.write_event(Event::Start(BytesStart::new(rpr_tag.as_str())))?;
                        write_rich_text_rpr(writer, prefix, &run.style)?;
                        writer.write_event(Event::End(BytesEnd::new(rpr_tag.as_str())))?;
                    }

                    let segment = rich.slice_run_text(run);
                    write_rich_text_t(writer, t_tag, segment)?;

                    writer.write_event(Event::End(BytesEnd::new(r_tag.as_str())))?;
                }
            }
            writer.write_event(Event::End(BytesEnd::new(is_tag)))?;
        }
        CellBodyKind::None => {}
    }

    Ok(())
}

fn write_rich_text_t(
    writer: &mut Writer<Vec<u8>>,
    t_tag: &str,
    text: &str,
) -> Result<(), XlsxError> {
    let mut t = BytesStart::new(t_tag);
    if needs_space_preserve(text) {
        t.push_attribute(("xml:space", "preserve"));
    }
    writer.write_event(Event::Start(t))?;
    writer.write_event(Event::Text(BytesText::new(text)))?;
    writer.write_event(Event::End(BytesEnd::new(t_tag)))?;
    Ok(())
}

fn write_rich_text_rpr(
    writer: &mut Writer<Vec<u8>>,
    prefix: Option<&str>,
    style: &RichTextRunStyle,
) -> Result<(), XlsxError> {
    if let Some(font) = &style.font {
        let mut rfont = BytesStart::new(prefixed_tag(prefix, "rFont"));
        rfont.push_attribute(("val", font.as_str()));
        writer.write_event(Event::Empty(rfont))?;
    }

    if let Some(size_100pt) = style.size_100pt {
        let mut sz = BytesStart::new(prefixed_tag(prefix, "sz"));
        let value = format_size_100pt(size_100pt);
        sz.push_attribute(("val", value.as_str()));
        writer.write_event(Event::Empty(sz))?;
    }

    if let Some(color) = style.color {
        let mut c = BytesStart::new(prefixed_tag(prefix, "color"));
        let value = format!("{:08X}", color.argb().unwrap_or(0));
        c.push_attribute(("rgb", value.as_str()));
        writer.write_event(Event::Empty(c))?;
    }

    if let Some(bold) = style.bold {
        let mut b = BytesStart::new(prefixed_tag(prefix, "b"));
        if !bold {
            b.push_attribute(("val", "0"));
        }
        writer.write_event(Event::Empty(b))?;
    }

    if let Some(italic) = style.italic {
        let mut i = BytesStart::new(prefixed_tag(prefix, "i"));
        if !italic {
            i.push_attribute(("val", "0"));
        }
        writer.write_event(Event::Empty(i))?;
    }

    if let Some(ul) = style.underline {
        let mut u = BytesStart::new(prefixed_tag(prefix, "u"));
        if let Some(val) = ul.to_ooxml() {
            u.push_attribute(("val", val));
        }
        writer.write_event(Event::Empty(u))?;
    }

    Ok(())
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

fn should_detach_shared_formula(f: &BytesStart<'_>, patch_formula: &str) -> bool {
    let trimmed = patch_formula.trim();
    let stripped = trimmed.strip_prefix('=').unwrap_or(trimmed).trim();
    if stripped.is_empty() {
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

fn should_preserve_unknown_t(t: &str) -> bool {
    !matches!(t, "s" | "b" | "e" | "n" | "str" | "inlineStr")
}

fn prefixed_tag(prefix: Option<&str>, local: &str) -> String {
    match prefix {
        Some(prefix) => format!("{prefix}:{local}"),
        None => local.to_string(),
    }
}

fn element_prefix(name: &[u8]) -> Option<&[u8]> {
    name.iter()
        .rposition(|b| *b == b':')
        .map(|idx| &name[..idx])
}

fn worksheet_has_default_spreadsheetml_ns(e: &BytesStart<'_>) -> Result<bool, XlsxError> {
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"xmlns" && attr.value.as_ref() == SPREADSHEETML_NS.as_bytes() {
            return Ok(true);
        }
    }
    Ok(false)
}

fn is_element_named(name: &[u8], expected_prefix: Option<&[u8]>, local: &[u8]) -> bool {
    // Most worksheet parts either:
    // - use the default SpreadsheetML namespace (no prefixes), or
    // - use a consistent explicit prefix for all SpreadsheetML elements.
    //
    // Some producers, however, mix the two: e.g. a prefixed `<x:c>` with unprefixed `<v>/<f>/<is>`
    // children where the unprefixed elements still resolve via a default `xmlns=...` declaration.
    //
    // The patcher is prefix-preserving when *writing*, but should be prefix-tolerant when *reading*
    // existing semantics to avoid unnecessary cell rewrites.
    let _ = expected_prefix;
    local_name(name) == local
}

fn parse_row_r(row: &BytesStart<'_>) -> Result<Option<u32>, XlsxError> {
    for attr in row.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"r" {
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
        let value = attr.unescape_value()?.into_owned();
        match attr.key.as_ref() {
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

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(idx) => &name[idx + 1..],
        None => name,
    }
}

fn needs_space_preserve(text: &str) -> bool {
    text.starts_with(char::is_whitespace) || text.ends_with(char::is_whitespace)
}

fn cell_patch_is_material_for_insertion(
    patch: &CellPatch,
    style_id_to_xf: Option<&HashMap<u32, u32>>,
) -> Result<bool, XlsxError> {
    let style_index = patch.style_index_override(style_id_to_xf)?;
    if style_index.is_some_and(|xf| xf != 0) {
        return Ok(true);
    }

    if patch.vm_override().flatten().is_some() || patch.cm_override().flatten().is_some() {
        return Ok(true);
    }

    match patch {
        CellPatch::Clear { .. } => Ok(false),
        CellPatch::Set { value, formula, .. } => Ok(!matches!(value, CellValue::Empty)
            || formula
                .as_deref()
                .is_some_and(|f| !crate::formula_text::normalize_display_formula(f).is_empty())),
    }
}

fn patch_bounds(
    patches: &WorksheetCellPatches,
    style_id_to_xf: Option<&HashMap<u32, u32>>,
) -> Result<Option<(u32, u32, u32, u32)>, XlsxError> {
    let mut min_row = u32::MAX;
    let mut min_col = u32::MAX;
    let mut max_row = 0u32;
    let mut max_col = 0u32;

    for (cell_ref, patch) in patches.iter() {
        if !cell_patch_is_material_for_insertion(patch, style_id_to_xf)? {
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

    Ok(if min_row == u32::MAX {
        None
    } else {
        Some((min_row, min_col, max_row, max_col))
    })
}

fn rewrite_dimension(
    e: &BytesStart<'_>,
    patch_bounds: (u32, u32, u32, u32),
) -> Result<BytesStart<'static>, XlsxError> {
    let (p_min_r, p_min_c, p_max_r, p_max_c) = patch_bounds;

    // Preserve the original element name (including any prefix) so prefixed worksheets round-trip
    // cleanly even when we update the dimension.
    let tag = String::from_utf8_lossy(e.name().as_ref()).into_owned();

    let mut existing_ref: Option<String> = None;
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"ref" {
            existing_ref = Some(attr.unescape_value()?.into_owned());
            break;
        }
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

    // Preserve attribute ordering where possible by rewriting `ref` in-place.
    let mut start = BytesStart::new(tag.as_str());
    let mut wrote_ref = false;
    for attr in e.attributes() {
        let attr = attr?;
        if attr.key.as_ref() == b"ref" {
            start.push_attribute(("ref", new_ref.as_str()));
            wrote_ref = true;
        } else {
            start.push_attribute((attr.key.as_ref(), attr.value.as_ref()));
        }
    }
    if !wrote_ref {
        start.push_attribute(("ref", new_ref.as_str()));
    }

    Ok(start.into_owned())
}

fn parse_dimension_ref(ref_str: &str) -> Option<(u32, u32, u32, u32)> {
    let s = ref_str.trim();
    let (a, b) = s.split_once(':').unwrap_or((s, s));
    let start = CellRef::from_a1(a).ok()?;
    let end = CellRef::from_a1(b).ok()?;
    Some((start.row + 1, start.col + 1, end.row + 1, end.col + 1))
}

fn format_dimension(min_r: u32, min_c: u32, max_r: u32, max_c: u32) -> String {
    let mut out = String::new();
    formula_model::push_a1_cell_range(
        min_r.saturating_sub(1),
        min_c.saturating_sub(1),
        max_r.saturating_sub(1),
        max_c.saturating_sub(1),
        false,
        false,
        &mut out,
    );
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::XlsxPackage;
    use formula_model::{Style, StyleTable};
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

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.finish().unwrap().into_inner()
    }

    #[test]
    fn apply_cell_patches_with_styles_ignores_external_workbook_styles_relationship() {
        let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

        // External styles relationship is listed first and must be ignored. Otherwise we'd resolve
        // it as an internal part name and fail to load `xl/styles.xml`.
        let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="https://example.com/styles.xml" TargetMode="External"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData/>
</worksheet>"#;

        let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.start_file("xl/styles.xml", options).unwrap();
        zip.write_all(styles_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        let mut style_table = StyleTable::default();
        let style_id = style_table.intern(Style {
            number_format: Some("0".to_string()),
            ..Default::default()
        });
        assert_eq!(style_id, 1);

        let mut patches = WorkbookCellPatches::default();
        patches.set_cell(
            "Sheet1",
            CellRef::from_a1("A1").unwrap(),
            CellPatch::set_value_with_style_id(CellValue::Number(1.0), style_id),
        );

        pkg.apply_cell_patches_with_styles(&patches, &style_table)
            .expect("apply_cell_patches_with_styles");
    }

    #[test]
    fn resolve_shared_strings_part_name_ignores_external_relationship() {
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="https://example.com/sharedStrings.xml" TargetMode="External"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"#;

        let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData/>
</worksheet>"#;

        let shared_strings_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>"#;

        let cursor = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(cursor);
        let options = zip::write::FileOptions::<()>::default()
            .compression_method(zip::CompressionMethod::Deflated);

        zip.start_file("xl/workbook.xml", options).unwrap();
        zip.write_all(workbook_xml.as_bytes()).unwrap();

        zip.start_file("xl/_rels/workbook.xml.rels", options)
            .unwrap();
        zip.write_all(workbook_rels.as_bytes()).unwrap();

        zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
        zip.write_all(worksheet_xml.as_bytes()).unwrap();

        zip.start_file("xl/sharedStrings.xml", options).unwrap();
        zip.write_all(shared_strings_xml.as_bytes()).unwrap();

        let bytes = zip.finish().unwrap().into_inner();
        let pkg = XlsxPackage::from_bytes(&bytes).expect("read pkg");

        assert_eq!(
            resolve_shared_strings_part_name(&pkg).expect("resolve shared strings"),
            Some("xl/sharedStrings.xml".to_string())
        );
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
