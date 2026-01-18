use std::cell::RefCell;
use std::collections::{BTreeMap, HashMap, HashSet};
use std::sync::Arc;

use serde::de::Error as _;
use serde::{Deserialize, Serialize};

use formula_columnar::{ColumnType as ColumnarType, ColumnarTable, Value as ColumnarValue};

use crate::drawings::DrawingObject;
use crate::{
    A1ParseError, Cell, CellKey, CellRef, CellValue, CellValueProvider, CfEvaluationResult, CfRule,
    CfStyleOverride, Comment, CommentError, CommentPatch, ConditionalFormattingEngine,
    DataValidation, DataValidationAssignment, DataValidationId, DifferentialFormatProvider,
    FormulaEvaluator, Hyperlink, MergeError, MergedRegions, Outline, OutlineEntry, Range, Reply,
    SheetAutoFilter, SheetProtection, SheetProtectionAction, SheetSelection, SheetView, StyleTable,
    Table,
};

/// Identifier for a worksheet within a workbook.
pub type WorksheetId = u32;

/// Sheet tab color.
///
/// XLSX stores this as a `CT_Color` payload, which can be specified as:
/// - `rgb` (ARGB hex)
/// - `theme` + `tint`
/// - `indexed`
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
pub struct TabColor {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub rgb: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub theme: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indexed: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tint: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto: Option<bool>,
}

impl TabColor {
    pub fn rgb(rgb: impl Into<String>) -> Self {
        Self {
            rgb: Some(rgb.into()),
            ..Default::default()
        }
    }
}

fn default_row_count() -> u32 {
    crate::cell::EXCEL_MAX_ROWS
}

fn default_col_count() -> u32 {
    crate::cell::EXCEL_MAX_COLS
}

fn default_zoom() -> f32 {
    crate::view::default_zoom()
}

fn is_default_zoom(z: &f32) -> bool {
    crate::view::is_default_zoom(z)
}

fn is_visible(v: &SheetVisibility) -> bool {
    matches!(v, SheetVisibility::Visible)
}

fn is_false(b: &bool) -> bool {
    !*b
}

fn is_default_outline(outline: &Outline) -> bool {
    outline == &Outline::default()
}

/// Sheet visibility state (Excel-compatible).
#[derive(Copy, Clone, Debug, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum SheetVisibility {
    Visible,
    Hidden,
    VeryHidden,
}

impl Default for SheetVisibility {
    fn default() -> Self {
        SheetVisibility::Visible
    }
}

/// Per-row overrides.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
pub struct RowProperties {
    /// Row height in points.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f32>,
    /// Whether the row is user-hidden (eg via "Hide row").
    ///
    /// This is treated as the persisted "user hidden" bit for row visibility. When using
    /// [`Worksheet::set_row_hidden`], this flag is kept in sync with
    /// `Worksheet::outline.rows[*].hidden.user`.
    #[serde(default, skip_serializing_if = "is_false")]
    pub hidden: bool,
    /// Optional default style id for all cells in this row.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_id: Option<u32>,
}

/// Per-column overrides.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
pub struct ColProperties {
    /// Column width in Excel "character" units (OOXML `col/@width`).
    ///
    /// This is the width value shown in Excel's "Column Width" UI and persisted in `.xlsx`
    /// files. It is **not pixels**; the pixel width depends on the workbook's default font
    /// (for Excel's default Calibri 11, max digit width is 7px with 5px padding).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<f32>,
    /// Whether the column is user-hidden (eg via "Hide column").
    ///
    /// This is treated as the persisted "user hidden" bit for column visibility. When using
    /// [`Worksheet::set_col_hidden`], this flag is kept in sync with
    /// `Worksheet::outline.cols[*].hidden.user`.
    #[serde(default, skip_serializing_if = "is_false")]
    pub hidden: bool,
    /// Optional default style id for all cells in this column.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_id: Option<u32>,
}

#[derive(Clone, Debug)]
struct ColumnarBackend {
    origin: CellRef,
    table: Arc<ColumnarTable>,
}

/// A worksheet (sheet tab) containing sparse cells and per-row/column metadata.
#[derive(Clone, Debug, Serialize)]
pub struct Worksheet {
    /// Stable worksheet identifier.
    pub id: WorksheetId,
    /// User-visible name.
    pub name: String,

    /// XLSX `sheetId` value preserved for round-trip fidelity.
    ///
    /// This is distinct from [`Worksheet::id`], which is the internal stable id
    /// used by the in-memory model.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub xlsx_sheet_id: Option<u32>,

    /// XLSX workbook relationship id (`r:id`) for this sheet, preserved for round-trip fidelity.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub xlsx_rel_id: Option<String>,

    /// Sheet visibility state.
    #[serde(default, skip_serializing_if = "is_visible")]
    pub visibility: SheetVisibility,

    /// Optional tab color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<TabColor>,
    /// Floating drawings (images, shapes, chart placeholders) anchored to the sheet.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub drawings: Vec<DrawingObject>,

    /// Sparse cell storage; only non-empty cells are stored.
    #[serde(default)]
    cells: HashMap<CellKey, Cell>,

    /// Bounding box of stored cells.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    used_range: Option<Range>,

    /// Merged-cell regions for this worksheet.
    ///
    /// Values are stored only in the top-left (anchor) cell. All cell addresses inside a
    /// merged region resolve to that anchor, matching Excel semantics.
    #[serde(default, skip_serializing_if = "MergedRegions::is_empty")]
    pub merged_regions: MergedRegions,

    /// Logical row count (may exceed the used range).
    #[serde(default = "default_row_count")]
    pub row_count: u32,

    /// Logical column count.
    #[serde(default = "default_col_count")]
    pub col_count: u32,

    /// Per-row formatting/visibility overrides.
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub row_properties: BTreeMap<u32, RowProperties>,

    /// Per-column formatting/visibility overrides.
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub col_properties: BTreeMap<u32, ColProperties>,

    /// Sheet default row height in points.
    ///
    /// This corresponds to OOXML `<sheetFormatPr defaultRowHeight="...">`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_row_height: Option<f32>,

    /// Sheet default column width in Excel "character" units.
    ///
    /// This corresponds to OOXML `<sheetFormatPr defaultColWidth="...">`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_col_width: Option<f32>,

    /// Base column width in characters.
    ///
    /// This corresponds to OOXML `<sheetFormatPr baseColWidth="...">`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub base_col_width: Option<u16>,

    /// Excel-style row/column outline (grouping) metadata.
    ///
    /// Indexes within the outline are 1-based (matching Excel / OOXML), whereas the sheet's
    /// cell grid and `row_properties`/`col_properties` are 0-based.
    #[serde(default, skip_serializing_if = "is_default_outline")]
    pub outline: Outline,

    /// Frozen pane row count (top).
    #[serde(default)]
    pub frozen_rows: u32,

    /// Frozen pane column count (left).
    #[serde(default)]
    pub frozen_cols: u32,

    /// Sheet zoom level (1.0 = 100%).
    #[serde(default = "default_zoom", skip_serializing_if = "is_default_zoom")]
    pub zoom: f32,

    /// Sheet view state (selection, pane splits, gridlines/headings visibility, zoom, etc).
    #[serde(default, skip_serializing_if = "SheetView::is_default")]
    pub view: SheetView,

    /// Excel tables (structured ranges) hosted on this worksheet.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tables: Vec<Table>,

    /// Worksheet-level AutoFilter (`<autoFilter>`), if present.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub auto_filter: Option<SheetAutoFilter>,

    /// Conditional formatting rules for this worksheet.
    #[serde(
        default,
        skip_serializing_if = "Vec::is_empty",
        alias = "conditional_formatting"
    )]
    pub conditional_formatting_rules: Vec<CfRule>,

    /// Differential formats referenced by conditional formatting rules.
    ///
    /// [`CfRule::dxf_id`] indexes into this vector.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub conditional_formatting_dxfs: Vec<CfStyleOverride>,

    /// Runtime cache for conditional formatting evaluation.
    #[serde(skip)]
    conditional_formatting_engine: RefCell<ConditionalFormattingEngine>,

    /// Optional columnar backing store for large imported datasets.
    ///
    /// This is runtime-only for now; persistence is handled by the storage layer.
    #[serde(skip)]
    columnar: Option<ColumnarBackend>,

    /// Hyperlinks anchored to cells/ranges in this worksheet.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub hyperlinks: Vec<Hyperlink>,

    /// Data validation rules assigned to ranges on this worksheet.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub data_validations: Vec<DataValidationAssignment>,

    /// Cell comments (legacy notes + modern threaded comments) anchored to this worksheet.
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    comments: BTreeMap<CellKey, Vec<Comment>>,

    /// Next data validation id to allocate (runtime-only).
    #[serde(skip)]
    next_data_validation_id: DataValidationId,
    /// Sheet protection options (Excel-compatible).
    #[serde(default, skip_serializing_if = "SheetProtection::is_default")]
    pub sheet_protection: SheetProtection,
}

impl Worksheet {
    /// Create a new empty worksheet.
    pub fn new(id: WorksheetId, name: impl Into<String>) -> Self {
        let frozen_rows = 0;
        let frozen_cols = 0;
        let zoom = default_zoom();
        let mut view = SheetView::default();
        view.pane.frozen_rows = frozen_rows;
        view.pane.frozen_cols = frozen_cols;
        view.zoom = zoom;
        Self {
            id,
            name: name.into(),
            xlsx_sheet_id: None,
            xlsx_rel_id: None,
            visibility: SheetVisibility::Visible,
            tab_color: None,
            drawings: Vec::new(),
            cells: HashMap::new(),
            used_range: None,
            merged_regions: MergedRegions::new(),
            row_count: default_row_count(),
            col_count: default_col_count(),
            row_properties: BTreeMap::new(),
            col_properties: BTreeMap::new(),
            default_row_height: None,
            default_col_width: None,
            base_col_width: None,
            outline: Outline::default(),
            frozen_rows,
            frozen_cols,
            zoom,
            view,
            tables: Vec::new(),
            auto_filter: None,
            conditional_formatting_rules: Vec::new(),
            conditional_formatting_dxfs: Vec::new(),
            conditional_formatting_engine: RefCell::new(ConditionalFormattingEngine::default()),
            columnar: None,
            hyperlinks: Vec::new(),
            data_validations: Vec::new(),
            comments: BTreeMap::new(),
            next_data_validation_id: 1,
            sheet_protection: SheetProtection::default(),
        }
    }

    /// Replace all conditional formatting rules and differential formats.
    ///
    /// This clears any cached evaluation results.
    pub fn set_conditional_formatting(&mut self, rules: Vec<CfRule>, dxfs: Vec<CfStyleOverride>) {
        self.conditional_formatting_rules = rules;
        self.conditional_formatting_dxfs = dxfs;
        self.clear_conditional_formatting_cache();
    }

    /// Append a conditional formatting rule.
    ///
    /// This clears any cached evaluation results.
    pub fn add_conditional_formatting_rule(&mut self, rule: CfRule) {
        self.conditional_formatting_rules.push(rule);
        self.clear_conditional_formatting_cache();
    }

    /// Remove all conditional formatting rules and differential formats.
    ///
    /// This clears any cached evaluation results.
    pub fn clear_conditional_formatting(&mut self) {
        self.conditional_formatting_rules.clear();
        self.conditional_formatting_dxfs.clear();
        self.clear_conditional_formatting_cache();
    }

    /// Invalidate cached conditional formatting results affected by the given changed cells.
    pub fn invalidate_conditional_formatting_cells<I: IntoIterator<Item = CellRef>>(
        &self,
        changed: I,
    ) {
        self.conditional_formatting_engine
            .borrow_mut()
            .invalidate_cells(changed);
    }

    /// Clear all cached conditional formatting evaluations.
    pub fn clear_conditional_formatting_cache(&self) {
        self.conditional_formatting_engine
            .borrow_mut()
            .clear_cache();
    }

    /// Evaluate conditional formatting rules for a visible viewport range.
    ///
    /// Results are cached per-visible-range and can be invalidated via
    /// [`Worksheet::invalidate_conditional_formatting_cells`].
    pub fn evaluate_conditional_formatting(
        &self,
        visible: Range,
        values: &dyn CellValueProvider,
        formula_evaluator: Option<&dyn FormulaEvaluator>,
    ) -> CfEvaluationResult {
        if self.conditional_formatting_rules.is_empty() {
            return CfEvaluationResult::new(visible);
        }

        let dxfs: Option<&dyn DifferentialFormatProvider> =
            if self.conditional_formatting_dxfs.is_empty() {
                None
            } else {
                Some(&self.conditional_formatting_dxfs)
            };

        self.conditional_formatting_engine
            .borrow_mut()
            .evaluate_visible_range(
                &self.conditional_formatting_rules,
                visible,
                values,
                formula_evaluator,
                dxfs,
            )
            .clone()
    }

    /// Returns whether a cell can be edited (i.e. its value/formula changed) given the current
    /// worksheet protection state.
    ///
    /// Note: this is a helper for UI/engine consumers; it does not enforce edits automatically.
    pub fn is_cell_editable(&self, cell_ref: CellRef, styles: &StyleTable) -> bool {
        if cell_ref.row >= self.row_count
            || cell_ref.col >= self.col_count
            || cell_ref.col >= crate::cell::EXCEL_MAX_COLS
        {
            return false;
        }

        if !self.sheet_protection.enabled {
            return true;
        }

        // Editing any cell in a merged region edits the anchor cell.
        let anchor = self.merged_regions.resolve_cell(cell_ref);
        let style_id = self
            .cells
            .get(&CellKey::from_ref(anchor))
            .map(|c| c.style_id)
            .unwrap_or(0);

        let style = styles.styles.get(style_id as usize);
        let locked = style
            .and_then(|s| s.protection.as_ref())
            .map(|p| p.locked)
            .unwrap_or(true);

        !locked
    }

    /// Returns whether an operation is allowed given the current worksheet protection state.
    ///
    /// Note: this is a helper for UI/engine consumers; it does not enforce edits automatically.
    pub fn can_perform(&self, action: SheetProtectionAction) -> bool {
        if !self.sheet_protection.enabled {
            return true;
        }

        match action {
            SheetProtectionAction::SelectLockedCells => self.sheet_protection.select_locked_cells,
            SheetProtectionAction::SelectUnlockedCells => {
                self.sheet_protection.select_unlocked_cells
            }
            SheetProtectionAction::FormatCells => self.sheet_protection.format_cells,
            SheetProtectionAction::FormatColumns => self.sheet_protection.format_columns,
            SheetProtectionAction::FormatRows => self.sheet_protection.format_rows,
            SheetProtectionAction::InsertColumns => self.sheet_protection.insert_columns,
            SheetProtectionAction::InsertRows => self.sheet_protection.insert_rows,
            SheetProtectionAction::InsertHyperlinks => self.sheet_protection.insert_hyperlinks,
            SheetProtectionAction::DeleteColumns => self.sheet_protection.delete_columns,
            SheetProtectionAction::DeleteRows => self.sheet_protection.delete_rows,
            SheetProtectionAction::Sort => self.sheet_protection.sort,
            SheetProtectionAction::AutoFilter => self.sheet_protection.auto_filter,
            SheetProtectionAction::PivotTables => self.sheet_protection.pivot_tables,
            SheetProtectionAction::EditObjects => self.sheet_protection.edit_objects,
            SheetProtectionAction::EditScenarios => self.sheet_protection.edit_scenarios,
        }
    }

    /// Assign a data validation rule to the given ranges.
    ///
    /// Returns the allocated data validation id.
    pub fn add_data_validation(
        &mut self,
        ranges: Vec<Range>,
        validation: DataValidation,
    ) -> DataValidationId {
        let id = self.next_data_validation_id;
        self.next_data_validation_id = self.next_data_validation_id.wrapping_add(1);
        self.data_validations.push(DataValidationAssignment {
            id,
            ranges,
            validation,
        });
        id
    }

    /// Remove a data validation rule by id.
    pub fn remove_data_validation(&mut self, id: DataValidationId) -> bool {
        let before = self.data_validations.len();
        self.data_validations.retain(|dv| dv.id != id);
        before != self.data_validations.len()
    }

    /// Return all data validation assignments that apply to `cell`.
    ///
    /// If `cell` is inside a merged region, validations applied to any part of the merged region
    /// are treated as applying to the anchor cell (Excel-like behavior).
    pub fn data_validations_for_cell(&self, cell: CellRef) -> Vec<&DataValidationAssignment> {
        if cell.row >= self.row_count
            || cell.col >= self.col_count
            || cell.col >= crate::cell::EXCEL_MAX_COLS
        {
            return Vec::new();
        }

        let anchor = self.merged_regions.resolve_cell(cell);
        let merged_range = self.merged_regions.containing_range(cell);

        let count = self
            .data_validations
            .iter()
            .filter(|assignment| {
                assignment.ranges.iter().any(|range| {
                    range.contains(anchor)
                        || merged_range.is_some_and(|merged| range.intersects(&merged))
                })
            })
            .count();
        let mut out: Vec<&DataValidationAssignment> = Vec::new();
        if out.try_reserve_exact(count).is_err() {
            debug_assert!(
                false,
                "allocation failed (data validations for cell, count={count})"
            );
            return Vec::new();
        }
        for assignment in self.data_validations.iter().filter(|assignment| {
            assignment.ranges.iter().any(|range| {
                range.contains(anchor) || merged_range.is_some_and(|merged| range.intersects(&merged))
            })
        }) {
            out.push(assignment);
        }
        out
    }

    /// Set the active cell and selection ranges.
    pub fn set_selection(&mut self, selection: SheetSelection) {
        self.view.selection = Some(selection);
    }

    /// Return the current selection state, if explicitly set.
    pub fn selection(&self) -> Option<&SheetSelection> {
        self.view.selection.as_ref()
    }

    /// Find the first table containing `cell`.
    pub fn table_for_cell(&self, cell: CellRef) -> Option<&Table> {
        self.tables.iter().find(|t| t.range.contains(cell))
    }

    /// Find a table by its workbook-scoped name (case-insensitive, like Excel).
    pub fn table_by_name_case_insensitive(&self, table_name: &str) -> Option<&Table> {
        self.tables.iter().find(|t| {
            t.name.eq_ignore_ascii_case(table_name)
                || t.display_name.eq_ignore_ascii_case(table_name)
        })
    }

    /// Find a table by id.
    pub fn table_by_id(&self, table_id: u32) -> Option<&Table> {
        self.tables.iter().find(|t| t.id == table_id)
    }

    /// Find a mutable table by its workbook-scoped name (case-insensitive, like Excel).
    pub fn table_mut_by_name_case_insensitive(&mut self, table_name: &str) -> Option<&mut Table> {
        self.tables.iter_mut().find(|t| {
            t.name.eq_ignore_ascii_case(table_name)
                || t.display_name.eq_ignore_ascii_case(table_name)
        })
    }

    /// Remove a table by name (case-insensitive).
    pub fn remove_table_by_name(&mut self, table_name: &str) -> Option<Table> {
        let idx = self.tables.iter().position(|t| {
            t.name.eq_ignore_ascii_case(table_name)
                || t.display_name.eq_ignore_ascii_case(table_name)
        })?;
        Some(self.tables.remove(idx))
    }

    /// Remove a table by id.
    pub fn remove_table_by_id(&mut self, table_id: u32) -> Option<Table> {
        let idx = self.tables.iter().position(|t| t.id == table_id)?;
        Some(self.tables.remove(idx))
    }

    /// Number of stored cells.
    ///
    /// This is proportional to memory usage for the sheet's cell content.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Get the current used range for *sparse stored cells*.
    ///
    /// This is the bounding box of the worksheet's in-memory cell map (edits, formulas,
    /// formatting overrides), and does **not** include any attached columnar backing table
    /// (see [`Worksheet::set_columnar_table`]).
    ///
    /// If the worksheet is backed by a columnar table and you want the effective extent of
    /// visible data, use [`Worksheet::effective_used_range`].
    pub fn used_range(&self) -> Option<Range> {
        self.used_range
    }

    /// Get the effective used range for this worksheet.
    ///
    /// This accounts for both:
    /// - sparse stored cells (see [`Worksheet::used_range`])
    /// - the optional columnar backing table (see [`Worksheet::columnar_range`])
    ///
    /// If both are present, the returned range is the bounding union of the two extents.
    pub fn effective_used_range(&self) -> Option<Range> {
        match (self.used_range(), self.columnar_range()) {
            (None, None) => None,
            (Some(r), None) | (None, Some(r)) => Some(r),
            (Some(a), Some(b)) => Some(a.bounding_box(&b)),
        }
    }

    /// Attach a columnar table as the backing store for this worksheet.
    ///
    /// The existing sparse cell map is retained as an "overlay" for edits, formulas,
    /// and styles.
    pub fn set_columnar_table(&mut self, origin: CellRef, table: Arc<ColumnarTable>) {
        let table_rows = table.row_count().min(u32::MAX as usize) as u32;
        let table_cols = table.column_count().min(u32::MAX as usize) as u32;
        self.row_count = self.row_count.max(origin.row.saturating_add(table_rows));
        self.col_count = self.col_count.max(origin.col.saturating_add(table_cols));
        self.columnar = Some(ColumnarBackend { origin, table });
        self.clear_conditional_formatting_cache();
    }

    /// Remove any columnar backing table.
    pub fn clear_columnar_table(&mut self) {
        self.columnar = None;
        self.clear_conditional_formatting_cache();
    }

    /// Returns the backing columnar table, if present.
    pub fn columnar_table(&self) -> Option<&Arc<ColumnarTable>> {
        self.columnar.as_ref().map(|c| &c.table)
    }

    /// Returns the origin cell for the columnar backing table, if present.
    pub fn columnar_origin(&self) -> Option<CellRef> {
        self.columnar.as_ref().map(|c| c.origin)
    }

    /// Returns the (origin, row_count, col_count) for the backing columnar table, clamped to
    /// the worksheet's dimensions.
    ///
    /// The returned `row_count` and `col_count` are suitable for iterating over cell positions
    /// within the table.
    pub fn columnar_table_extent(&self) -> Option<(CellRef, usize, usize)> {
        let backend = self.columnar.as_ref()?;
        let origin = backend.origin;
        if origin.row >= self.row_count
            || origin.col >= self.col_count
            || origin.col >= crate::cell::EXCEL_MAX_COLS
        {
            return None;
        }

        let max_rows = (self.row_count - origin.row) as usize;
        let max_cols =
            (self.col_count - origin.col).min(crate::cell::EXCEL_MAX_COLS - origin.col) as usize;
        let rows = backend.table.row_count().min(max_rows);
        let cols = backend.table.column_count().min(max_cols);
        if rows == 0 || cols == 0 {
            return None;
        }

        Some((origin, rows, cols))
    }

    /// Returns the bounding range covered by the backing columnar table, clamped to the
    /// worksheet's dimensions.
    pub fn columnar_range(&self) -> Option<Range> {
        let (origin, rows, cols) = self.columnar_table_extent()?;
        let end = CellRef::new(
            origin.row.saturating_add(rows.saturating_sub(1) as u32),
            origin.col.saturating_add(cols.saturating_sub(1) as u32),
        );
        Some(Range::new(origin, end))
    }

    /// Get per-row properties if an override exists.
    pub fn row_properties(&self, row: u32) -> Option<&RowProperties> {
        self.row_properties.get(&row)
    }

    /// Get per-column properties if an override exists.
    pub fn col_properties(&self, col: u32) -> Option<&ColProperties> {
        self.col_properties.get(&col)
    }

    /// Returns the outline entry for a row using 1-based Excel indexing.
    ///
    /// Note: `RowProperties.hidden` is treated as the persisted "user hidden" bit and is
    /// kept in sync with `OutlineEntry.hidden.user` when using [`Worksheet::set_row_hidden`].
    pub fn row_outline_entry(&self, row_1based: u32) -> OutlineEntry {
        assert!(
            row_1based >= 1,
            "row_1based must be >= 1 (got {row_1based})"
        );
        let mut entry = self.outline.rows.entry(row_1based);
        let row_0based = row_1based - 1;
        if self
            .row_properties
            .get(&row_0based)
            .map(|p| p.hidden)
            .unwrap_or(false)
        {
            entry.hidden.user = true;
        }
        entry
    }

    /// Returns the outline entry for a column using 1-based Excel indexing.
    ///
    /// Note: `ColProperties.hidden` is treated as the persisted "user hidden" bit and is
    /// kept in sync with `OutlineEntry.hidden.user` when using [`Worksheet::set_col_hidden`].
    pub fn col_outline_entry(&self, col_1based: u32) -> OutlineEntry {
        assert!(
            col_1based >= 1,
            "col_1based must be >= 1 (got {col_1based})"
        );
        let mut entry = self.outline.cols.entry(col_1based);
        let col_0based = col_1based - 1;
        if self
            .col_properties
            .get(&col_0based)
            .map(|p| p.hidden)
            .unwrap_or(false)
        {
            entry.hidden.user = true;
        }
        entry
    }

    /// Returns whether the row is effectively hidden under Excel semantics.
    ///
    /// Effective hidden combines:
    /// - user-hidden (`RowProperties.hidden` / `OutlineEntry.hidden.user`)
    /// - outline-hidden (`OutlineEntry.hidden.outline`)
    /// - filter-hidden (`OutlineEntry.hidden.filter`)
    pub fn is_row_hidden_effective(&self, row_1based: u32) -> bool {
        self.row_outline_entry(row_1based).hidden.is_hidden()
    }

    /// Returns whether the column is effectively hidden under Excel semantics.
    pub fn is_col_hidden_effective(&self, col_1based: u32) -> bool {
        self.col_outline_entry(col_1based).hidden.is_hidden()
    }

    /// Group a row range, increasing the outline level by 1 for each row.
    pub fn group_rows(&mut self, start_1based: u32, end_1based: u32) {
        assert!(
            start_1based >= 1 && end_1based >= 1,
            "row indexes must be 1-based (got {start_1based}..={end_1based})"
        );
        let max = start_1based.max(end_1based);
        self.row_count = self.row_count.max(max);
        self.outline.group_rows(start_1based, end_1based);
    }

    /// Ungroup a row range, decreasing the outline level by 1 for each row.
    pub fn ungroup_rows(&mut self, start_1based: u32, end_1based: u32) {
        assert!(
            start_1based >= 1 && end_1based >= 1,
            "row indexes must be 1-based (got {start_1based}..={end_1based})"
        );
        let max = start_1based.max(end_1based);
        self.row_count = self.row_count.max(max);
        self.outline.ungroup_rows(start_1based, end_1based);
    }

    /// Collapse or expand a row outline group.
    pub fn toggle_row_group(&mut self, summary_index_1based: u32) -> bool {
        assert!(
            summary_index_1based >= 1,
            "summary_index_1based must be >= 1 (got {summary_index_1based})"
        );
        self.row_count = self.row_count.max(summary_index_1based);
        self.outline.toggle_row_group(summary_index_1based)
    }

    /// Group a column range, increasing the outline level by 1 for each column.
    pub fn group_cols(&mut self, start_1based: u32, end_1based: u32) {
        assert!(
            start_1based >= 1 && end_1based >= 1,
            "col indexes must be 1-based (got {start_1based}..={end_1based})"
        );
        let max = start_1based.max(end_1based);
        assert!(
            max <= crate::cell::EXCEL_MAX_COLS,
            "col out of Excel bounds: {max}"
        );
        self.col_count = self.col_count.max(max);
        self.outline.group_cols(start_1based, end_1based);
    }

    /// Ungroup a column range, decreasing the outline level by 1 for each column.
    pub fn ungroup_cols(&mut self, start_1based: u32, end_1based: u32) {
        assert!(
            start_1based >= 1 && end_1based >= 1,
            "col indexes must be 1-based (got {start_1based}..={end_1based})"
        );
        let max = start_1based.max(end_1based);
        assert!(
            max <= crate::cell::EXCEL_MAX_COLS,
            "col out of Excel bounds: {max}"
        );
        self.col_count = self.col_count.max(max);
        self.outline.ungroup_cols(start_1based, end_1based);
    }

    /// Collapse or expand a column outline group.
    pub fn toggle_col_group(&mut self, summary_index_1based: u32) -> bool {
        assert!(
            summary_index_1based >= 1,
            "summary_index_1based must be >= 1 (got {summary_index_1based})"
        );
        assert!(
            summary_index_1based <= crate::cell::EXCEL_MAX_COLS,
            "col out of Excel bounds: {summary_index_1based}"
        );
        self.col_count = self.col_count.max(summary_index_1based);
        self.outline.toggle_col_group(summary_index_1based)
    }

    /// Sets whether the row at `row_1based` is hidden by an AutoFilter.
    pub fn set_filter_hidden_row(&mut self, row_1based: u32, hidden: bool) {
        assert!(
            row_1based >= 1,
            "row_1based must be >= 1 (got {row_1based})"
        );
        self.row_count = self.row_count.max(row_1based);
        self.outline.rows.set_filter_hidden(row_1based, hidden);
    }

    /// Clears filter-hidden flags within `[start_1based, end_1based]`.
    pub fn clear_filter_hidden_range(&mut self, start_1based: u32, end_1based: u32) {
        assert!(
            start_1based >= 1 && end_1based >= 1,
            "row indexes must be 1-based (got {start_1based}..={end_1based})"
        );
        self.outline
            .rows
            .clear_filter_hidden_range(start_1based, end_1based);
    }

    /// Set (or clear) the height override for a row.
    ///
    /// Passing `None` removes the height override. If the row has no overrides
    /// remaining, its entry is removed from the map.
    pub fn set_row_height(&mut self, row: u32, height: Option<f32>) {
        self.row_count = self.row_count.max(row.saturating_add(1));
        match self.row_properties.get_mut(&row) {
            Some(props) => {
                props.height = height;
                if props.height.is_none() && !props.hidden && props.style_id.is_none() {
                    self.row_properties.remove(&row);
                }
            }
            None => {
                if height.is_none() {
                    return;
                }
                self.row_properties.insert(
                    row,
                    RowProperties {
                        height,
                        hidden: false,
                        style_id: None,
                    },
                );
            }
        }
    }

    /// Set the hidden flag for a row.
    ///
    /// If the row ends up with no overrides (not hidden and no height), its entry
    /// is removed from the map.
    pub fn set_row_hidden(&mut self, row: u32, hidden: bool) {
        self.row_count = self.row_count.max(row.saturating_add(1));
        self.outline
            .rows
            .set_user_hidden(row.saturating_add(1), hidden);
        match self.row_properties.get_mut(&row) {
            Some(props) => {
                props.hidden = hidden;
                if props.height.is_none() && !props.hidden && props.style_id.is_none() {
                    self.row_properties.remove(&row);
                }
            }
            None => {
                if !hidden {
                    return;
                }
                self.row_properties.insert(
                    row,
                    RowProperties {
                        height: None,
                        hidden,
                        style_id: None,
                    },
                );
            }
        }
    }

    /// Set (or clear) the default style id for all cells in a row.
    ///
    /// Passing `None` (or `Some(0)`) removes the style override. If the row has no overrides
    /// remaining, its entry is removed from the map.
    pub fn set_row_style_id(&mut self, row_0based: u32, style_id: Option<u32>) {
        self.row_count = self.row_count.max(row_0based.saturating_add(1));
        let style_id = style_id.filter(|id| *id != 0);
        match self.row_properties.get_mut(&row_0based) {
            Some(props) => {
                props.style_id = style_id;
                if props.height.is_none() && !props.hidden && props.style_id.is_none() {
                    self.row_properties.remove(&row_0based);
                }
            }
            None => {
                let Some(style_id) = style_id else {
                    return;
                };
                self.row_properties.insert(
                    row_0based,
                    RowProperties {
                        height: None,
                        hidden: false,
                        style_id: Some(style_id),
                    },
                );
            }
        }
    }

    /// Set (or clear) the width override for a column.
    ///
    /// `width` is expressed in Excel "character" units (OOXML `col/@width`), **not pixels**.
    pub fn set_col_width(&mut self, col: u32, width: Option<f32>) {
        assert!(
            col < crate::cell::EXCEL_MAX_COLS,
            "col out of Excel bounds: {col}"
        );
        self.col_count = self.col_count.max(col.saturating_add(1));
        match self.col_properties.get_mut(&col) {
            Some(props) => {
                props.width = width;
                if props.width.is_none() && !props.hidden && props.style_id.is_none() {
                    self.col_properties.remove(&col);
                }
            }
            None => {
                if width.is_none() {
                    return;
                }
                self.col_properties.insert(
                    col,
                    ColProperties {
                        width,
                        hidden: false,
                        style_id: None,
                    },
                );
            }
        }
    }

    /// Set the hidden flag for a column.
    pub fn set_col_hidden(&mut self, col: u32, hidden: bool) {
        assert!(
            col < crate::cell::EXCEL_MAX_COLS,
            "col out of Excel bounds: {col}"
        );
        self.col_count = self.col_count.max(col.saturating_add(1));
        self.outline
            .cols
            .set_user_hidden(col.saturating_add(1), hidden);
        match self.col_properties.get_mut(&col) {
            Some(props) => {
                props.hidden = hidden;
                if props.width.is_none() && !props.hidden && props.style_id.is_none() {
                    self.col_properties.remove(&col);
                }
            }
            None => {
                if !hidden {
                    return;
                }
                self.col_properties.insert(
                    col,
                    ColProperties {
                        width: None,
                        hidden,
                        style_id: None,
                    },
                );
            }
        }
    }

    /// Set (or clear) the default style id for all cells in a column.
    ///
    /// Passing `None` (or `Some(0)`) removes the style override. If the column has no overrides
    /// remaining, its entry is removed from the map.
    pub fn set_col_style_id(&mut self, col_0based: u32, style_id: Option<u32>) {
        assert!(
            col_0based < crate::cell::EXCEL_MAX_COLS,
            "col out of Excel bounds: {col_0based}"
        );
        self.col_count = self.col_count.max(col_0based.saturating_add(1));
        let style_id = style_id.filter(|id| *id != 0);
        match self.col_properties.get_mut(&col_0based) {
            Some(props) => {
                props.style_id = style_id;
                if props.width.is_none() && !props.hidden && props.style_id.is_none() {
                    self.col_properties.remove(&col_0based);
                }
            }
            None => {
                let Some(style_id) = style_id else {
                    return;
                };
                self.col_properties.insert(
                    col_0based,
                    ColProperties {
                        width: None,
                        hidden: false,
                        style_id: Some(style_id),
                    },
                );
            }
        }
    }

    /// Get a cell record if it is present in the sparse store.
    pub fn cell(&self, cell: CellRef) -> Option<&Cell> {
        if cell.row >= self.row_count
            || cell.col >= self.col_count
            || cell.col >= crate::cell::EXCEL_MAX_COLS
        {
            return None;
        }
        let anchor = self.merged_regions.resolve_cell(cell);
        self.cells.get(&CellKey::from_ref(anchor))
    }

    /// Get a mutable cell record if it is present in the sparse store.
    pub fn cell_mut(&mut self, cell: CellRef) -> Option<&mut Cell> {
        if cell.row >= self.row_count
            || cell.col >= self.col_count
            || cell.col >= crate::cell::EXCEL_MAX_COLS
        {
            return None;
        }
        self.cells.get_mut(&CellKey::from_ref(cell))
    }

    /// Get a cell record from an A1 reference (e.g. `A1`, `$B$2`).
    pub fn cell_a1(&self, a1: &str) -> Result<Option<&Cell>, A1ParseError> {
        Ok(self.cell(CellRef::from_a1(a1)?))
    }

    /// Get a cell's value, returning [`CellValue::Empty`] if unset.
    pub fn value(&self, cell: CellRef) -> CellValue {
        if let Some(cell) = self.cell(cell) {
            return cell.value.clone();
        }

        self.columnar_value(cell).unwrap_or(CellValue::Empty)
    }

    fn columnar_value(&self, cell: CellRef) -> Option<CellValue> {
        let backend = self.columnar.as_ref()?;
        if cell.row < backend.origin.row || cell.col < backend.origin.col {
            return None;
        }

        let row = (cell.row - backend.origin.row) as usize;
        let col = (cell.col - backend.origin.col) as usize;
        if row >= backend.table.row_count() || col >= backend.table.column_count() {
            return None;
        }

        let col_type = backend.table.schema().get(col)?.column_type;
        let value = backend.table.get_cell(row, col);
        Some(columnar_to_cell_value(value, col_type))
    }

    /// Fetch a rectangular region as a column-major payload suitable for
    /// virtualized grid rendering.
    pub fn get_range_batch(&self, range: Range) -> RangeBatch {
        let mut buffer = RangeBatchBuffer::default();
        self.get_range_batch_into(range, &mut buffer);
        RangeBatch {
            start: range.start,
            columns: buffer.columns,
        }
    }

    /// Fetch a rectangular region into a reusable output buffer.
    ///
    /// This is the allocation-friendly variant of [`Worksheet::get_range_batch`]. Callers can
    /// retain a [`RangeBatchBuffer`] across frames/requests to avoid repeated allocation of
    /// nested `Vec<Vec<CellValue>>` buffers.
    ///
    /// The returned [`RangeBatchRef`] borrows from `out`.
    pub fn get_range_batch_into<'a>(
        &self,
        range: Range,
        out: &'a mut RangeBatchBuffer,
    ) -> RangeBatchRef<'a> {
        let rows = (range.end.row - range.start.row + 1) as usize;
        let cols = (range.end.col - range.start.col + 1) as usize;
        out.columns.resize_with(cols, Vec::new);
        for column in &mut out.columns {
            column.clear();
            column.resize(rows, CellValue::Empty);
        }

        // Bulk fill from columnar backing (if present).
        if let Some(backend) = &self.columnar {
            fill_from_columnar(&mut out.columns, range, backend);
        }

        // Overlay sparse cells (edits / formulas / formatting) on top.
        for r_off in 0..rows {
            let row = range.start.row + r_off as u32;
            for c_off in 0..cols {
                let col = range.start.col + c_off as u32;
                let cell_ref = CellRef::new(row, col);
                if let Some(cell) = self.cell(cell_ref) {
                    out.columns[c_off][r_off] = cell.value.clone();
                }
            }
        }

        RangeBatchRef {
            start: range.start,
            columns: &out.columns,
        }
    }

    /// Get a cell's value from an A1 reference, returning [`CellValue::Empty`] if unset.
    pub fn value_a1(&self, a1: &str) -> Result<CellValue, A1ParseError> {
        Ok(self.value(CellRef::from_a1(a1)?))
    }

    /// Get the formula text for a cell, if present.
    pub fn formula(&self, cell: CellRef) -> Option<&str> {
        self.cell(cell).and_then(|c| c.formula.as_deref())
    }

    /// Get the phonetic metadata for a cell, if present.
    pub fn phonetic(&self, cell: CellRef) -> Option<&str> {
        self.cell(cell).and_then(|c| c.phonetic.as_deref())
    }

    /// Get the formula text for a cell addressed by an A1 reference.
    pub fn formula_a1(&self, a1: &str) -> Result<Option<&str>, A1ParseError> {
        Ok(self.formula(CellRef::from_a1(a1)?))
    }

    /// Get the phonetic metadata for a cell addressed by an A1 reference.
    pub fn phonetic_a1(&self, a1: &str) -> Result<Option<&str>, A1ParseError> {
        Ok(self.phonetic(CellRef::from_a1(a1)?))
    }

    /// Set a cell formula.
    ///
    /// Setting a formula to `None` clears the formula and removes the cell from the sparse store
    /// if it becomes "truly empty".
    pub fn set_formula(&mut self, cell_ref: CellRef, formula: Option<String>) {
        let formula = formula.and_then(|formula| crate::normalize_formula_text(&formula));

        let anchor = self.merged_regions.resolve_cell(cell_ref);
        let key = CellKey::from(anchor);

        match self.cells.get_mut(&key) {
            Some(cell) => {
                cell.formula = formula;
                if cell.is_truly_empty() {
                    self.cells.remove(&key);
                    self.on_cell_removed(anchor);
                }
            }
            None => {
                let Some(formula) = formula else { return };
                let mut cell = Cell::default();
                cell.formula = Some(formula);
                self.cells.insert(key, cell);
                self.on_cell_inserted(anchor);
            }
        }

        self.invalidate_conditional_formatting_cells([anchor]);
    }

    /// Set a cell's phonetic metadata.
    ///
    /// Phonetic metadata is treated as observable content: setting phonetic on an empty cell
    /// creates a stored cell record, while clearing phonetic removes the cell record if it
    /// becomes "truly empty".
    pub fn set_phonetic(&mut self, cell_ref: CellRef, phonetic: Option<String>) {
        let anchor = self.merged_regions.resolve_cell(cell_ref);
        let key = CellKey::from(anchor);

        match self.cells.get_mut(&key) {
            Some(cell) => {
                cell.phonetic = phonetic;
                if cell.is_truly_empty() {
                    self.cells.remove(&key);
                    self.on_cell_removed(anchor);
                }
            }
            None => {
                let Some(phonetic) = phonetic else { return };
                let mut cell = Cell::default();
                cell.phonetic = Some(phonetic);
                self.cells.insert(key, cell);
                self.on_cell_inserted(anchor);
            }
        }
    }

    /// Set a cell formula using an A1 reference.
    pub fn set_formula_a1(
        &mut self,
        a1: &str,
        formula: Option<String>,
    ) -> Result<(), A1ParseError> {
        let cell_ref = CellRef::from_a1(a1)?;
        self.set_formula(cell_ref, formula);
        Ok(())
    }

    /// Set a cell's phonetic metadata using an A1 reference.
    pub fn set_phonetic_a1(
        &mut self,
        a1: &str,
        phonetic: Option<String>,
    ) -> Result<(), A1ParseError> {
        let cell_ref = CellRef::from_a1(a1)?;
        self.set_phonetic(cell_ref, phonetic);
        Ok(())
    }

    /// Set the cell's style id.
    ///
    /// Cells with a non-zero style id are stored even if the value is empty,
    /// matching Excel's ability to format empty cells.
    pub fn set_style_id(&mut self, cell_ref: CellRef, style_id: u32) {
        let key = CellKey::from(cell_ref);

        match self.cells.get_mut(&key) {
            Some(cell) => {
                cell.style_id = style_id;
                if cell.is_truly_empty() {
                    self.cells.remove(&key);
                    self.on_cell_removed(cell_ref);
                }
            }
            None => {
                if style_id == 0 {
                    return;
                }
                let mut cell = Cell::default();
                cell.style_id = style_id;
                self.cells.insert(key, cell);
                self.on_cell_inserted(cell_ref);
            }
        }

        self.invalidate_conditional_formatting_cells([cell_ref]);
    }

    /// Set the cell's style id using an A1 reference.
    pub fn set_style_id_a1(&mut self, a1: &str, style_id: u32) -> Result<(), A1ParseError> {
        let cell_ref = CellRef::from_a1(a1)?;
        self.set_style_id(cell_ref, style_id);
        Ok(())
    }

    /// Set or replace a cell record.
    ///
    /// If the cell becomes "truly empty", it is removed from storage.
    pub fn set_cell(&mut self, cell_ref: CellRef, mut cell: Cell) {
        if let Some(formula) = cell.formula.take() {
            cell.formula = crate::normalize_formula_text(&formula);
        }

        let anchor = self.merged_regions.resolve_cell(cell_ref);
        let key = CellKey::from(anchor);

        if cell.is_truly_empty() {
            let removed = self.cells.remove(&key).is_some();
            if removed {
                self.on_cell_removed(anchor);
                self.invalidate_conditional_formatting_cells([anchor]);
            }
            return;
        }

        let is_new = self.cells.insert(key, cell).is_none();
        if is_new {
            self.on_cell_inserted(anchor);
        } else {
            // Existing cell updated; used range can only expand if the sheet was empty.
            if self.used_range.is_none() {
                self.used_range = Some(Range::new(anchor, anchor));
            }
        }

        self.invalidate_conditional_formatting_cells([anchor]);
    }

    /// Convenience: set only the value for a cell.
    ///
    /// If the target cell does not exist yet, it is created with default style.
    pub fn set_value(&mut self, cell_ref: CellRef, value: CellValue) {
        let anchor = self.merged_regions.resolve_cell(cell_ref);
        let key = CellKey::from(anchor);

        match self.cells.get_mut(&key) {
            Some(cell) => {
                cell.value = value;
                if cell.is_truly_empty() {
                    self.cells.remove(&key);
                    self.on_cell_removed(anchor);
                }
            }
            None => {
                if value == CellValue::Empty {
                    return;
                }
                self.cells.insert(key, Cell::new(value));
                self.on_cell_inserted(anchor);
            }
        }

        self.invalidate_conditional_formatting_cells([anchor]);
    }

    /// Set a cell value using an A1 reference (e.g. `C3`).
    pub fn set_value_a1(&mut self, a1: &str, value: CellValue) -> Result<(), A1ParseError> {
        let cell_ref = CellRef::from_a1(a1)?;
        self.set_value(cell_ref, value);
        Ok(())
    }

    /// Remove any stored record for this cell.
    pub fn clear_cell(&mut self, cell_ref: CellRef) {
        let anchor = self.merged_regions.resolve_cell(cell_ref);
        let key = CellKey::from(anchor);
        if self.cells.remove(&key).is_some() {
            self.on_cell_removed(anchor);
            self.invalidate_conditional_formatting_cells([anchor]);
        }
    }

    /// Remove any stored record for the cell addressed by an A1 reference.
    pub fn clear_cell_a1(&mut self, a1: &str) -> Result<(), A1ParseError> {
        let cell_ref = CellRef::from_a1(a1)?;
        self.clear_cell(cell_ref);
        Ok(())
    }

    /// Merge the given range.
    ///
    /// If the range intersects existing merged regions, they are unmerged first (Excel-like).
    /// When merging, only the top-left cell's value/formula/style is kept; all other stored
    /// cells in the range are cleared.
    pub fn merge_range(&mut self, range: Range) -> Result<(), MergeError> {
        if range.is_single_cell() {
            return Ok(());
        }

        self.merged_regions.unmerge_range(range);

        let anchor = range.start;
        let mut removed_any = false;
        let mut moved_comments: Vec<Comment> = Vec::new();
        for row in range.start.row..=range.end.row {
            for col in range.start.col..=range.end.col {
                let cell = CellRef::new(row, col);
                if cell != anchor {
                    removed_any |= self.cells.remove(&CellKey::from(cell)).is_some();
                    if let Some(mut comments) = self.comments.remove(&CellKey::from(cell)) {
                        for comment in &mut comments {
                            comment.cell_ref = anchor;
                        }
                        moved_comments.append(&mut comments);
                    }
                }
            }
        }

        if removed_any {
            self.recompute_used_range();
        }

        if !moved_comments.is_empty() {
            self.comments
                .entry(CellKey::from(anchor))
                .or_default()
                .extend(moved_comments);
        }

        let res = self.merged_regions.add(range);
        self.clear_conditional_formatting_cache();
        res
    }

    /// Unmerge any merged regions that intersect `range`.
    pub fn unmerge_range(&mut self, range: Range) -> usize {
        let count = self.merged_regions.unmerge_range(range);
        if count > 0 {
            self.clear_conditional_formatting_cache();
        }
        count
    }

    /// Iterate over all stored cells.
    pub fn iter_cells(&self) -> impl Iterator<Item = (CellRef, &Cell)> {
        self.cells.iter().map(|(k, v)| (k.to_ref(), v))
    }

    /// Iterate over all stored cells that fall within `range`.
    ///
    /// This is O(n) in the number of stored cells.
    pub fn iter_cells_in_range(&self, range: Range) -> impl Iterator<Item = (CellRef, &Cell)> {
        self.cells.iter().filter_map(move |(k, v)| {
            let cell_ref = k.to_ref();
            range.contains(cell_ref).then_some((cell_ref, v))
        })
    }

    /// Clear all stored cell records that fall within `range`.
    ///
    /// This recomputes the sheet's used range once after removals.
    pub fn clear_range(&mut self, range: Range) {
        let before = self.cells.len();
        self.cells.retain(|k, _| !range.contains(k.to_ref()));
        if before == self.cells.len() {
            return;
        }

        self.used_range = compute_used_range(&self.cells);
        self.clear_conditional_formatting_cache();
    }

    /// Iterate over all stored cells mutably.
    pub fn iter_cells_mut(&mut self) -> impl Iterator<Item = (CellRef, &mut Cell)> {
        self.cells.iter_mut().map(|(k, v)| (k.to_ref(), v))
    }

    /// Return the first hyperlink whose anchor range contains `cell`.
    pub fn hyperlink_at(&self, cell: CellRef) -> Option<&Hyperlink> {
        self.hyperlinks.iter().find(|h| h.range.contains(cell))
    }

    /// Normalize a cell reference for comment anchoring.
    ///
    /// Excel treats merged cells as a single cell anchored at the region's top-left.
    /// Comments inside a merged region are anchored to that top-left cell.
    pub fn normalize_comment_anchor(&self, cell_ref: CellRef) -> CellRef {
        self.merged_regions.resolve_cell(cell_ref)
    }

    /// Get all comments anchored to `cell_ref`.
    ///
    /// If `cell_ref` lies inside a merged region, this returns the comments anchored to the
    /// merged region's top-left cell.
    pub fn comments_for_cell(&self, cell_ref: CellRef) -> &[Comment] {
        let anchor = self.normalize_comment_anchor(cell_ref);
        let key = CellKey::from(anchor);
        self.comments.get(&key).map(Vec::as_slice).unwrap_or(&[])
    }

    /// Iterate over all comments in this worksheet, in deterministic order.
    ///
    /// Ordering is row-major by anchor cell, then insertion order within each cell.
    pub fn iter_comments(&self) -> impl Iterator<Item = (CellRef, &Comment)> {
        self.comments
            .iter()
            .flat_map(|(k, comments)| comments.iter().map(move |c| (k.to_ref(), c)))
    }

    fn comment_id_exists(&self, comment_id: &str) -> bool {
        self.iter_comments().any(|(_, c)| c.id == comment_id)
    }

    fn reply_id_exists(&self, reply_id: &str) -> bool {
        self.iter_comments()
            .flat_map(|(_, c)| c.replies.iter())
            .any(|r| r.id == reply_id)
    }

    fn normalize_comment_storage(&mut self) -> Result<(), CommentError> {
        if self.comments.is_empty() {
            return Ok(());
        }

        let existing = std::mem::take(&mut self.comments);
        let mut normalized: BTreeMap<CellKey, Vec<Comment>> = BTreeMap::new();
        let mut comment_ids: HashSet<String> = HashSet::new();
        let mut reply_ids: HashSet<String> = HashSet::new();

        for (key, comments) in existing {
            let anchor = self.normalize_comment_anchor(key.to_ref());
            let anchor_key = CellKey::from(anchor);

            for mut comment in comments {
                if comment.id.is_empty() {
                    comment.id = crate::new_uuid().to_string();
                }
                if !comment_ids.insert(comment.id.clone()) {
                    return Err(CommentError::DuplicateCommentId(comment.id));
                }

                comment.cell_ref = anchor;

                let mut local_reply_ids: HashSet<String> = HashSet::new();
                for reply in &mut comment.replies {
                    if reply.id.is_empty() {
                        reply.id = crate::new_uuid().to_string();
                    }
                    if !local_reply_ids.insert(reply.id.clone())
                        || !reply_ids.insert(reply.id.clone())
                    {
                        return Err(CommentError::DuplicateReplyId(reply.id.clone()));
                    }
                }

                normalized.entry(anchor_key).or_default().push(comment);
            }
        }

        self.comments = normalized;
        Ok(())
    }

    /// Add a comment (note or threaded) anchored to `cell_ref`.
    ///
    /// If `cell_ref` lies inside a merged region, the comment is anchored to the merged region's
    /// top-left cell. If `comment.id` is empty, an id is generated.
    pub fn add_comment(
        &mut self,
        cell_ref: CellRef,
        mut comment: Comment,
    ) -> Result<String, CommentError> {
        let anchor = self.normalize_comment_anchor(cell_ref);
        let key = CellKey::from(anchor);

        if comment.id.is_empty() {
            comment.id = crate::new_uuid().to_string();
        }
        if self.comment_id_exists(&comment.id) {
            return Err(CommentError::DuplicateCommentId(comment.id));
        }

        comment.cell_ref = anchor;

        // Ensure reply ids are unique (and usable for global lookup APIs).
        let mut reply_ids = HashSet::new();
        for reply in &mut comment.replies {
            if reply.id.is_empty() {
                reply.id = crate::new_uuid().to_string();
            }
            if !reply_ids.insert(reply.id.clone()) || self.reply_id_exists(&reply.id) {
                return Err(CommentError::DuplicateReplyId(reply.id.clone()));
            }
        }

        let id = comment.id.clone();
        self.comments.entry(key).or_default().push(comment);
        Ok(id)
    }

    /// Apply a partial update to an existing comment.
    pub fn update_comment(
        &mut self,
        comment_id: &str,
        patch: CommentPatch,
    ) -> Result<(), CommentError> {
        for comments in self.comments.values_mut() {
            for comment in comments.iter_mut() {
                if comment.id != comment_id {
                    continue;
                }

                if let Some(author) = patch.author {
                    comment.author = author;
                }
                if let Some(updated_at) = patch.updated_at {
                    comment.updated_at = updated_at;
                }
                if let Some(resolved) = patch.resolved {
                    comment.resolved = resolved;
                }
                if let Some(kind) = patch.kind {
                    comment.kind = kind;
                }
                if let Some(content) = patch.content {
                    comment.content = content;
                }
                if let Some(mentions) = patch.mentions {
                    comment.mentions = mentions;
                }

                return Ok(());
            }
        }

        Err(CommentError::CommentNotFound(comment_id.to_string()))
    }

    /// Delete a comment by id, returning the removed comment.
    pub fn delete_comment(&mut self, comment_id: &str) -> Result<Comment, CommentError> {
        let mut target: Option<(CellKey, usize)> = None;
        for (key, comments) in &self.comments {
            if let Some(idx) = comments.iter().position(|c| c.id == comment_id) {
                target = Some((*key, idx));
                break;
            }
        }

        let Some((key, idx)) = target else {
            return Err(CommentError::CommentNotFound(comment_id.to_string()));
        };

        let Some(comments) = self.comments.get_mut(&key) else {
            debug_assert!(
                false,
                "comment key disappeared while deleting comment: {key:?}"
            );
            return Err(CommentError::CommentNotFound(comment_id.to_string()));
        };
        let removed = comments.remove(idx);
        if comments.is_empty() {
            self.comments.remove(&key);
        }
        Ok(removed)
    }

    /// Add a reply to an existing comment.
    ///
    /// If `reply.id` is empty, an id is generated.
    pub fn add_reply(
        &mut self,
        comment_id: &str,
        mut reply: Reply,
    ) -> Result<String, CommentError> {
        if reply.id.is_empty() {
            reply.id = crate::new_uuid().to_string();
        }
        if self.reply_id_exists(&reply.id) {
            return Err(CommentError::DuplicateReplyId(reply.id));
        }

        for comments in self.comments.values_mut() {
            for comment in comments.iter_mut() {
                if comment.id != comment_id {
                    continue;
                }
                if comment.replies.iter().any(|r| r.id == reply.id) {
                    return Err(CommentError::DuplicateReplyId(reply.id));
                }
                let reply_id = reply.id.clone();
                comment.replies.push(reply);
                return Ok(reply_id);
            }
        }

        Err(CommentError::CommentNotFound(comment_id.to_string()))
    }

    /// Delete a reply by id.
    pub fn delete_reply(&mut self, reply_id: &str) -> Result<Reply, CommentError> {
        for comments in self.comments.values_mut() {
            for comment in comments.iter_mut() {
                if let Some(idx) = comment.replies.iter().position(|r| r.id == reply_id) {
                    return Ok(comment.replies.remove(idx));
                }
            }
        }

        Err(CommentError::ReplyNotFound(reply_id.to_string()))
    }

    fn on_cell_inserted(&mut self, cell_ref: CellRef) {
        self.row_count = self.row_count.max(cell_ref.row.saturating_add(1));
        self.col_count = self.col_count.max(cell_ref.col.saturating_add(1));

        match self.used_range {
            None => self.used_range = Some(Range::new(cell_ref, cell_ref)),
            Some(r) => {
                let start =
                    CellRef::new(r.start.row.min(cell_ref.row), r.start.col.min(cell_ref.col));
                let end = CellRef::new(r.end.row.max(cell_ref.row), r.end.col.max(cell_ref.col));
                self.used_range = Some(Range::new(start, end));
            }
        }
    }

    fn on_cell_removed(&mut self, cell_ref: CellRef) {
        let Some(r) = self.used_range else {
            return;
        };
        if self.cells.is_empty() {
            self.used_range = None;
            return;
        }

        // Only recompute if we removed a boundary cell.
        if cell_ref.row == r.start.row
            || cell_ref.row == r.end.row
            || cell_ref.col == r.start.col
            || cell_ref.col == r.end.col
        {
            self.recompute_used_range();
        }
    }

    fn recompute_used_range(&mut self) {
        self.used_range = compute_used_range(&self.cells);
    }
}

/// Column-major range payload for a virtualized grid viewport.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct RangeBatch {
    pub start: CellRef,
    pub columns: Vec<Vec<CellValue>>,
}

/// Reusable output buffer for [`Worksheet::get_range_batch_into`].
#[derive(Clone, Debug, Default)]
pub struct RangeBatchBuffer {
    pub columns: Vec<Vec<CellValue>>,
}

/// Borrowed view into a range batch produced by [`Worksheet::get_range_batch_into`].
#[derive(Copy, Clone, Debug)]
pub struct RangeBatchRef<'a> {
    pub start: CellRef,
    pub columns: &'a [Vec<CellValue>],
}

impl<'a> RangeBatchRef<'a> {
    /// Convert this borrowed range payload into an owned [`RangeBatch`].
    pub fn to_owned(self) -> RangeBatch {
        RangeBatch {
            start: self.start,
            columns: self.columns.to_vec(),
        }
    }
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

fn fill_from_columnar(dest: &mut [Vec<CellValue>], range: Range, backend: &ColumnarBackend) {
    let origin_row = backend.origin.row as u64;
    let origin_col = backend.origin.col as u64;
    let table_rows = backend.table.row_count() as u64;
    let table_cols = backend.table.column_count() as u64;

    if table_rows == 0 || table_cols == 0 {
        return;
    }

    let table_row_end = origin_row.saturating_add(table_rows - 1);
    let table_col_end = origin_col.saturating_add(table_cols - 1);

    let view_row_start = range.start.row as u64;
    let view_row_end = range.end.row as u64;
    let view_col_start = range.start.col as u64;
    let view_col_end = range.end.col as u64;

    let overlap_row_start = view_row_start.max(origin_row);
    let overlap_row_end = view_row_end.min(table_row_end);
    let overlap_col_start = view_col_start.max(origin_col);
    let overlap_col_end = view_col_end.min(table_col_end);

    if overlap_row_start > overlap_row_end || overlap_col_start > overlap_col_end {
        return;
    }

    let rel_row_start = (overlap_row_start - origin_row) as usize;
    let rel_row_end = (overlap_row_end - origin_row + 1) as usize;
    let rel_col_start = (overlap_col_start - origin_col) as usize;
    let rel_col_end = (overlap_col_end - origin_col + 1) as usize;

    let fetched = backend
        .table
        .get_range(rel_row_start, rel_row_end, rel_col_start, rel_col_end);

    let dest_row_off = (overlap_row_start - view_row_start) as usize;
    let dest_col_off = (overlap_col_start - view_col_start) as usize;

    let fetched_col_start = fetched.col_start;
    for (local_col, values) in fetched.columns.into_iter().enumerate() {
        let table_col_idx = fetched_col_start + local_col;
        let column_type = backend
            .table
            .schema()
            .get(table_col_idx)
            .map(|c| c.column_type)
            .unwrap_or(ColumnarType::String);

        for (local_row, v) in values.into_iter().enumerate() {
            let out_col = dest_col_off + local_col;
            let out_row = dest_row_off + local_row;
            if let Some(col_vec) = dest.get_mut(out_col) {
                if let Some(cell) = col_vec.get_mut(out_row) {
                    *cell = columnar_to_cell_value(v, column_type);
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn set_phonetic_creates_and_removes_records() {
        let mut sheet = Worksheet::new(1, "Sheet1");
        let cell_ref = CellRef::new(0, 0);

        assert!(sheet.cell(cell_ref).is_none());

        sheet.set_phonetic(cell_ref, Some("PHO".to_string()));
        assert_eq!(sheet.phonetic(cell_ref), Some("PHO"));
        assert!(
            sheet.cell(cell_ref).is_some(),
            "phonetic should create a cell record"
        );

        sheet.set_phonetic(cell_ref, None);
        assert!(
            sheet.cell(cell_ref).is_none(),
            "clearing phonetic on an otherwise-empty cell should remove the record"
        );

        // Clearing phonetic should not remove non-empty cells.
        sheet.set_value(cell_ref, CellValue::Number(1.0));
        sheet.set_phonetic(cell_ref, Some("PHO".to_string()));
        sheet.set_phonetic(cell_ref, None);
        assert!(sheet.cell(cell_ref).is_some());
        assert_eq!(sheet.value(cell_ref), CellValue::Number(1.0));
    }
}

impl<'de> Deserialize<'de> for Worksheet {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        #[derive(Deserialize)]
        struct Helper {
            id: WorksheetId,
            name: String,
            #[serde(default)]
            xlsx_sheet_id: Option<u32>,
            #[serde(default)]
            xlsx_rel_id: Option<String>,
            #[serde(default)]
            visibility: SheetVisibility,
            #[serde(default)]
            tab_color: Option<TabColor>,
            #[serde(default)]
            drawings: Vec<DrawingObject>,
            #[serde(default)]
            cells: HashMap<CellKey, Cell>,
            #[serde(default)]
            tables: Vec<Table>,
            #[serde(default)]
            auto_filter: Option<SheetAutoFilter>,
            #[serde(default, alias = "conditional_formatting")]
            conditional_formatting_rules: Vec<CfRule>,
            #[serde(default)]
            conditional_formatting_dxfs: Vec<CfStyleOverride>,
            #[serde(default = "default_row_count")]
            row_count: u32,
            #[serde(default = "default_col_count")]
            col_count: u32,
            #[serde(default)]
            merged_regions: MergedRegions,
            #[serde(default)]
            row_properties: BTreeMap<u32, RowProperties>,
            #[serde(default)]
            col_properties: BTreeMap<u32, ColProperties>,
            #[serde(default)]
            default_row_height: Option<f32>,
            #[serde(default)]
            default_col_width: Option<f32>,
            #[serde(default)]
            base_col_width: Option<u16>,
            #[serde(default)]
            outline: Outline,
            #[serde(default)]
            frozen_rows: u32,
            #[serde(default)]
            frozen_cols: u32,
            #[serde(default = "default_zoom")]
            zoom: f32,
            #[serde(default)]
            view: Option<SheetView>,
            #[serde(default)]
            hyperlinks: Vec<Hyperlink>,
            #[serde(default)]
            data_validations: Vec<DataValidationAssignment>,
            #[serde(default)]
            comments: BTreeMap<CellKey, Vec<Comment>>,
            #[serde(default)]
            sheet_protection: SheetProtection,
        }

        let helper = Helper::deserialize(deserializer)?;
        let used_range = compute_used_range(&helper.cells);

        if helper.row_count == 0 {
            return Err(D::Error::custom("row_count must be >= 1"));
        }
        if helper.col_count == 0 {
            return Err(D::Error::custom("col_count must be >= 1"));
        }
        if helper.col_count > crate::cell::EXCEL_MAX_COLS {
            return Err(D::Error::custom(format!(
                "col_count out of Excel bounds: {}",
                helper.col_count
            )));
        }

        for col in helper.col_properties.keys() {
            if *col >= crate::cell::EXCEL_MAX_COLS {
                return Err(D::Error::custom(format!(
                    "col_properties col out of Excel bounds: {col}"
                )));
            }
        }

        let mut row_count = helper.row_count;
        let mut col_count = helper.col_count;

        if let Some(used) = used_range {
            row_count = row_count.max(used.end.row.saturating_add(1));
            col_count = col_count.max(used.end.col.saturating_add(1));
        }

        if let Some(max_row) = helper.row_properties.keys().max().copied() {
            row_count = row_count.max(max_row.saturating_add(1));
        }
        if let Some(max_col) = helper.col_properties.keys().max().copied() {
            col_count = col_count.max(max_col.saturating_add(1));
        }

        if let Some(max_row_1based) = helper.outline.rows.iter().map(|(k, _)| k).max() {
            row_count = row_count.max(max_row_1based);
        }
        if let Some(max_col_1based) = helper.outline.cols.iter().map(|(k, _)| k).max() {
            col_count = col_count.max(max_col_1based);
        }
        if col_count > crate::cell::EXCEL_MAX_COLS {
            return Err(D::Error::custom(format!(
                "col_count out of Excel bounds: {col_count}"
            )));
        }

        let view_provided = helper.view.is_some();
        let mut view = helper.view.unwrap_or_default();
        let (frozen_rows, frozen_cols, zoom) = if view_provided {
            (view.pane.frozen_rows, view.pane.frozen_cols, view.zoom)
        } else {
            view.pane.frozen_rows = helper.frozen_rows;
            view.pane.frozen_cols = helper.frozen_cols;
            view.zoom = helper.zoom;
            (helper.frozen_rows, helper.frozen_cols, helper.zoom)
        };

        if frozen_rows > row_count {
            return Err(D::Error::custom(format!(
                "frozen_rows exceeds row_count: {} > {row_count}",
                frozen_rows
            )));
        }
        if frozen_cols > col_count {
            return Err(D::Error::custom(format!(
                "frozen_cols exceeds col_count: {} > {col_count}",
                frozen_cols
            )));
        }

        if zoom <= 0.0 {
            return Err(D::Error::custom(format!("zoom must be > 0 (got {})", zoom)));
        }

        let next_data_validation_id = helper
            .data_validations
            .iter()
            .map(|dv| dv.id)
            .max()
            .unwrap_or(0)
            .wrapping_add(1);

        let mut outline = helper.outline;
        outline.recompute_outline_hidden_rows();
        outline.recompute_outline_hidden_cols();

        let mut sheet = Worksheet {
            id: helper.id,
            name: helper.name,
            xlsx_sheet_id: helper.xlsx_sheet_id,
            xlsx_rel_id: helper.xlsx_rel_id,
            visibility: helper.visibility,
            tab_color: helper.tab_color,
            drawings: helper.drawings,
            cells: helper.cells,
            used_range,
            merged_regions: helper.merged_regions,
            row_count,
            col_count,
            row_properties: helper.row_properties,
            col_properties: helper.col_properties,
            default_row_height: helper.default_row_height,
            default_col_width: helper.default_col_width,
            base_col_width: helper.base_col_width,
            outline,
            frozen_rows,
            frozen_cols,
            zoom,
            view,
            tables: helper.tables,
            auto_filter: helper.auto_filter,
            conditional_formatting_rules: helper.conditional_formatting_rules,
            conditional_formatting_dxfs: helper.conditional_formatting_dxfs,
            conditional_formatting_engine: RefCell::new(ConditionalFormattingEngine::default()),
            columnar: None,
            hyperlinks: helper.hyperlinks,
            data_validations: helper.data_validations,
            comments: helper.comments,
            next_data_validation_id,
            sheet_protection: helper.sheet_protection,
        };

        sheet
            .normalize_comment_storage()
            .map_err(D::Error::custom)?;
        sheet.sync_user_hidden_bits();
        Ok(sheet)
    }
}

impl Worksheet {
    fn sync_user_hidden_bits(&mut self) {
        // RowProperties/ColProperties carry the persisted "user hidden" bit; keep the
        // outline's user-hidden state in sync for Excel visibility semantics.
        for (&row, props) in &self.row_properties {
            if props.hidden {
                self.outline
                    .rows
                    .set_user_hidden(row.saturating_add(1), true);
            }
        }
        for (&col, props) in &self.col_properties {
            if props.hidden {
                self.outline
                    .cols
                    .set_user_hidden(col.saturating_add(1), true);
            }
        }

        // Also sync in the other direction to support payloads that only store
        // `OutlineEntry.hidden.user` (e.g. XLSX round-trip) without corresponding
        // row/col properties.
        let user_hidden_row_count = self
            .outline
            .rows
            .iter()
            .filter(|(_, entry)| entry.hidden.user)
            .count();
        let mut user_hidden_rows: Vec<u32> = Vec::new();
        if user_hidden_rows
            .try_reserve_exact(user_hidden_row_count)
            .is_err()
        {
            debug_assert!(
                false,
                "allocation failed (sync hidden rows, count={user_hidden_row_count})"
            );
            return;
        }
        for (row_1based, entry) in self.outline.rows.iter() {
            if entry.hidden.user {
                user_hidden_rows.push(row_1based);
            }
        }
        for row_1based in user_hidden_rows {
            if row_1based == 0 {
                continue;
            }
            let row_0based = row_1based - 1;
            self.row_count = self.row_count.max(row_0based.saturating_add(1));
            self.row_properties
                .entry(row_0based)
                .and_modify(|p| p.hidden = true)
                .or_insert_with(|| RowProperties {
                    height: None,
                    hidden: true,
                    style_id: None,
                });
        }

        let user_hidden_col_count = self
            .outline
            .cols
            .iter()
            .filter(|(_, entry)| entry.hidden.user)
            .count();
        let mut user_hidden_cols: Vec<u32> = Vec::new();
        if user_hidden_cols
            .try_reserve_exact(user_hidden_col_count)
            .is_err()
        {
            debug_assert!(
                false,
                "allocation failed (sync hidden cols, count={user_hidden_col_count})"
            );
            return;
        }
        for (col_1based, entry) in self.outline.cols.iter() {
            if entry.hidden.user {
                user_hidden_cols.push(col_1based);
            }
        }
        for col_1based in user_hidden_cols {
            if col_1based == 0 {
                continue;
            }
            let col_0based = col_1based - 1;
            self.col_count = self.col_count.max(col_0based.saturating_add(1));
            self.col_properties
                .entry(col_0based)
                .and_modify(|p| p.hidden = true)
                .or_insert_with(|| ColProperties {
                    width: None,
                    hidden: true,
                    style_id: None,
                });
        }
    }
}

fn compute_used_range(cells: &HashMap<CellKey, Cell>) -> Option<Range> {
    let mut iter = cells.keys();
    let Some(first) = iter.next().copied() else {
        return None;
    };

    let mut min_row = first.row();
    let mut max_row = first.row();
    let mut min_col = first.col();
    let mut max_col = first.col();

    for key in iter.copied() {
        min_row = min_row.min(key.row());
        max_row = max_row.max(key.row());
        min_col = min_col.min(key.col());
        max_col = max_col.max(key.col());
    }

    Some(Range::new(
        CellRef::new(min_row, min_col),
        CellRef::new(max_row, max_col),
    ))
}
