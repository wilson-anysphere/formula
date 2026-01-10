use std::collections::{BTreeMap, HashMap};
use std::sync::Arc;

use serde::de::Error as _;
use serde::{Deserialize, Serialize};

use formula_columnar::{ColumnType as ColumnarType, ColumnarTable, Value as ColumnarValue};

use crate::{
    A1ParseError, Cell, CellKey, CellRef, CellValue, Hyperlink, MergeError, MergedRegions, Range,
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
    1.0
}

fn is_default_zoom(z: &f32) -> bool {
    (*z - 1.0).abs() < f32::EPSILON
}

fn is_visible(v: &SheetVisibility) -> bool {
    matches!(v, SheetVisibility::Visible)
}

fn is_false(b: &bool) -> bool {
    !*b
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
    /// Whether the row is hidden.
    #[serde(default, skip_serializing_if = "is_false")]
    pub hidden: bool,
}

/// Per-column overrides.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize, Default)]
pub struct ColProperties {
    /// Column width in Excel "character" units.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<f32>,
    /// Whether the column is hidden.
    #[serde(default, skip_serializing_if = "is_false")]
    pub hidden: bool,
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

    /// Sheet visibility state.
    #[serde(default, skip_serializing_if = "is_visible")]
    pub visibility: SheetVisibility,

    /// Optional tab color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_color: Option<TabColor>,

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

    /// Frozen pane row count (top).
    #[serde(default)]
    pub frozen_rows: u32,

    /// Frozen pane column count (left).
    #[serde(default)]
    pub frozen_cols: u32,

    /// Sheet zoom level (1.0 = 100%).
    #[serde(default = "default_zoom", skip_serializing_if = "is_default_zoom")]
    pub zoom: f32,

    /// Excel tables (structured ranges) hosted on this worksheet.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub tables: Vec<Table>,

    /// Optional columnar backing store for large imported datasets.
    ///
    /// This is runtime-only for now; persistence is handled by the storage layer.
    #[serde(skip)]
    columnar: Option<ColumnarBackend>,

    /// Hyperlinks anchored to cells/ranges in this worksheet.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub hyperlinks: Vec<Hyperlink>,
}

impl Worksheet {
    /// Create a new empty worksheet.
    pub fn new(id: WorksheetId, name: impl Into<String>) -> Self {
        Self {
            id,
            name: name.into(),
            visibility: SheetVisibility::Visible,
            tab_color: None,
            cells: HashMap::new(),
            used_range: None,
            merged_regions: MergedRegions::new(),
            row_count: default_row_count(),
            col_count: default_col_count(),
            row_properties: BTreeMap::new(),
            col_properties: BTreeMap::new(),
            frozen_rows: 0,
            frozen_cols: 0,
            zoom: default_zoom(),
            tables: Vec::new(),
            columnar: None,
            hyperlinks: Vec::new(),
        }
    }

    /// Find the first table containing `cell`.
    pub fn table_for_cell(&self, cell: CellRef) -> Option<&Table> {
        self.tables.iter().find(|t| t.range.contains(cell))
    }

    /// Number of stored cells.
    ///
    /// This is proportional to memory usage for the sheet's cell content.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Get the current used range (bounding box of stored cells).
    pub fn used_range(&self) -> Option<Range> {
        self.used_range
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
    }

    /// Remove any columnar backing table.
    pub fn clear_columnar_table(&mut self) {
        self.columnar = None;
    }

    /// Returns the backing columnar table, if present.
    pub fn columnar_table(&self) -> Option<&Arc<ColumnarTable>> {
        self.columnar.as_ref().map(|c| &c.table)
    }

    /// Get per-row properties if an override exists.
    pub fn row_properties(&self, row: u32) -> Option<&RowProperties> {
        self.row_properties.get(&row)
    }

    /// Get per-column properties if an override exists.
    pub fn col_properties(&self, col: u32) -> Option<&ColProperties> {
        self.col_properties.get(&col)
    }

    /// Set (or clear) the height override for a row.
    ///
    /// Passing `None` removes the height override. If the row has no overrides
    /// remaining, its entry is removed from the map.
    pub fn set_row_height(&mut self, row: u32, height: Option<f32>) {
        assert!(
            row < crate::cell::EXCEL_MAX_ROWS,
            "row out of Excel bounds: {row}"
        );
        match self.row_properties.get_mut(&row) {
            Some(props) => {
                props.height = height;
                if props.height.is_none() && !props.hidden {
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
        assert!(
            row < crate::cell::EXCEL_MAX_ROWS,
            "row out of Excel bounds: {row}"
        );
        match self.row_properties.get_mut(&row) {
            Some(props) => {
                props.hidden = hidden;
                if props.height.is_none() && !props.hidden {
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
                    },
                );
            }
        }
    }

    /// Set (or clear) the width override for a column.
    pub fn set_col_width(&mut self, col: u32, width: Option<f32>) {
        assert!(
            col < crate::cell::EXCEL_MAX_COLS,
            "col out of Excel bounds: {col}"
        );
        match self.col_properties.get_mut(&col) {
            Some(props) => {
                props.width = width;
                if props.width.is_none() && !props.hidden {
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
        match self.col_properties.get_mut(&col) {
            Some(props) => {
                props.hidden = hidden;
                if props.width.is_none() && !props.hidden {
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
                    },
                );
            }
        }
    }

    /// Get a cell record if it is present in the sparse store.
    pub fn cell(&self, cell: CellRef) -> Option<&Cell> {
        if cell.row >= crate::cell::EXCEL_MAX_ROWS || cell.col >= crate::cell::EXCEL_MAX_COLS {
            return None;
        }
        let anchor = self.merged_regions.resolve_cell(cell);
        self.cells.get(&CellKey::from_ref(anchor))
    }

    /// Get a mutable cell record if it is present in the sparse store.
    pub fn cell_mut(&mut self, cell: CellRef) -> Option<&mut Cell> {
        if cell.row >= crate::cell::EXCEL_MAX_ROWS || cell.col >= crate::cell::EXCEL_MAX_COLS {
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
        let rows = (range.end.row - range.start.row + 1) as usize;
        let cols = (range.end.col - range.start.col + 1) as usize;
        let mut columns = vec![vec![CellValue::Empty; rows]; cols];

        // Bulk fill from columnar backing (if present).
        if let Some(backend) = &self.columnar {
            fill_from_columnar(&mut columns, range, backend);
        }

        // Overlay sparse cells (edits / formulas / formatting) on top.
        for r_off in 0..rows {
            let row = range.start.row + r_off as u32;
            for c_off in 0..cols {
                let col = range.start.col + c_off as u32;
                let cell_ref = CellRef::new(row, col);
                if let Some(cell) = self.cell(cell_ref) {
                    columns[c_off][r_off] = cell.value.clone();
                }
            }
        }

        RangeBatch {
            start: range.start,
            columns,
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

    /// Set a cell formula.
    ///
    /// Setting a formula to `None` clears the formula and removes the cell from the sparse store
    /// if it becomes "truly empty".
    pub fn set_formula(&mut self, cell_ref: CellRef, formula: Option<String>) {
        let key = CellKey::from(cell_ref);

        match self.cells.get_mut(&key) {
            Some(cell) => {
                cell.formula = formula;
                if cell.is_truly_empty() {
                    self.cells.remove(&key);
                    self.on_cell_removed(cell_ref);
                }
            }
            None => {
                let Some(formula) = formula else {
                    return;
                };
                let mut cell = Cell::default();
                cell.formula = Some(formula);
                self.cells.insert(key, cell);
                self.on_cell_inserted(cell_ref);
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
    pub fn set_cell(&mut self, cell_ref: CellRef, cell: Cell) {
        let anchor = self.merged_regions.resolve_cell(cell_ref);
        let key = CellKey::from(anchor);

        if cell.is_truly_empty() {
            let removed = self.cells.remove(&key).is_some();
            if removed {
                self.on_cell_removed(anchor);
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
        for row in range.start.row..=range.end.row {
            for col in range.start.col..=range.end.col {
                let cell = CellRef::new(row, col);
                if cell != anchor {
                    removed_any |= self.cells.remove(&CellKey::from(cell)).is_some();
                }
            }
        }

        if removed_any {
            self.recompute_used_range();
        }

        self.merged_regions.add(range)
    }

    /// Unmerge any merged regions that intersect `range`.
    pub fn unmerge_range(&mut self, range: Range) -> usize {
        self.merged_regions.unmerge_range(range)
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
        let keys: Vec<CellKey> = self
            .cells
            .keys()
            .filter(|k| range.contains(k.to_ref()))
            .copied()
            .collect();

        if keys.is_empty() {
            return;
        }

        for key in keys {
            self.cells.remove(&key);
        }

        self.used_range = compute_used_range(&self.cells);
    }

    /// Iterate over all stored cells mutably.
    pub fn iter_cells_mut(&mut self) -> impl Iterator<Item = (CellRef, &mut Cell)> {
        self.cells.iter_mut().map(|(k, v)| (k.to_ref(), v))
    }

    /// Return the first hyperlink whose anchor range contains `cell`.
    pub fn hyperlink_at(&self, cell: CellRef) -> Option<&Hyperlink> {
        self.hyperlinks.iter().find(|h| h.range.contains(cell))
    }

    fn on_cell_inserted(&mut self, cell_ref: CellRef) {
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
            visibility: SheetVisibility,
            #[serde(default)]
            tab_color: Option<TabColor>,
            #[serde(default)]
            cells: HashMap<CellKey, Cell>,
            #[serde(default)]
            tables: Vec<Table>,
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
            frozen_rows: u32,
            #[serde(default)]
            frozen_cols: u32,
            #[serde(default = "default_zoom")]
            zoom: f32,
            #[serde(default)]
            hyperlinks: Vec<Hyperlink>,
        }

        let helper = Helper::deserialize(deserializer)?;
        let used_range = compute_used_range(&helper.cells);

        if helper.row_count == 0 {
            return Err(D::Error::custom("row_count must be >= 1"));
        }
        if helper.col_count == 0 {
            return Err(D::Error::custom("col_count must be >= 1"));
        }
        if helper.row_count > crate::cell::EXCEL_MAX_ROWS {
            return Err(D::Error::custom(format!(
                "row_count out of Excel bounds: {}",
                helper.row_count
            )));
        }
        if helper.col_count > crate::cell::EXCEL_MAX_COLS {
            return Err(D::Error::custom(format!(
                "col_count out of Excel bounds: {}",
                helper.col_count
            )));
        }

        for row in helper.row_properties.keys() {
            if *row >= crate::cell::EXCEL_MAX_ROWS {
                return Err(D::Error::custom(format!(
                    "row_properties row out of Excel bounds: {row}"
                )));
            }
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
            row_count = row_count.max(used.end.row + 1);
            col_count = col_count.max(used.end.col + 1);
        }

        if helper.frozen_rows > row_count {
            return Err(D::Error::custom(format!(
                "frozen_rows exceeds row_count: {} > {row_count}",
                helper.frozen_rows
            )));
        }
        if helper.frozen_cols > col_count {
            return Err(D::Error::custom(format!(
                "frozen_cols exceeds col_count: {} > {col_count}",
                helper.frozen_cols
            )));
        }

        if helper.zoom <= 0.0 {
            return Err(D::Error::custom(format!(
                "zoom must be > 0 (got {})",
                helper.zoom
            )));
        }

        Ok(Worksheet {
            id: helper.id,
            name: helper.name,
            visibility: helper.visibility,
            tab_color: helper.tab_color,
            cells: helper.cells,
            used_range,
            merged_regions: helper.merged_regions,
            row_count,
            col_count,
            row_properties: helper.row_properties,
            col_properties: helper.col_properties,
            frozen_rows: helper.frozen_rows,
            frozen_cols: helper.frozen_cols,
            zoom: helper.zoom,
            tables: helper.tables,
            columnar: None,
            hyperlinks: helper.hyperlinks,
        })
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
