use std::collections::{BTreeMap, HashMap};

use serde::{Deserialize, Serialize};

use crate::{Cell, CellKey, CellRef, CellValue, Range};

/// Identifier for a worksheet within a workbook.
pub type WorksheetId = u32;

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

fn is_false(b: &bool) -> bool {
    !*b
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

/// A worksheet (sheet tab) containing sparse cells and per-row/column metadata.
#[derive(Clone, Debug, Serialize)]
pub struct Worksheet {
    /// Stable worksheet identifier.
    pub id: WorksheetId,
    /// User-visible name.
    pub name: String,

    /// Sparse cell storage; only non-empty cells are stored.
    #[serde(default)]
    cells: HashMap<CellKey, Cell>,

    /// Bounding box of stored cells.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    used_range: Option<Range>,

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
}

impl Worksheet {
    /// Create a new empty worksheet.
    pub fn new(id: WorksheetId, name: impl Into<String>) -> Self {
        Self {
            id,
            name: name.into(),
            cells: HashMap::new(),
            used_range: None,
            row_count: default_row_count(),
            col_count: default_col_count(),
            row_properties: BTreeMap::new(),
            col_properties: BTreeMap::new(),
            frozen_rows: 0,
            frozen_cols: 0,
            zoom: default_zoom(),
        }
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

    /// Get a cell record if it is present in the sparse store.
    pub fn cell(&self, cell: CellRef) -> Option<&Cell> {
        self.cells.get(&CellKey::from(cell))
    }

    /// Get a cell's value, returning [`CellValue::Empty`] if unset.
    pub fn value(&self, cell: CellRef) -> CellValue {
        self.cell(cell)
            .map(|c| c.value.clone())
            .unwrap_or(CellValue::Empty)
    }

    /// Set or replace a cell record.
    ///
    /// If the cell becomes "truly empty", it is removed from storage.
    pub fn set_cell(&mut self, cell_ref: CellRef, cell: Cell) {
        let key = CellKey::from(cell_ref);

        if cell.is_truly_empty() {
            let removed = self.cells.remove(&key).is_some();
            if removed {
                self.on_cell_removed(cell_ref);
            }
            return;
        }

        let is_new = self.cells.insert(key, cell).is_none();
        if is_new {
            self.on_cell_inserted(cell_ref);
        } else {
            // Existing cell updated; used range can only expand if the sheet was empty.
            if self.used_range.is_none() {
                self.used_range = Some(Range::new(cell_ref, cell_ref));
            }
        }
    }

    /// Convenience: set only the value for a cell.
    ///
    /// If the target cell does not exist yet, it is created with default style.
    pub fn set_value(&mut self, cell_ref: CellRef, value: CellValue) {
        let key = CellKey::from(cell_ref);

        match self.cells.get_mut(&key) {
            Some(cell) => {
                cell.value = value;
                if cell.is_truly_empty() {
                    self.cells.remove(&key);
                    self.on_cell_removed(cell_ref);
                }
            }
            None => {
                if value == CellValue::Empty {
                    return;
                }
                self.cells.insert(key, Cell::new(value));
                self.on_cell_inserted(cell_ref);
            }
        }
    }

    /// Remove any stored record for this cell.
    pub fn clear_cell(&mut self, cell_ref: CellRef) {
        let key = CellKey::from(cell_ref);
        if self.cells.remove(&key).is_some() {
            self.on_cell_removed(cell_ref);
        }
    }

    /// Iterate over all stored cells.
    pub fn iter_cells(&self) -> impl Iterator<Item = (CellRef, &Cell)> {
        self.cells.iter().map(|(k, v)| (k.to_ref(), v))
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
            cells: HashMap<CellKey, Cell>,
            #[serde(default = "default_row_count")]
            row_count: u32,
            #[serde(default = "default_col_count")]
            col_count: u32,
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
        }

        let helper = Helper::deserialize(deserializer)?;
        let used_range = compute_used_range(&helper.cells);

        Ok(Worksheet {
            id: helper.id,
            name: helper.name,
            cells: helper.cells,
            used_range,
            row_count: helper.row_count,
            col_count: helper.col_count,
            row_properties: helper.row_properties,
            col_properties: helper.col_properties,
            frozen_rows: helper.frozen_rows,
            frozen_cols: helper.frozen_cols,
            zoom: helper.zoom,
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
