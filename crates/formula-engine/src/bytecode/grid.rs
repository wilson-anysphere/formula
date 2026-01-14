use super::value::{CellCoord, ErrorKind, SheetId, Value};
use ahash::AHashMap;
use formula_model::{EXCEL_MAX_COLS, EXCEL_MAX_ROWS};

const EXCEL_MAX_ROWS_I32: i32 = EXCEL_MAX_ROWS as i32;
const EXCEL_MAX_COLS_I32: i32 = EXCEL_MAX_COLS as i32;

pub trait Grid: Sync {
    fn get_value(&self, coord: CellCoord) -> Value;

    /// Returns the current worksheet tab order index for `sheet_id`.
    ///
    /// Excel defines the ordering of multi-area references (e.g. 3D sheet spans like
    /// `Sheet1:Sheet3!A1`, or `INDEX(..., area_num)`) based on workbook sheet tab order, not the
    /// internal numeric sheet id.
    ///
    /// Most single-sheet grid backends historically used sheet ids that matched tab order, so the
    /// default implementation preserves the old behavior by treating the sheet id itself as the
    /// order index. Multi-sheet backends with stable sheet ids should override this to return the
    /// current tab position.
    #[inline]
    fn sheet_order_index(&self, sheet_id: usize) -> Option<usize> {
        Some(sheet_id)
    }

    /// Return the sheet tab order for an external workbook.
    ///
    /// This is used to order multi-area references that span multiple external sheets (e.g. unions
    /// like `([Book.xlsx]Sheet2!A1,[Book.xlsx]Sheet10!A1)`) in a way that matches Excel semantics
    /// for 3D references: order is defined by workbook tab order, not by lexicographic sheet name.
    ///
    /// Returning `None` indicates the sheet order is unavailable, in which case the bytecode
    /// runtime falls back to lexicographic ordering for external sheet keys.
    #[inline]
    fn external_sheet_order(&self, _workbook: &str) -> Option<Vec<String>> {
        None
    }

    /// Get a value from a specific sheet.
    ///
    /// Bytecode formulas that don't use explicit sheet-qualified references can ignore the sheet
    /// id. Multi-sheet backends (like the engine) should override this to support 3D references.
    #[inline]
    fn get_value_on_sheet(&self, sheet: &SheetId, coord: CellCoord) -> Value {
        match sheet {
            SheetId::Local(_) => self.get_value(coord),
            SheetId::External(_) => Value::Error(ErrorKind::Ref),
        }
    }

    /// Record that a cell/range reference was dereferenced during evaluation.
    ///
    /// The bytecode runtime invokes this hook when it materializes reference values (e.g. turning a
    /// `Value::Range` into an array) or scans ranges for functions like `SUM`.
    ///
    /// Engine-side grids can override this to collect dynamic dependency information for formulas
    /// whose precedents are not statically known (e.g. `OFFSET`, `INDIRECT`).
    #[inline]
    fn record_reference(&self, sheet: usize, start: CellCoord, end: CellCoord) {
        let _ = sheet;
        let _ = start;
        let _ = end;
    }

    /// Sheet-aware variant of [`Grid::record_reference`].
    ///
    /// This is invoked for sheet-qualified references (e.g. `Sheet2!A1` or external workbook
    /// references). The default implementation records local sheets and ignores external sheets.
    #[inline]
    fn record_reference_on_sheet(&self, sheet: &SheetId, start: CellCoord, end: CellCoord) {
        if let SheetId::Local(sheet_id) = sheet {
            self.record_reference(*sheet_id, start, end);
        }
    }

    /// Source worksheet id for the grid.
    ///
    /// This is used for deterministic volatile behavior in the bytecode backend (e.g. RAND,
    /// RANDBETWEEN) so results differ across sheets.
    ///
    /// Grids that are not associated with a particular sheet can rely on the default `0`.
    #[inline]
    fn sheet_id(&self) -> usize {
        0
    }

    #[inline]
    fn get_number(&self, coord: CellCoord) -> Option<f64> {
        match self.get_value(coord) {
            Value::Number(v) => Some(v),
            Value::Bool(v) => Some(if v { 1.0 } else { 0.0 }),
            _ => None,
        }
    }

    /// Return a contiguous slice for a single-column row range if the backing storage is columnar.
    fn column_slice(&self, col: i32, row_start: i32, row_end: i32) -> Option<&[f64]>;

    /// Like [`Grid::column_slice`], but only when the underlying cells are strictly numeric/blank.
    ///
    /// This is used by lookup functions where treating logical/text values as `NaN` would produce
    /// incorrect comparison ordering (e.g. for approximate matches). Implementations that allow
    /// non-numeric values in `column_slice` should override this to apply stricter validation.
    #[inline]
    fn column_slice_strict_numeric(
        &self,
        col: i32,
        row_start: i32,
        row_end: i32,
    ) -> Option<&[f64]> {
        self.column_slice(col, row_start, row_end)
    }

    /// Optional sparse iteration over populated cells.
    ///
    /// When implemented, this should yield coordinates and values for cells that have a stored
    /// value in the backing grid (i.e. non-implicit blanks). The iterator does **not** need to
    /// include implicit empty cells; callers can account for them separately when needed.
    ///
    /// This is primarily used to optimize large-range aggregates like `SUM(A:A)` on sparse sheets
    /// without allocating/visiting every cell in the range.
    #[inline]
    fn iter_cells(&self) -> Option<Box<dyn Iterator<Item = (CellCoord, Value)> + '_>> {
        None
    }

    /// Sheet-aware variant of [`Grid::iter_cells`].
    ///
    /// This is used to support sparse iteration for multi-sheet references (e.g. 3D sheet spans)
    /// without forcing dense column caches or scanning every implicit blank cell.
    #[inline]
    fn iter_cells_on_sheet(
        &self,
        sheet: &SheetId,
    ) -> Option<Box<dyn Iterator<Item = (CellCoord, Value)> + '_>> {
        match sheet {
            SheetId::Local(_) => self.iter_cells(),
            SheetId::External(_) => None,
        }
    }

    /// Sheet-aware variant of [`Grid::column_slice`].
    #[inline]
    fn column_slice_on_sheet(
        &self,
        sheet: &SheetId,
        col: i32,
        row_start: i32,
        row_end: i32,
    ) -> Option<&[f64]> {
        match sheet {
            SheetId::Local(_) => self.column_slice(col, row_start, row_end),
            SheetId::External(_) => None,
        }
    }

    /// Sheet-aware variant of [`Grid::column_slice_strict_numeric`].
    ///
    /// Implementations that allow non-numeric values in `column_slice_on_sheet` should override
    /// this to apply stricter validation.
    #[inline]
    fn column_slice_on_sheet_strict_numeric(
        &self,
        sheet: &SheetId,
        col: i32,
        row_start: i32,
        row_end: i32,
    ) -> Option<&[f64]> {
        self.column_slice_on_sheet(sheet, col, row_start, row_end)
    }
    fn bounds(&self) -> (i32, i32);

    /// Sheet-aware variant of [`Grid::bounds`].
    #[inline]
    fn bounds_on_sheet(&self, sheet: &SheetId) -> (i32, i32) {
        match sheet {
            SheetId::Local(_) => self.bounds(),
            // External sheets do not expose true dimensions, but references are still valid within
            // Excel's fixed maximum grid. This matches the AST evaluator which treats external
            // bounds as unknown/valid.
            SheetId::External(_) => (EXCEL_MAX_ROWS_I32, EXCEL_MAX_COLS_I32),
        }
    }

    /// Resolve a worksheet display name to an internal sheet id.
    ///
    /// This is used by volatile reference functions like `INDIRECT` that parse sheet names at
    /// runtime. Multi-sheet backends (like the engine) should override this. Single-sheet
    /// backends can rely on the default implementation, which does not resolve any names.
    ///
    /// Expected semantics: match Excel's Unicode-aware sheet name comparison by applying Unicode
    /// NFKC (compatibility normalization) and then Unicode uppercasing (see
    /// [`formula_model::sheet_name_eq_case_insensitive`]).
    #[inline]
    fn resolve_sheet_name(&self, _name: &str) -> Option<usize> {
        None
    }

    #[inline]
    fn in_bounds(&self, coord: CellCoord) -> bool {
        let (rows, cols) = self.bounds();
        coord.row >= 0 && coord.col >= 0 && coord.row < rows && coord.col < cols
    }

    #[inline]
    fn in_bounds_on_sheet(&self, sheet: &SheetId, coord: CellCoord) -> bool {
        let (rows, cols) = self.bounds_on_sheet(sheet);
        coord.row >= 0 && coord.col >= 0 && coord.row < rows && coord.col < cols
    }

    /// If `addr` is part of a spilled array, returns the spill origin cell.
    ///
    /// This mirrors the semantics of [`crate::eval::ValueResolver::spill_origin`]. Bytecode
    /// backends that don't support dynamic arrays can leave the default implementation, which
    /// behaves as if there are no spills.
    #[inline]
    fn spill_origin(
        &self,
        _sheet_id: &SheetId,
        _addr: crate::eval::CellAddr,
    ) -> Option<crate::eval::CellAddr> {
        None
    }

    /// If `origin` is the origin of a spilled array, returns the full spill range (inclusive).
    ///
    /// This mirrors the semantics of [`crate::eval::ValueResolver::spill_range`]. Bytecode
    /// backends that don't support dynamic arrays can leave the default implementation, which
    /// behaves as if there are no spills.
    #[inline]
    fn spill_range(
        &self,
        _sheet_id: &SheetId,
        _origin: crate::eval::CellAddr,
    ) -> Option<(crate::eval::CellAddr, crate::eval::CellAddr)> {
        None
    }
}

pub trait GridMut: Grid {
    fn set_value(&mut self, coord: CellCoord, value: Value);
}

/// HashMap-backed sparse cell storage.
#[derive(Default)]
pub struct SparseGrid {
    rows: i32,
    cols: i32,
    cells: AHashMap<(i32, i32), Value>,
}

impl SparseGrid {
    pub fn new(rows: i32, cols: i32) -> Self {
        Self {
            rows,
            cols,
            cells: AHashMap::new(),
        }
    }
}

impl Grid for SparseGrid {
    fn get_value(&self, coord: CellCoord) -> Value {
        if !self.in_bounds(coord) {
            return Value::Error(ErrorKind::Ref);
        }
        self.cells
            .get(&(coord.row, coord.col))
            .cloned()
            .unwrap_or(Value::Empty)
    }

    fn column_slice(&self, _col: i32, _row_start: i32, _row_end: i32) -> Option<&[f64]> {
        None
    }

    fn bounds(&self) -> (i32, i32) {
        (self.rows, self.cols)
    }
}

impl GridMut for SparseGrid {
    fn set_value(&mut self, coord: CellCoord, value: Value) {
        if !self.in_bounds(coord) {
            return;
        }
        match value {
            Value::Empty => {
                self.cells.remove(&(coord.row, coord.col));
            }
            v => {
                self.cells.insert((coord.row, coord.col), v);
            }
        }
    }
}

/// Column-major numeric storage with NaN representing empty/non-numeric cells.
pub struct ColumnarGrid {
    rows: i32,
    cols: i32,
    cols_data: Vec<Vec<f64>>,
}

impl ColumnarGrid {
    pub fn new(rows: i32, cols: i32) -> Self {
        let mut cols_data = Vec::with_capacity(cols as usize);
        for _ in 0..cols {
            cols_data.push(vec![f64::NAN; rows as usize]);
        }
        Self {
            rows,
            cols,
            cols_data,
        }
    }

    #[inline]
    pub fn set_number(&mut self, coord: CellCoord, value: f64) {
        if coord.row < 0 || coord.col < 0 || coord.row >= self.rows || coord.col >= self.cols {
            return;
        }
        self.cols_data[coord.col as usize][coord.row as usize] = value;
    }

    #[inline]
    pub fn get_raw_number(&self, coord: CellCoord) -> Option<f64> {
        if coord.row < 0 || coord.col < 0 || coord.row >= self.rows || coord.col >= self.cols {
            return None;
        }
        Some(self.cols_data[coord.col as usize][coord.row as usize])
    }
}

impl Grid for ColumnarGrid {
    fn get_value(&self, coord: CellCoord) -> Value {
        let v = match self.get_raw_number(coord) {
            Some(v) => v,
            None => return Value::Error(ErrorKind::Ref),
        };
        if v.is_nan() {
            Value::Empty
        } else {
            Value::Number(v)
        }
    }

    fn column_slice(&self, col: i32, row_start: i32, row_end: i32) -> Option<&[f64]> {
        if col < 0 || row_start < 0 || row_end < 0 || col >= self.cols || row_end >= self.rows {
            return None;
        }
        let col_vec = self.cols_data.get(col as usize)?;
        let start = row_start as usize;
        let end = row_end as usize;
        if start > end || end >= col_vec.len() {
            return None;
        }
        Some(&col_vec[start..=end])
    }

    fn bounds(&self) -> (i32, i32) {
        (self.rows, self.cols)
    }
}

impl GridMut for ColumnarGrid {
    fn set_value(&mut self, coord: CellCoord, value: Value) {
        if coord.row < 0 || coord.col < 0 || coord.row >= self.rows || coord.col >= self.cols {
            return;
        }
        let slot = &mut self.cols_data[coord.col as usize][coord.row as usize];
        *slot = match value {
            Value::Number(v) => v,
            Value::Bool(v) => {
                if v {
                    1.0
                } else {
                    0.0
                }
            }
            Value::Empty => f64::NAN,
            _ => f64::NAN,
        };
    }
}
