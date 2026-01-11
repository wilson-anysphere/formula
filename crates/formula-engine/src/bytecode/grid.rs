use super::value::{CellCoord, ErrorKind, Value};
use ahash::AHashMap;

pub trait Grid: Sync {
    fn get_value(&self, coord: CellCoord) -> Value;

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

    fn bounds(&self) -> (i32, i32);

    #[inline]
    fn in_bounds(&self, coord: CellCoord) -> bool {
        let (rows, cols) = self.bounds();
        coord.row >= 0 && coord.col >= 0 && coord.row < rows && coord.col < cols
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
