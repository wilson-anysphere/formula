use std::fmt;
use std::sync::Arc;

/// 0-indexed cell coordinate.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
pub struct CellCoord {
    pub row: i32,
    pub col: i32,
}

impl CellCoord {
    #[inline]
    pub const fn new(row: i32, col: i32) -> Self {
        Self { row, col }
    }
}

/// Excel-style cell/range reference represented as row/col that can be either absolute
/// (stored as 0-indexed coordinate) or relative (stored as offset from the formula cell).
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Ref {
    pub row: i32,
    pub col: i32,
    pub row_abs: bool,
    pub col_abs: bool,
}

impl Ref {
    #[inline]
    pub const fn new(row: i32, col: i32, row_abs: bool, col_abs: bool) -> Self {
        Self {
            row,
            col,
            row_abs,
            col_abs,
        }
    }

    #[inline]
    pub fn resolve(self, base: CellCoord) -> CellCoord {
        let row = if self.row_abs { self.row } else { base.row + self.row };
        let col = if self.col_abs { self.col } else { base.col + self.col };
        CellCoord { row, col }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct RangeRef {
    pub start: Ref,
    pub end: Ref,
}

impl RangeRef {
    #[inline]
    pub const fn new(start: Ref, end: Ref) -> Self {
        Self { start, end }
    }

    #[inline]
    pub fn resolve(self, base: CellCoord) -> ResolvedRange {
        let a = self.start.resolve(base);
        let b = self.end.resolve(base);
        ResolvedRange::from_coords(a, b)
    }
}

/// Resolved absolute range (inclusive bounds).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ResolvedRange {
    pub row_start: i32,
    pub row_end: i32,
    pub col_start: i32,
    pub col_end: i32,
}

impl ResolvedRange {
    #[inline]
    pub fn from_coords(a: CellCoord, b: CellCoord) -> Self {
        let row_start = a.row.min(b.row);
        let row_end = a.row.max(b.row);
        let col_start = a.col.min(b.col);
        let col_end = a.col.max(b.col);
        Self {
            row_start,
            row_end,
            col_start,
            col_end,
        }
    }

    #[inline]
    pub fn rows(self) -> i32 {
        self.row_end - self.row_start + 1
    }

    #[inline]
    pub fn cols(self) -> i32 {
        self.col_end - self.col_start + 1
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ErrorKind {
    Div0,
    Ref,
    Value,
    Name,
    Num,
    NA,
    Calc,
}

#[derive(Clone, Debug)]
pub struct Array {
    pub rows: usize,
    pub cols: usize,
    pub values: Vec<f64>,
}

impl Array {
    #[inline]
    pub fn new(rows: usize, cols: usize, values: Vec<f64>) -> Self {
        debug_assert_eq!(rows * cols, values.len());
        Self { rows, cols, values }
    }

    #[inline]
    pub fn len(&self) -> usize {
        self.values.len()
    }

    #[inline]
    pub fn as_slice(&self) -> &[f64] {
        &self.values
    }
}

/// Runtime value used by both AST and bytecode evaluators.
#[derive(Clone, Debug)]
pub enum Value {
    Number(f64),
    Bool(bool),
    Text(Arc<str>),
    Array(Array),
    Range(RangeRef),
    Empty,
    Error(ErrorKind),
}

impl Value {
    #[inline]
    pub fn as_f64(&self) -> Option<f64> {
        match self {
            Value::Number(v) => Some(*v),
            Value::Bool(v) => Some(if *v { 1.0 } else { 0.0 }),
            Value::Empty => None,
            _ => None,
        }
    }
}

impl PartialEq for Value {
    fn eq(&self, other: &Self) -> bool {
        use Value::*;
        match (self, other) {
            (Number(a), Number(b)) => a.to_bits() == b.to_bits(),
            (Bool(a), Bool(b)) => a == b,
            (Text(a), Text(b)) => a == b,
            (Empty, Empty) => true,
            (Error(a), Error(b)) => a == b,
            (Range(a), Range(b)) => a == b,
            (Array(a), Array(b)) => a.rows == b.rows
                && a.cols == b.cols
                && a.values.len() == b.values.len()
                && a.values
                    .iter()
                    .zip(&b.values)
                    .all(|(x, y)| x.to_bits() == y.to_bits()),
            _ => false,
        }
    }
}

impl Eq for Value {}

impl fmt::Display for ErrorKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let s = match self {
            ErrorKind::Div0 => "#DIV/0!",
            ErrorKind::Ref => "#REF!",
            ErrorKind::Value => "#VALUE!",
            ErrorKind::Name => "#NAME?",
            ErrorKind::Num => "#NUM!",
            ErrorKind::NA => "#N/A",
            ErrorKind::Calc => "#CALC!",
        };
        f.write_str(s)
    }
}

