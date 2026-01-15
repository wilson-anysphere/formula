use super::program::LambdaTemplate;
use std::fmt;
use std::sync::Arc;

use crate::value::{EntityValue, RecordValue};

/// Worksheet identifier used by the bytecode backend.
///
/// This must be stable and uniquely identify the sheet for bytecode program caching.
/// In particular, external sheet references are keyed by a canonical string
/// (e.g. `"[Book.xlsx]Sheet1"`) so that cached programs do not collide across different external
/// workbooks/sheets.
#[derive(Clone, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum SheetId {
    /// 0-indexed sheet id within the current workbook/engine snapshot.
    Local(usize),
    /// Canonical external sheet key string (e.g. `"[Book.xlsx]Sheet1"`).
    External(Arc<str>),
    /// External-workbook 3D sheet span (`[Book.xlsx]Sheet1:Sheet3`) that must be expanded using the
    /// external workbook's sheet tab order at evaluation time.
    ///
    /// This variant exists so the bytecode backend can represent external 3D references without
    /// baking a particular sheet-order expansion into the bytecode cache key.
    ExternalSpan {
        workbook: Arc<str>,
        start: Arc<str>,
        end: Arc<str>,
    },
}

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
        let row = if self.row_abs {
            self.row
        } else {
            base.row + self.row
        };
        let col = if self.col_abs {
            self.col
        } else {
            base.col + self.col
        };
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

/// A range on a specific sheet.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub struct SheetRangeRef {
    pub sheet: SheetId,
    pub range: RangeRef,
}

impl SheetRangeRef {
    #[inline]
    pub fn new(sheet: SheetId, range: RangeRef) -> Self {
        Self { sheet, range }
    }
}

/// A discontiguous reference consisting of multiple per-sheet range areas.
///
/// This is used to represent 3D sheet spans like `Sheet1:Sheet3!A1:B2`, which
/// expand to one rectangular area per sheet.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MultiRangeRef {
    pub areas: Arc<[SheetRangeRef]>,
}

impl MultiRangeRef {
    #[inline]
    pub fn new(areas: Arc<[SheetRangeRef]>) -> Self {
        Self { areas }
    }

    #[inline]
    pub fn is_empty(&self) -> bool {
        self.areas.is_empty()
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
    Null,
    Div0,
    Value,
    Ref,
    Name,
    Num,
    NA,
    GettingData,
    Spill,
    Calc,
    Field,
    Connect,
    Blocked,
    Unknown,
}

impl ErrorKind {
    pub fn from_code(raw: &str) -> Option<Self> {
        crate::value::ErrorKind::from_code(raw).map(Self::from)
    }

    pub fn as_code(self) -> &'static str {
        crate::value::ErrorKind::from(self).as_code()
    }
}

impl From<crate::value::ErrorKind> for ErrorKind {
    fn from(value: crate::value::ErrorKind) -> Self {
        match value {
            crate::value::ErrorKind::Null => ErrorKind::Null,
            crate::value::ErrorKind::Div0 => ErrorKind::Div0,
            crate::value::ErrorKind::Value => ErrorKind::Value,
            crate::value::ErrorKind::Ref => ErrorKind::Ref,
            crate::value::ErrorKind::Name => ErrorKind::Name,
            crate::value::ErrorKind::Num => ErrorKind::Num,
            crate::value::ErrorKind::NA => ErrorKind::NA,
            crate::value::ErrorKind::GettingData => ErrorKind::GettingData,
            crate::value::ErrorKind::Spill => ErrorKind::Spill,
            crate::value::ErrorKind::Calc => ErrorKind::Calc,
            crate::value::ErrorKind::Field => ErrorKind::Field,
            crate::value::ErrorKind::Connect => ErrorKind::Connect,
            crate::value::ErrorKind::Blocked => ErrorKind::Blocked,
            crate::value::ErrorKind::Unknown => ErrorKind::Unknown,
        }
    }
}

impl From<ErrorKind> for crate::value::ErrorKind {
    fn from(value: ErrorKind) -> Self {
        match value {
            ErrorKind::Null => crate::value::ErrorKind::Null,
            ErrorKind::Div0 => crate::value::ErrorKind::Div0,
            ErrorKind::Ref => crate::value::ErrorKind::Ref,
            ErrorKind::Value => crate::value::ErrorKind::Value,
            ErrorKind::Name => crate::value::ErrorKind::Name,
            ErrorKind::Num => crate::value::ErrorKind::Num,
            ErrorKind::NA => crate::value::ErrorKind::NA,
            ErrorKind::GettingData => crate::value::ErrorKind::GettingData,
            ErrorKind::Spill => crate::value::ErrorKind::Spill,
            ErrorKind::Calc => crate::value::ErrorKind::Calc,
            ErrorKind::Field => crate::value::ErrorKind::Field,
            ErrorKind::Connect => crate::value::ErrorKind::Connect,
            ErrorKind::Blocked => crate::value::ErrorKind::Blocked,
            ErrorKind::Unknown => crate::value::ErrorKind::Unknown,
        }
    }
}

#[derive(Clone, Debug)]
pub struct Array {
    pub rows: usize,
    pub cols: usize,
    /// Row-major order values (length = rows * cols).
    pub values: Arc<Vec<Value>>,
}

impl Array {
    #[inline]
    pub fn new(rows: usize, cols: usize, values: Vec<Value>) -> Self {
        debug_assert_eq!(rows * cols, values.len());
        Self {
            rows,
            cols,
            values: Arc::new(values),
        }
    }

    #[inline]
    pub fn len(&self) -> usize {
        self.values.len()
    }

    #[inline]
    pub fn get(&self, row: usize, col: usize) -> Option<&Value> {
        if row >= self.rows || col >= self.cols {
            return None;
        }
        self.values.get(row * self.cols + col)
    }

    #[inline]
    pub fn iter(&self) -> impl Iterator<Item = &Value> {
        self.values.iter()
    }
}

#[derive(Clone)]
pub struct Lambda {
    pub template: Arc<LambdaTemplate>,
    /// Captured values aligned with `template.captures`.
    pub captures: Arc<[Value]>,
}

impl fmt::Debug for Lambda {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Lambda")
            .field("params", &self.template.params)
            .field("body_key", &self.template.body.key())
            .field("captures_len", &self.captures.len())
            .finish()
    }
}

impl PartialEq for Lambda {
    fn eq(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.template, &other.template) && Arc::ptr_eq(&self.captures, &other.captures)
    }
}

impl Eq for Lambda {}

/// Runtime value used by both AST and bytecode evaluators.
#[derive(Clone, Debug)]
pub enum Value {
    Number(f64),
    Bool(bool),
    Text(Arc<str>),
    Entity(Arc<EntityValue>),
    Record(Arc<RecordValue>),
    Array(Array),
    Range(RangeRef),
    MultiRange(MultiRangeRef),
    Lambda(Lambda),
    Empty,
    /// Placeholder for a missing/omitted function argument (e.g. `IF(,1,2)`).
    ///
    /// This is distinct from `Empty` (blank cell value) so bytecode runtime implementations can
    /// preserve Excel's semantics for optional arguments where "omitted" differs from "blank".
    Missing,
    Error(ErrorKind),
}

impl Value {
    #[inline]
    pub fn as_f64(&self) -> Option<f64> {
        match self {
            Value::Number(v) => Some(*v),
            Value::Bool(v) => Some(if *v { 1.0 } else { 0.0 }),
            Value::Empty | Value::Missing => None,
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
            (Text(a), Text(b)) => Arc::ptr_eq(a, b) || a == b,
            (Entity(a), Entity(b)) => Arc::ptr_eq(a, b) || a == b,
            (Record(a), Record(b)) => Arc::ptr_eq(a, b) || a == b,
            (Empty, Empty) => true,
            (Missing, Missing) => true,
            (Error(a), Error(b)) => a == b,
            (Range(a), Range(b)) => a == b,
            (MultiRange(a), MultiRange(b)) => Arc::ptr_eq(&a.areas, &b.areas) || a.areas == b.areas,
            (Lambda(a), Lambda(b)) => a == b,
            (Array(a), Array(b)) => {
                a.rows == b.rows
                    && a.cols == b.cols
                    && (Arc::ptr_eq(&a.values, &b.values) || a.values == b.values)
            }
            _ => false,
        }
    }
}

impl Eq for Value {}
impl fmt::Display for ErrorKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.as_code())
    }
}
