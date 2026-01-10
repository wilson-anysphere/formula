use crate::value::Value;
use formula_model::{CellRef, Range};

/// High-level structural edit operation, intended to behave like Excel.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum EditOp {
    InsertRows { sheet: String, row: u32, count: u32 },
    DeleteRows { sheet: String, row: u32, count: u32 },
    InsertCols { sheet: String, col: u32, count: u32 },
    DeleteCols { sheet: String, col: u32, count: u32 },
    InsertCellsShiftRight { sheet: String, range: Range },
    InsertCellsShiftDown { sheet: String, range: Range },
    DeleteCellsShiftLeft { sheet: String, range: Range },
    DeleteCellsShiftUp { sheet: String, range: Range },
    MoveRange {
        sheet: String,
        src: Range,
        dst_top_left: CellRef,
    },
    CopyRange {
        sheet: String,
        src: Range,
        dst_top_left: CellRef,
    },
    Fill { sheet: String, src: Range, dst: Range },
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub enum EditError {
    SheetNotFound(String),
    InvalidCount,
    InvalidRange,
    OverlappingMove,
    Engine(String),
}

#[derive(Clone, Debug, PartialEq)]
pub struct CellSnapshot {
    pub value: Value,
    pub formula: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct CellChange {
    pub sheet: String,
    pub cell: CellRef,
    pub before: Option<CellSnapshot>,
    pub after: Option<CellSnapshot>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MovedRange {
    pub sheet: String,
    pub from: Range,
    pub to: Range,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct FormulaRewrite {
    pub sheet: String,
    pub cell: CellRef,
    pub before: String,
    pub after: String,
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct EditResult {
    pub changed_cells: Vec<CellChange>,
    pub moved_ranges: Vec<MovedRange>,
    pub formula_rewrites: Vec<FormulaRewrite>,
}

