mod ops;
pub(crate) mod rewrite;

pub use ops::{CellChange, CellSnapshot, EditError, EditOp, EditResult, FormulaRewrite, MovedRange};
