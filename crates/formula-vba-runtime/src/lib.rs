//! `formula-vba-runtime` provides a small, sandboxed VBA interpreter intended to
//! power Formula's Excel compatibility layer (L4 Execute).
//!
//! This crate is **not** a full VBA implementation; it targets a pragmatic
//! subset that is sufficient to run common recorded macros that manipulate
//! worksheets (e.g. `Range("A1").Value = 1`).
//!
//! The runtime exposes:
//! - A parser that turns VBA into an AST (`VbaProgram`).
//! - An interpreter (`VbaRuntime`) that executes procedures with an Excel-like
//!   object model subset.
//! - A sandbox policy (`VbaSandboxPolicy`) that enforces permissions and a time
//!   limit.

mod ast;
mod lexer;
mod object_model;
mod parser;
mod runtime;
mod sandbox;
mod value;

pub use crate::ast::{ProcedureKind, VbaProgram};
pub use crate::ast::{
    ArrayDim, BinOp, CallArg, CaseComparisonOp, CaseCondition, ConstDecl, Expr, LoopConditionKind,
    ParamDef, ProcedureDef, SelectCaseArm, Stmt, UnOp, VarDecl, VbaType,
};
pub use crate::object_model::{
    a1_to_row_col, row_col_to_a1, InMemoryWorkbook, Spreadsheet, VbaRangeRef,
};
pub use crate::runtime::{ExecutionResult, VbaError, VbaRuntime};
pub use crate::sandbox::{Permission, PermissionChecker, VbaSandboxPolicy};
pub use crate::value::VbaValue;

/// Parse VBA source code into a [`VbaProgram`].
pub fn parse_program(source: &str) -> Result<VbaProgram, VbaError> {
    parser::parse_program(source)
}
