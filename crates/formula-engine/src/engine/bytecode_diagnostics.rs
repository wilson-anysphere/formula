use crate::bytecode;
use crate::eval::CellAddr;
use std::collections::HashMap;
use std::sync::Arc;

/// Reason a formula was not compiled to the bytecode backend.
///
/// This is intended for coverage/benchmark harnesses so they can measure bytecode backend
/// adoption and prioritize which unsupported constructs should be implemented next.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum BytecodeCompileReason {
    /// The bytecode backend is disabled via [`crate::Engine::set_bytecode_enabled`].
    Disabled,
    /// The formula was rejected because it is volatile.
    ///
    /// Note: The bytecode backend now supports *thread-safe* volatile formulas (e.g. `NOW()`,
    /// `TODAY()`), so this reason is no longer emitted by the engine. It is kept for backward
    /// compatibility with earlier coverage-reporting tooling.
    Volatile,
    /// The formula calls a non-thread-safe function and cannot be evaluated in the bytecode VM.
    NotThreadSafe,
    /// The formula uses dynamic dependency functions (e.g. `OFFSET`, `INDIRECT`) which require
    /// dependency tracing during evaluation.
    ///
    /// Note: The bytecode backend now implements dependency tracing for bytecode evaluation, so
    /// this reason is no longer emitted by the engine. It is kept for backward compatibility with
    /// earlier coverage-reporting tooling.
    DynamicDependencies,
    /// Lowering the canonical formula AST into the bytecode AST failed.
    LowerError(bytecode::LowerError),
    /// The bytecode backend does not yet support this expression shape (even if lowering succeeded).
    IneligibleExpr,
    /// The formula calls a worksheet function that the bytecode backend does not yet implement.
    ///
    /// This is reported separately from `IneligibleExpr` so coverage tools can see which missing
    /// functions account for the majority of AST fallbacks.
    UnsupportedFunction(Arc<str>),
    /// The formula references cells/ranges that fall outside the Excel grid.
    ExceedsGridLimits,
    /// The formula contains a range reference that exceeds the bytecode backend's cell-count limit.
    ExceedsRangeCellLimit,
    /// The workbook has custom sheet dimensions that don't match Excel's fixed worksheet bounds.
    ///
    /// The bytecode backend currently assumes Excel's default 1,048,576 x 16,384 grid when
    /// lowering whole-row / whole-column references and validating range sizes.
    NonDefaultSheetDimensions,
}

/// Aggregate bytecode compilation coverage statistics.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct BytecodeCompileStats {
    /// Number of cells in the workbook that currently contain formulas.
    pub total_formula_cells: usize,
    /// Number of formula cells compiled to bytecode.
    pub compiled: usize,
    /// Number of formula cells that fell back to the AST evaluator.
    pub fallback: usize,
    /// Breakdown of fallback reasons.
    pub fallback_reasons: HashMap<BytecodeCompileReason, usize>,
}

/// Per-cell bytecode compilation report entry (only includes fallbacks).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BytecodeCompileReportEntry {
    pub sheet: String,
    pub addr: CellAddr,
    pub reason: BytecodeCompileReason,
}
