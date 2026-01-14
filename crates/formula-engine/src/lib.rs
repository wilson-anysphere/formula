#![forbid(unsafe_code)]
#![deny(unreachable_patterns)]

//! Locale-aware formula parsing and (re-)stringification, plus a core evaluation engine.
//!
//! Formulas are persisted in a canonical (Excel / en-US) form and can be translated
//! for display using [`locale::localize_formula`] and [`locale::canonicalize_formula`].
//!
//! For editor workflows (syntax highlighting, autocomplete, shared formulas) the
//! crate also exposes a syntax-only lexer/parser that produces a normalized AST via
//! [`parse_formula`] and [`parser::parse_formula_partial`].
//!
//! ## External workbook references
//!
//! The engine supports Excel-style external workbook references (e.g. `=[Book.xlsx]Sheet1!A1`) via
//! a host-provided [`ExternalValueProvider`]. Hosts attach a provider with
//! [`Engine::set_external_value_provider`].
//!
//! See [`ExternalValueProvider`] for the canonical external sheet-key format (`"[workbook]sheet"`),
//! external 3D span expansion rules (`"[workbook]Sheet1:Sheet3"`), and the required
//! [`ExternalValueProvider::sheet_order`] implementation for evaluating external 3D spans.
//!
//! ## Workbook text codepage (DBCS text functions)
//!
//! Excel's legacy DBCS ("double-byte character set") text functions depend on the workbook's
//! configured ANSI text codepage (not the host OS locale). This affects:
//! - Byte-count variants like `LENB` / `LEFTB` / `MIDB` / `RIGHTB` / `FINDB` / `SEARCHB` / `REPLACEB`
//! - Fullwidth/halfwidth conversions via `ASC` / `DBCS`
//!
//! Hosts can configure the active workbook text codepage via [`Engine::set_text_codepage`]
//! (default: 1252 / en-US). Changing this setting invalidates compiled formulas so dependent cells
//! are recalculated in automatic modes.
//!
//! Performance is a feature (see `docs/16-performance-targets.md`). This crate exposes a
//! small benchmark harness via [`run_benchmarks`] so CI can detect regressions in the core
//! parsing/evaluation/recalc paths as the engine evolves.

pub mod bytecode;
pub mod calc_settings;
pub mod coercion;
pub mod date;
pub mod debug;
pub mod display;
pub mod editing;
pub mod error;
pub mod eval;
pub mod functions;
pub mod graph;
pub mod iterative;
pub mod locale;
pub mod metadata;
pub mod pivot;
pub mod pivot_registry;
pub mod simd;
pub mod solver;
pub mod sort_filter;
pub mod style_bridge;
pub mod value;
pub mod what_if;

/// Excel's hard limit for the maximum number of arguments in a single function call.
///
/// This applies to both built-in function calls (e.g. `SUM(...)`) and call expressions
/// used for LAMBDA invocation syntax (e.g. `LAMBDA(...)(...)`).
pub const EXCEL_MAX_ARGS: usize = 255;

mod engine;
mod parallel;
mod perf;
pub mod structured_refs;

#[cfg(target_arch = "wasm32")]
mod wasm_smoke;

mod ast;
pub mod parser;

pub use crate::error::{ExcelError, ExcelResult};
pub use ast::*;
pub use editing::{
    CellChange, CellSnapshot, EditError, EditOp, EditResult, FormulaRewrite, MovedRange,
};
pub use engine::{
    BytecodeCompileReason, BytecodeCompileReportEntry, BytecodeCompileStats, Engine, EngineError,
    EngineInfo, ExternalDataProvider, ExternalValueProvider, NameDefinition, NameScope,
    PrecedentNode, RecalcMode, RecalcValueChange, SheetId, SheetLifecycleError,
};
pub use parser::{
    lex, lex_partial, parse_formula_partial, FunctionContext, ParseContext, PartialLex, PartialParse,
    Token, TokenKind,
};
pub use perf::{run_benchmarks, BenchmarkResult};
pub use value::{Entity, ErrorKind, Record, Value};

/// Parse a formula into an [`Ast`].
///
/// The input may optionally start with `=`. Locale-specific separators are
/// controlled via [`ParseOptions`].
pub fn parse_formula(formula: &str, opts: ParseOptions) -> Result<Ast, ParseError> {
    parser::parse_formula(formula, opts)
}
