mod address;
mod ast;
mod compiler;
mod evaluator;
mod parser;

/// Internal sentinel prefix used to track omitted LAMBDA parameters.
///
/// The leading NUL character ensures the key cannot be referenced by user formulas.
pub(crate) const LAMBDA_OMITTED_PREFIX: &str = "\u{0}LAMBDA_OMITTED:";

pub use address::{parse_a1, AddressParseError, CellAddr};
pub use ast::{
    BinaryOp, CellRef, CompareOp, CompiledExpr, Expr, NameRef, ParsedExpr, PostfixOp, RangeRef,
    SheetReference, UnaryOp,
};
pub use compiler::{compile_canonical_expr, lower_ast, lower_expr};
pub(crate) use evaluator::{
    is_valid_external_sheet_key, split_external_sheet_key, split_external_sheet_span_key,
};
pub use evaluator::{
    DependencyTrace, EvalContext, Evaluator, RecalcContext, ResolvedName, ValueResolver,
};
pub(crate) use evaluator::MAX_MATERIALIZED_ARRAY_CELLS;
pub use parser::{FormulaParseError, Parser};
