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
pub use evaluator::{EvalContext, Evaluator, RecalcContext, ResolvedName, ValueResolver};
pub use parser::{FormulaParseError, Parser};
