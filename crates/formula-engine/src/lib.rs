#![forbid(unsafe_code)]

//! Locale-aware formula parsing and (re-)stringification, plus a core evaluation engine.
//!
//! Formulas are persisted in a canonical (Excel / en-US) form and can be translated
//! for display using [`locale::localize_formula`] and [`locale::canonicalize_formula`].
//!
//! For editor workflows (syntax highlighting, autocomplete, shared formulas) the
//! crate also exposes a syntax-only lexer/parser that produces a normalized AST via
//! [`parse_formula`] and [`parser::parse_formula_partial`].

pub mod date;
pub mod display;
pub mod error;
pub mod eval;
pub mod functions;
pub mod graph;
pub mod locale;
pub mod pivot;
pub mod solver;
pub mod value;
pub mod what_if;

pub mod debug;
pub mod sort_filter;

mod engine;

mod ast;
pub mod parser;

pub use ast::*;
pub use error::{ExcelError, ExcelResult};
pub use engine::{Engine, EngineError, RecalcMode};
pub use parser::{
    lex, parse_formula_partial, FunctionContext, ParseContext, PartialParse, Token, TokenKind,
};
pub use value::{ErrorKind, Value};

/// Parse a formula into an [`Ast`].
///
/// The input may optionally start with `=`. Locale-specific separators are
/// controlled via [`ParseOptions`].
pub fn parse_formula(formula: &str, opts: ParseOptions) -> Result<Ast, ParseError> {
    parser::parse_formula(formula, opts)
}
