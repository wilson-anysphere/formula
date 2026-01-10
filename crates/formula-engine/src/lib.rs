#![forbid(unsafe_code)]

//! Locale-aware formula parsing and (re-)stringification, plus a core evaluation
//! engine.
//!
//! The parsing surface is intentionally small for now: we only model enough
//! syntax to prove out internationalization requirements:
//! - Locale dependent argument separators (`,` vs `;`)
//! - Locale dependent decimal separators (`.` vs `,`)
//! - Localized function names (e.g. `SUMME` ↔ `SUM`)
//! - Round-tripping (localized display; canonical persistence)

pub mod date;
pub mod error;
pub mod functions;
pub mod locale;
mod parser;

pub use crate::error::{ExcelError, ExcelResult};
pub use parser::{parse_formula, Expr, Formula, ParseError};
pub mod what_if;
pub mod solver;

pub mod eval;
pub mod value;

mod engine;

pub use engine::{Engine, EngineError, RecalcMode};
pub use value::{ErrorKind, Value};
