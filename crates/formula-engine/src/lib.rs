#![forbid(unsafe_code)]

//! Locale-aware formula parsing and (re-)stringification, plus a core evaluation
//! engine.
//!
//! Formulas are persisted in a canonical (Excel / en-US) form and can be
//! translated for display using [`locale::localize_formula`] and
//! [`locale::canonicalize_formula`].

pub mod date;
pub mod display;
pub mod error;
pub mod functions;
pub mod graph;
pub mod locale;
pub mod pivot;

pub use crate::error::{ExcelError, ExcelResult};
pub mod what_if;
pub mod solver;

pub mod eval;
pub mod value;

mod engine;

pub use engine::{Engine, EngineError, RecalcMode};
pub use value::{ErrorKind, Value};
pub mod debug;
