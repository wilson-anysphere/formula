//! Structured references (`Table1[Column]`, `Table1[[#Headers],[Column]]`, …) are used in two
//! distinct layers:
//!
//! - **Parser/Lexer disambiguation**: find the correct closing `]` when `]]` may mean either
//!   “escaped `]`” or “close nested group”. The common fast path is allocation-free; ambiguous
//!   `]]` cases may allocate small temporary buffers while selecting the correct close.
//! - **Evaluation/Compilation**: parse into a `StructuredRef` (items + column selection) so it can
//!   be resolved against table metadata.
//!
//! Prefer:
//! - `find_structured_ref_end` / `scan_structured_ref` for disambiguation (hot path)
//! - `parse_structured_ref_parts{,_unchecked}` when you already have `(table_name, spec)`
//! - `parse_structured_ref` when parsing from raw formula text starting at a byte index

mod parser;
mod resolver;
mod types;

pub use parser::parse_structured_ref;
pub(crate) use parser::find_structured_ref_end;
pub(crate) use parser::find_structured_ref_end_lenient;
pub(crate) use parser::parse_structured_ref_parts;
pub(crate) use parser::parse_structured_ref_parts_unchecked;
pub(crate) use parser::scan_structured_ref;
pub use resolver::{resolve_structured_ref, resolve_structured_ref_in_table};
pub use types::{StructuredColumn, StructuredColumns, StructuredRef, StructuredRefItem};
