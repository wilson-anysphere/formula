mod parser;
mod resolver;
mod types;

pub use parser::parse_structured_ref;
pub use resolver::{resolve_structured_ref, resolve_structured_ref_in_table};
pub use types::{StructuredColumn, StructuredColumns, StructuredRef, StructuredRefItem};
