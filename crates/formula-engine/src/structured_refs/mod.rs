mod parser;
mod resolver;
mod types;

pub use parser::parse_structured_ref;
pub use resolver::resolve_structured_ref;
pub use types::{StructuredColumns, StructuredRef, StructuredRefItem};
