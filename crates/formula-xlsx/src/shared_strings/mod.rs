mod parse;
pub(crate) mod preserve;
mod write;
#[cfg(not(target_arch = "wasm32"))]
mod xlsx;

use formula_model::rich_text::RichText;

pub use parse::parse_shared_strings_xml;
pub use parse::SharedStringsError;
pub use write::write_shared_strings_xml;
pub use write::WriteSharedStringsError;
#[cfg(not(target_arch = "wasm32"))]
pub use xlsx::SharedStringsXlsxError;
#[cfg(not(target_arch = "wasm32"))]
pub use xlsx::{read_shared_strings_from_xlsx, write_shared_strings_to_xlsx};

/// Shared strings table (`xl/sharedStrings.xml`).
#[derive(Clone, Debug, Default, PartialEq)]
pub struct SharedStrings {
    pub items: Vec<RichText>,
}

impl SharedStrings {
    pub fn len(&self) -> usize {
        self.items.len()
    }

    pub fn is_empty(&self) -> bool {
        self.items.is_empty()
    }
}
