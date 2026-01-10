mod parse;
mod write;
mod xlsx;

use formula_model::rich_text::RichText;

pub use parse::parse_shared_strings_xml;
pub use parse::SharedStringsError;
pub use write::write_shared_strings_xml;
pub use write::WriteSharedStringsError;
pub use xlsx::SharedStringsXlsxError;
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
