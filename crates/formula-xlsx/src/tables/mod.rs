mod xml;

pub use xml::{parse_table, write_table_xml};

/// Relationship type for table parts from worksheet rels.
pub const TABLE_REL_TYPE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

#[derive(Debug, Clone)]
pub struct TablePart {
    pub r_id: String,
}

