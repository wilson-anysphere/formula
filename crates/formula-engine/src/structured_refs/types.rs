#[derive(Debug, Clone, PartialEq, Eq)]
pub enum StructuredRefItem {
    All,
    Data,
    Headers,
    Totals,
    ThisRow,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum StructuredColumns {
    All,
    Single(String),
    Range { start: String, end: String },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StructuredRef {
    pub table_name: Option<String>,
    pub item: Option<StructuredRefItem>,
    pub columns: StructuredColumns,
}

