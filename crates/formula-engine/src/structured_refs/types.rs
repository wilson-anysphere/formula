#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum StructuredRefItem {
    All,
    Data,
    Headers,
    Totals,
    ThisRow,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum StructuredColumn {
    Single(String),
    Range { start: String, end: String },
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum StructuredColumns {
    All,
    Single(String),
    Range {
        start: String,
        end: String,
    },
    /// A union of non-contiguous column selections (single columns and/or ranges).
    Multi(Vec<StructuredColumn>),
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct StructuredRef {
    pub table_name: Option<String>,
    /// Special table item specifiers like `#Headers`/`#Data`.
    ///
    /// An empty list means the default item selection (Excel's implicit `#Data`).
    pub items: Vec<StructuredRefItem>,
    pub columns: StructuredColumns,
}
