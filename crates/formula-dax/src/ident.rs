use crate::DataModel;

/// Escape a DAX bracket identifier (the part inside `[ ... ]`).
///
/// In DAX, a literal `]` inside a bracket identifier is escaped as `]]`.
pub(crate) fn escape_dax_bracket_identifier(ident: &str) -> String {
    ident.replace(']', "]]")
}

/// Format a DAX table name as a single-quoted identifier.
///
/// Table names are always quoted to avoid edge cases with spaces and reserved words. A literal
/// single quote inside the name is escaped as `''`.
pub(crate) fn format_dax_table_name(table: &str) -> String {
    let escaped = table.replace('\'', "''");
    format!("'{escaped}'")
}

/// Format a fully-qualified DAX column reference: `'Table'[Column]`.
pub(crate) fn format_dax_column_ref(table: &str, column: &str) -> String {
    let escaped = escape_dax_bracket_identifier(column);
    format!("{}[{}]", format_dax_table_name(table), escaped)
}

/// Format a DAX measure reference: `[Measure]`.
pub(crate) fn format_dax_measure_ref(measure: &str) -> String {
    let name = DataModel::normalize_measure_name(measure);
    let escaped = escape_dax_bracket_identifier(name);
    format!("[{escaped}]")
}

