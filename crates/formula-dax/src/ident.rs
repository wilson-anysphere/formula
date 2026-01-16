/// Escape a DAX bracket identifier (the part inside `[ ... ]`).
///
/// In DAX, a literal `]` inside a bracket identifier is escaped as `]]`.
pub(crate) fn escape_dax_bracket_identifier(ident: &str) -> String {
    if !ident.contains(']') {
        return ident.to_string();
    }

    let extra = ident.as_bytes().iter().filter(|&&b| b == b']').count();
    let mut out = String::with_capacity(ident.len() + extra);
    push_escaped_dax_bracket_identifier(ident, &mut out);
    out
}

#[cfg(feature = "pivot-model")]
use crate::DataModel;

fn push_escaped_dax_bracket_identifier(raw: &str, out: &mut String) {
    let mut start = 0usize;
    for (i, ch) in raw.char_indices() {
        if ch != ']' {
            continue;
        }

        out.push_str(&raw[start..i]);
        out.push_str("]]");
        start = i + 1; // `]` is a single-byte UTF-8 codepoint.
    }
    out.push_str(&raw[start..]);
}

#[cfg(feature = "pivot-model")]
fn push_escaped_dax_single_quotes(raw: &str, out: &mut String) {
    let mut start = 0usize;
    for (i, ch) in raw.char_indices() {
        if ch != '\'' {
            continue;
        }

        out.push_str(&raw[start..i]);
        out.push_str("''");
        start = i + 1; // `'` is a single-byte UTF-8 codepoint.
    }
    out.push_str(&raw[start..]);
}

/// Format a DAX table name as a single-quoted identifier.
///
/// Table names are always quoted to avoid edge cases with spaces and reserved words. A literal
/// single quote inside the name is escaped as `''`.
#[cfg(feature = "pivot-model")]
pub(crate) fn format_dax_table_name(table: &str) -> String {
    let extra = table.as_bytes().iter().filter(|&&b| b == b'\'').count();
    let mut out = String::with_capacity(table.len() + extra + 2);
    out.push('\'');
    push_escaped_dax_single_quotes(table, &mut out);
    out.push('\'');
    out
}

/// Format a fully-qualified DAX column reference: `'Table'[Column]`.
#[cfg(feature = "pivot-model")]
pub(crate) fn format_dax_column_ref(table: &str, column: &str) -> String {
    let extra_table = table.as_bytes().iter().filter(|&&b| b == b'\'').count();
    let extra_col = column.as_bytes().iter().filter(|&&b| b == b']').count();
    let mut out = String::with_capacity(table.len() + column.len() + 4 + extra_table + extra_col);

    out.push('\'');
    push_escaped_dax_single_quotes(table, &mut out);
    out.push('\'');
    out.push('[');
    push_escaped_dax_bracket_identifier(column, &mut out);
    out.push(']');
    out
}

/// Format a DAX measure reference: `[Measure]`.
#[cfg(feature = "pivot-model")]
pub(crate) fn format_dax_measure_ref(measure: &str) -> String {
    let name = DataModel::normalize_measure_name(measure);
    let extra = name.as_bytes().iter().filter(|&&b| b == b']').count();
    let mut out = String::with_capacity(name.len() + extra + 2);
    out.push('[');
    push_escaped_dax_bracket_identifier(name, &mut out);
    out.push(']');
    out
}
