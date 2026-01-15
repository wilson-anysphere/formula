//! Helpers for parsing canonical external workbook sheet keys.
//!
//! The engine represents external workbook references using a bracketed "external sheet key"
//! string such as `"[Book.xlsx]Sheet1"`. Centralizing parsing here ensures consistent validation
//! across the evaluator, engine, and debug tooling.

/// Split an external workbook key on the bracketed workbook boundary.
///
/// Accepts both single-sheet keys (`"[Book]Sheet"`) and 3D span keys (`"[Book]Start:End"`).
/// The returned `sheet_part` is everything after the closing bracket (it may contain `:`).
pub(crate) fn split_external_sheet_key_parts(key: &str) -> Option<(&str, &str)> {
    if !key.starts_with('[') {
        return None;
    }

    // External workbook ids can include a path prefix (e.g. from a quoted reference like
    // `'C:\[foo]\[Book.xlsx]Sheet1'!A1`), and that prefix may itself contain `[` / `]`. Locate
    // the *last* closing bracket to recover the full workbook id.
    let end = key.rfind(']')?;
    let workbook = &key[1..end];
    let sheet_part = &key[end + 1..];

    if workbook.is_empty() || sheet_part.is_empty() {
        return None;
    }

    Some((workbook, sheet_part))
}

/// Parse an external workbook sheet key in the canonical bracketed form: `"[Book]Sheet"`.
///
/// Returns the workbook name and sheet name slices (borrowed from `key`).
///
/// Notes:
/// - External 3D spans (`"[Book]Sheet1:Sheet3"`) are **not** accepted here; use
///   [`parse_external_span_key`] instead.
pub(crate) fn parse_external_key(key: &str) -> Option<(&str, &str)> {
    let (workbook, sheet) = split_external_sheet_key_parts(key)?;

    // `:` indicates an external workbook 3D span like `"[Book]Sheet1:Sheet3"`. Those are parsed
    // separately by [`parse_external_span_key`].
    if sheet.contains(':') {
        return None;
    }

    Some((workbook, sheet))
}

/// Parse an external workbook 3D span key in the canonical bracketed form: `"[Book]Start:End"`.
///
/// Returns the workbook name, start sheet, and end sheet slices (borrowed from `key`).
pub(crate) fn parse_external_span_key(key: &str) -> Option<(&str, &str, &str)> {
    let (workbook, sheet_part) = split_external_sheet_key_parts(key)?;

    let (start, end) = sheet_part.split_once(':')?;
    if start.is_empty() || end.is_empty() {
        return None;
    }

    // Sheet names cannot contain `:` in Excel, so reject additional separators to avoid
    // ambiguity.
    if end.contains(':') {
        return None;
    }

    Some((workbook, start, end))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_external_key_accepts_spaces_and_hyphens_in_workbook() {
        let (workbook, sheet) = parse_external_key("[My Book-1.xlsx]Sheet1").expect("parse");
        assert_eq!(workbook, "My Book-1.xlsx");
        assert_eq!(sheet, "Sheet1");
    }

    #[test]
    fn parse_external_key_accepts_spaces_in_sheet() {
        let (workbook, sheet) = parse_external_key("[Book.xlsx]My Sheet").expect("parse");
        assert_eq!(workbook, "Book.xlsx");
        assert_eq!(sheet, "My Sheet");
    }

    #[test]
    fn parse_external_key_uses_last_closing_bracket_for_workbook_id() {
        // Workbook ids can contain `[` / `]` in a path prefix, so we must locate the *last* `]`.
        let (workbook, sheet) = parse_external_key("[C:\\[foo]\\Book.xlsx]Sheet1").expect("parse");
        assert_eq!(workbook, "C:\\[foo]\\Book.xlsx");
        assert_eq!(sheet, "Sheet1");
    }

    #[test]
    fn parse_external_span_key_parses_start_and_end_sheets() {
        let (workbook, start, end) =
            parse_external_span_key("[Book.xlsx]Sheet1:Sheet3").expect("parse");
        assert_eq!(workbook, "Book.xlsx");
        assert_eq!(start, "Sheet1");
        assert_eq!(end, "Sheet3");
    }

    #[test]
    fn parse_external_key_rejects_missing_bracket() {
        assert!(parse_external_key("[Book.xlsxSheet1").is_none());
        assert!(parse_external_span_key("[Book.xlsxSheet1:Sheet3").is_none());
    }

    #[test]
    fn parse_external_key_rejects_empty_workbook_or_sheet() {
        assert!(parse_external_key("[]Sheet1").is_none());
        assert!(parse_external_key("[Book.xlsx]").is_none());
    }

    #[test]
    fn parse_external_key_rejects_spans() {
        assert!(parse_external_key("[Book.xlsx]Sheet1:Sheet3").is_none());
        assert!(parse_external_span_key("[Book.xlsx]Sheet1:Sheet3").is_some());
    }
}
