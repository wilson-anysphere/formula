//! Helpers for parsing canonical external workbook sheet keys.
//!
//! The engine represents external workbook references using a bracketed "external sheet key"
//! string such as `"[Book.xlsx]Sheet1"`. Centralizing parsing here ensures consistent validation
//! across the evaluator, engine, and debug tooling.

/// Find the end of a raw Excel external workbook prefix that starts with `[` (e.g. `[Book.xlsx]`).
///
/// Returns the index *after* the closing bracket.
///
/// Notes:
/// - Excel escapes literal `]` characters inside workbook identifiers by doubling them: `]]` -> `]`.
/// - Workbook identifiers may contain `[` characters; treat them as plain text (no nesting).
pub(crate) fn find_external_workbook_prefix_end(src: &str, start: usize) -> Option<usize> {
    formula_model::external_refs::find_external_workbook_prefix_end(src, start)
}

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

/// Parse a workbook-only external key in the canonical bracketed form: `"[Book]"`.
///
/// This is used for workbook-scoped external structured references like `[Book.xlsx]Table1[Col]`,
/// which lower to a `SheetReference::External("[Book.xlsx]")` key (no explicit sheet name).
///
/// Returns the workbook identifier slice (borrowed from `key`).
pub(crate) fn parse_external_workbook_key(key: &str) -> Option<&str> {
    if !key.starts_with('[') {
        return None;
    }
    let end = key.rfind(']')?;
    if end + 1 != key.len() {
        return None;
    }
    let workbook = &key[1..end];
    (!workbook.is_empty()).then_some(workbook)
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
    fn find_external_workbook_prefix_end_parses_escaped_brackets() {
        let src = "[Book]]Name.xlsx]Sheet1";
        let end = find_external_workbook_prefix_end(src, 0).expect("end");
        assert_eq!(&src[..end], "[Book]]Name.xlsx]");
        assert_eq!(&src[end..], "Sheet1");
    }

    #[test]
    fn find_external_workbook_prefix_end_requires_leading_open_bracket() {
        assert_eq!(find_external_workbook_prefix_end("Book.xlsx]Sheet1", 0), None);
        assert_eq!(find_external_workbook_prefix_end("[]Sheet1", 0), Some(2));
    }

    #[test]
    fn parse_external_workbook_key_parses_workbook_only() {
        let workbook = parse_external_workbook_key("[Book.xlsx]").expect("parse");
        assert_eq!(workbook, "Book.xlsx");
    }

    #[test]
    fn parse_external_workbook_key_accepts_bracketed_directories_in_paths() {
        let workbook = parse_external_workbook_key("[C:\\[foo]\\Book.xlsx]").expect("parse");
        assert_eq!(workbook, "C:\\[foo]\\Book.xlsx");
    }

    #[test]
    fn parse_external_workbook_key_rejects_missing_or_empty_workbook() {
        assert!(parse_external_workbook_key("[Book.xlsx").is_none());
        assert!(parse_external_workbook_key("[]").is_none());
    }

    #[test]
    fn parse_external_workbook_key_rejects_sheet_qualified_keys() {
        assert!(parse_external_workbook_key("[Book.xlsx]Sheet1").is_none());
    }

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
    fn parse_external_key_accepts_escaped_close_brackets_in_workbook() {
        let (workbook, sheet) = parse_external_key("[Book]]Name.xlsx]Sheet1").expect("parse");
        assert_eq!(workbook, "Book]]Name.xlsx");
        assert_eq!(sheet, "Sheet1");
    }

    #[test]
    fn parse_external_workbook_key_accepts_escaped_close_brackets() {
        let workbook = parse_external_workbook_key("[Book]]Name.xlsx]").expect("parse");
        assert_eq!(workbook, "Book]]Name.xlsx");
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
