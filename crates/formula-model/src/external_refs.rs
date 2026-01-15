//! Helpers for external workbook references.
//!
//! Excel external workbook references embed a workbook identifier inside brackets, e.g.
//! `[Book.xlsx]Sheet1!A1`. Literal `]` characters inside the workbook identifier are escaped by
//! doubling them (`]]`).
//!
//! This module provides utilities for both:
//! - scanning/splitting raw formula text segments like `C:\path\[Book.xlsx]Sheet1`
//! - parsing canonical bracketed keys like `"[Book.xlsx]Sheet1"` that appear in internal data
//!   structures

/// Find the end of a raw Excel external workbook prefix that starts with `[` (e.g. `[Book.xlsx]`).
///
/// Returns the index *after* the closing bracket.
///
/// Notes:
/// - Excel escapes literal `]` characters inside workbook identifiers by doubling them: `]]` -> `]`.
/// - Workbook identifiers may contain `[` characters; treat them as plain text (no nesting).
pub fn find_external_workbook_prefix_end(src: &str, start: usize) -> Option<usize> {
    let bytes = src.as_bytes();
    if bytes.get(start) != Some(&b'[') {
        return None;
    }

    let mut i = start + 1;
    while i < bytes.len() {
        if bytes[i] == b']' {
            if bytes.get(i + 1) == Some(&b']') {
                i += 2;
                continue;
            }
            return Some(i + 1);
        }

        // Advance by UTF-8 char boundaries so we don't accidentally interpret `[` / `]` bytes
        // inside multi-byte sequences as actual bracket characters.
        let ch = src[i..].chars().next()?;
        i += ch.len_utf8();
    }

    None
}

fn find_external_workbook_delimiter_end(raw: &str, start: usize) -> Option<usize> {
    let bytes = raw.as_bytes();
    if bytes.get(start) != Some(&b'[') {
        return None;
    }

    let mut i = start + 1;
    while i < bytes.len() {
        if bytes[i] == b']' {
            // Escaped literal `]` inside bracketed segments: `]]` -> `]`.
            if bytes.get(i + 1) == Some(&b']') {
                i += 2;
                continue;
            }

            let end = i + 1;
            if end >= raw.len() {
                return None;
            }

            // Heuristic: reject candidates that still look like path segments (e.g. the `]` in a
            // bracketed directory like `C:\[foo]\[Book.xlsx]Sheet1`).
            let remainder = raw[end..].trim_start();
            let next = remainder.chars().next()?;
            if matches!(next, '\\' | '/' | '[' | ']' | '!') {
                i += 1;
                continue;
            }

            return Some(end);
        }

        // Advance by UTF-8 char boundaries so we don't accidentally interpret `[` / `]` bytes
        // inside multi-byte sequences as actual bracket characters.
        let ch = raw[i..].chars().next()?;
        i += ch.len_utf8();
    }

    None
}

/// Split a raw sheet spec/workbook reference on its external workbook prefix (if present).
///
/// For example:
/// - `"[Book.xlsx]Sheet1"` -> `Some(("[Book.xlsx]", "Sheet1"))`
/// - `"C:\\path\\[Book.xlsx]Sheet1"` -> `Some(("C:\\path\\[Book.xlsx]", "Sheet1"))`
/// - `"Sheet1"` -> `None`
///
/// Note:
/// - External references can include path components that contain `[` / `]` (e.g.
///   `'C:\[foo]\[Book.xlsx]Sheet1'!A1`). In such cases, the workbook delimiter is the last
///   bracketed segment in the spec, so this routine scans for the last `[...]` segment and splits
///   there.
pub fn split_external_workbook_prefix(raw: &str) -> Option<(&str, &str)> {
    let bytes = raw.as_bytes();
    let mut best_end: Option<usize> = None;

    let mut i = 0usize;
    while i < bytes.len() {
        if bytes[i] == b'[' {
            if let Some(end) = find_external_workbook_delimiter_end(raw, i) {
                best_end = Some(end);
            }
        }

        // Advance by UTF-8 char boundaries so we don't accidentally interpret `[` / `]` bytes
        // inside multi-byte sequences as actual bracket characters.
        let ch = raw[i..].chars().next()?;
        i += ch.len_utf8();
    }

    let end = best_end?;
    let (prefix, remainder) = raw.split_at(end);
    (!remainder.is_empty()).then_some((prefix, remainder))
}

/// Split a canonical external sheet key on the workbook boundary.
///
/// Canonical external keys are the bracketed form used by internal data structures, e.g.
/// - `"[Book.xlsx]Sheet1"`
/// - `"[Book.xlsx]Sheet1:Sheet3"`
///
/// The returned `workbook` slice is the workbook identifier without the surrounding brackets, and
/// `sheet_part` is everything after the closing bracket (it may contain `:` for 3D spans).
pub fn split_external_sheet_key_parts(key: &str) -> Option<(&str, &str)> {
    if !key.starts_with('[') {
        return None;
    }

    // Workbook identifiers can include bracketed path segments (e.g. `C:\[foo]\Book.xlsx`), so we
    // locate the *last* closing bracket to recover the full workbook id.
    let end = key.rfind(']')?;
    let workbook = &key[1..end];
    let sheet_part = &key[end + 1..];

    if workbook.is_empty() || sheet_part.is_empty() {
        return None;
    }

    Some((workbook, sheet_part))
}

/// Parse a workbook-only canonical external key: `"[Book]"`.
///
/// Returns the workbook identifier slice (borrowed from `key`).
pub fn parse_external_workbook_key(key: &str) -> Option<&str> {
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

/// Parse a single-sheet canonical external key: `"[Book]Sheet"`.
///
/// External 3D spans (`"[Book]Sheet1:Sheet3"`) are not accepted; use
/// [`parse_external_span_key`] instead.
pub fn parse_external_key(key: &str) -> Option<(&str, &str)> {
    let (workbook, sheet) = split_external_sheet_key_parts(key)?;
    if sheet.contains(':') {
        return None;
    }
    Some((workbook, sheet))
}

/// Parse a 3D-span canonical external key: `"[Book]Start:End"`.
pub fn parse_external_span_key(key: &str) -> Option<(&str, &str, &str)> {
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
    fn split_external_workbook_prefix_splits_prefix_and_remainder() {
        assert_eq!(
            split_external_workbook_prefix("[Book.xlsx]Sheet1"),
            Some(("[Book.xlsx]", "Sheet1"))
        );
    }

    #[test]
    fn split_external_workbook_prefix_rejects_missing_or_empty_remainder() {
        assert_eq!(split_external_workbook_prefix("Sheet1"), None);
        assert_eq!(split_external_workbook_prefix("[Book.xlsx]"), None);
    }

    #[test]
    fn split_external_workbook_prefix_accepts_paths_with_brackets() {
        assert_eq!(
            split_external_workbook_prefix("C:\\[foo]\\[Book.xlsx]Sheet1"),
            Some(("C:\\[foo]\\[Book.xlsx]", "Sheet1"))
        );
        assert_eq!(
            split_external_workbook_prefix("C:\\path\\[Book.xlsx]Sheet1"),
            Some(("C:\\path\\[Book.xlsx]", "Sheet1"))
        );
    }

    #[test]
    fn split_external_workbook_prefix_accepts_leading_bracketed_paths() {
        assert_eq!(
            split_external_workbook_prefix("[C:\\[foo]\\[Book.xlsx]Sheet1"),
            Some(("[C:\\[foo]\\[Book.xlsx]", "Sheet1"))
        );
    }

    #[test]
    fn split_external_workbook_prefix_accepts_canonical_keys_with_unescaped_brackets() {
        assert_eq!(
            split_external_workbook_prefix("[C:\\[foo]\\Book.xlsx]Sheet1"),
            Some(("[C:\\[foo]\\Book.xlsx]", "Sheet1"))
        );
    }

    #[test]
    fn split_external_sheet_key_parts_splits_workbook_and_sheet_part() {
        assert_eq!(
            split_external_sheet_key_parts("[Book.xlsx]Sheet1"),
            Some(("Book.xlsx", "Sheet1"))
        );
        assert_eq!(
            split_external_sheet_key_parts("[Book.xlsx]Sheet1:Sheet3"),
            Some(("Book.xlsx", "Sheet1:Sheet3"))
        );
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
        let (workbook, sheet) = parse_external_key("[C:\\[foo]\\Book.xlsx]Sheet1").expect("parse");
        assert_eq!(workbook, "C:\\[foo]\\Book.xlsx");
        assert_eq!(sheet, "Sheet1");
    }

    #[test]
    fn parse_external_span_key_parses_start_and_end_sheets() {
        let (workbook, start, end) = parse_external_span_key("[Book.xlsx]Sheet1:Sheet3").expect("parse");
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

