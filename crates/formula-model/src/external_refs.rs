//! Helpers for parsing Excel external workbook prefixes.
//!
//! Excel external workbook references embed a workbook identifier inside brackets, e.g.
//! `[Book.xlsx]Sheet1!A1`. Literal `]` characters inside the workbook identifier are escaped by
//! doubling them (`]]`).

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
}

