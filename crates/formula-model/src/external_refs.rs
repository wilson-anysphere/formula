//! Helpers for external workbook references.
//!
//! Excel external workbook references embed a workbook identifier inside brackets, e.g.
//! `[Book.xlsx]Sheet1!A1`. Literal `]` characters inside the workbook identifier are escaped by
//! doubling them (`]]`).
//!
//! This module provides utilities for both:
//! - scanning/splitting *sheet-spec tokens* (the `Sheet` portion of `Sheet!A1`) that may contain an
//!   embedded workbook id, e.g. `C:\path\[Book.xlsx]Sheet1` or `[Book.xlsx]Sheet1`
//! - disambiguating external-workbook prefixes inside *full formula text* (where other `[...]`
//!   constructs like structured references may appear)
//! - parsing canonical bracketed keys like `"[Book.xlsx]Sheet1"` that appear in internal data
//!   structures

use std::borrow::Cow;

#[derive(Clone, Copy)]
enum WorkbookDelimiterDecision {
    Accept,
    Continue,
    Reject,
}

fn skip_ws_pos(src: &str, mut pos: usize) -> usize {
    while let Some(ch) = src.get(pos..).and_then(|s| s.chars().next()) {
        if ch.is_whitespace() {
            pos += ch.len_utf8();
        } else {
            break;
        }
    }
    pos
}

fn scan_quoted_sheet_name_end(src: &str, start: usize) -> Option<usize> {
    if !src.get(start..)?.starts_with('\'') {
        return None;
    }
    let mut pos = start + 1;
    loop {
        match src[pos..].chars().next() {
            Some('\'') => {
                if src[pos..].starts_with("''") {
                    pos += 2;
                    continue;
                }
                return Some(pos + 1);
            }
            Some(ch) => pos += ch.len_utf8(),
            None => return None,
        }
    }
}

fn is_unquoted_name_start_char(ch: char) -> bool {
    ch == '_' || ch.is_alphabetic()
}

fn is_unquoted_name_cont_char(ch: char) -> bool {
    ch.is_alphanumeric() || matches!(ch, '_' | '.' | '$')
}

fn scan_unquoted_name_end(src: &str, start: usize) -> Option<usize> {
    let first = src[start..].chars().next()?;
    if !is_unquoted_name_start_char(first) {
        return None;
    }
    let mut i = start + first.len_utf8();
    while i < src.len() {
        let ch = src[i..].chars().next()?;
        if is_unquoted_name_cont_char(ch) {
            i += ch.len_utf8();
            continue;
        }
        break;
    }
    Some(i)
}

fn scan_sheet_name_token_end(src: &str, start: usize) -> Option<usize> {
    let i = skip_ws_pos(src, start);
    if i >= src.len() {
        return None;
    }
    match src[i..].chars().next()? {
        '\'' => scan_quoted_sheet_name_end(src, i),
        _ => scan_unquoted_name_end(src, i),
    }
}

fn is_valid_sheet_spec_remainder(remainder: &str) -> bool {
    let mut pos = 0usize;
    let Some(first_end) = scan_sheet_name_token_end(remainder, pos) else {
        return false;
    };
    pos = skip_ws_pos(remainder, first_end);
    if pos == remainder.len() {
        return true;
    }
    if remainder[pos..].starts_with(':') {
        pos = skip_ws_pos(remainder, pos + 1);
        let Some(second_end) = scan_sheet_name_token_end(remainder, pos) else {
            return false;
        };
        pos = skip_ws_pos(remainder, second_end);
        return pos == remainder.len();
    }
    false
}

/// Validate a workbook-scoped name remainder in a *sheet-spec token* context.
///
/// This is used when splitting sheet-spec-like tokens (e.g. `C:\path\[Book.xlsx]Sheet1`) where the
/// workbook prefix delimiter selection should be based on whether the remainder could be:
/// - a sheet name token (handled elsewhere), or
/// - a workbook-scoped name token (e.g. `[Book.xlsx]MyName` or `[Book.xlsx]Table1[Col]`).
///
/// In this context, we require the remainder to end after the name token (optionally followed by a
/// structured-ref `[...]` suffix), because there are no formula operators.
fn is_valid_workbook_scoped_name_remainder_in_sheet_spec(remainder: &str) -> bool {
    let pos = skip_ws_pos(remainder, 0);
    let Some(end) = scan_unquoted_name_end(remainder, pos) else {
        return false;
    };
    let end = skip_ws_pos(remainder, end);
    end == remainder.len() || remainder[end..].starts_with('[')
}

/// Check whether a remainder in *formula text* starts with a workbook-scoped name token.
///
/// In formula text, workbook-scoped names like `[Book.xlsx]MyName` may be immediately followed by
/// operators (e.g. `=[Book.xlsx]MyName+1`), so we only require that a valid name token starts here.
fn starts_with_workbook_scoped_name_token_in_formula_text(remainder: &str) -> bool {
    let pos = skip_ws_pos(remainder, 0);
    scan_unquoted_name_end(remainder, pos).is_some()
}

fn scan_external_workbook_delimiter_end<F>(raw: &str, start: usize, mut decide: F) -> Option<usize>
where
    F: FnMut(usize) -> WorkbookDelimiterDecision,
{
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
            match decide(end) {
                WorkbookDelimiterDecision::Accept => return Some(end),
                WorkbookDelimiterDecision::Continue => {
                    i += 1;
                    continue;
                }
                WorkbookDelimiterDecision::Reject => return None,
            }
        }

        // Advance by UTF-8 char boundaries so we don't accidentally interpret `[` / `]` bytes
        // inside multi-byte sequences as actual bracket characters.
        let ch = raw[i..].chars().next()?;
        i += ch.len_utf8();
    }

    None
}

/// Find the end of a raw Excel external workbook prefix that starts with `[` (e.g. `[Book.xlsx]`).
///
/// Returns the index *after* the closing bracket.
///
/// Notes:
/// - Excel escapes literal `]` characters inside workbook identifiers by doubling them: `]]` -> `]`.
/// - Workbook identifiers may contain `[` characters; treat them as plain text (no nesting).
/// - This function does **not** validate what follows the closing `]`. If you are scanning full
///   formula text, prefer [`find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token`]
///   to disambiguate from other `[...]` constructs (like structured references).
pub fn find_external_workbook_prefix_end(src: &str, start: usize) -> Option<usize> {
    scan_external_workbook_delimiter_end(src, start, |_| WorkbookDelimiterDecision::Accept)
}

/// Escape a workbook identifier for use inside an Excel external workbook prefix.
///
/// In Excel formula syntax, literal `]` characters inside the bracketed workbook identifier are
/// escaped by doubling them (`]]`).
///
/// This helper is intentionally small and allocation-aware: it borrows when no escaping is needed.
///
/// Despite the workbook-centric name in surrounding docs, this escaping rule is also used by other
/// bracketed identifier constructs in Excel formula text (notably structured references), so the
/// API is intentionally generic.
pub fn escape_bracketed_identifier_content(raw: &str) -> Cow<'_, str> {
    if !raw.contains(']') {
        return Cow::Borrowed(raw);
    }

    let extra = raw.as_bytes().iter().filter(|&&b| b == b']').count();
    let mut out = String::new();
    if out.try_reserve_exact(raw.len().saturating_add(extra)).is_err() {
        debug_assert!(
            false,
            "allocation failed (escape bracketed identifier, len={})",
            raw.len()
        );
    }
    push_escaped_bracketed_identifier_content(raw, &mut out);
    Cow::Owned(out)
}

/// Unescape bracketed identifier content from Excel formula text.
///
/// Excel represents literal `]` characters inside some bracketed identifier contexts by doubling
/// them (`]]`). This helper performs the inverse mapping (`]]` -> `]`) and borrows when no
/// unescaping is needed.
pub fn unescape_bracketed_identifier_content(escaped: &str) -> Cow<'_, str> {
    if !escaped.contains("]]") {
        return Cow::Borrowed(escaped);
    }

    let mut out = String::new();
    if out.try_reserve_exact(escaped.len()).is_err() {
        debug_assert!(
            false,
            "allocation failed (unescape bracketed identifier, len={})",
            escaped.len()
        );
    }
    let mut chars = escaped.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == ']' && chars.peek() == Some(&']') {
            chars.next();
            out.push(']');
        } else {
            out.push(ch);
        }
    }
    Cow::Owned(out)
}

pub fn push_escaped_bracketed_identifier_content(raw: &str, out: &mut String) {
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

/// Escape a workbook identifier for use inside an Excel external workbook prefix.
///
/// This is a small semantic wrapper around [`escape_bracketed_identifier_content`].
pub fn escape_external_workbook_name_for_prefix(workbook: &str) -> Cow<'_, str> {
    escape_bracketed_identifier_content(workbook)
}

/// Find the end of a bracketed external workbook prefix (`[Book.xlsx]`) in formula text.
///
/// This helper disambiguates external workbook prefixes from other `[...]` constructs (like
/// structured references) while accounting for workbook identifiers that contain unescaped `]`
/// characters (typically from bracketed path components).
///
/// Selection rules:
/// - Prefer the first unescaped `]` that can be parsed as a sheet-qualified prefix
///   (`[workbook]Sheet!` or `[workbook]Sheet1:Sheet3!`).
/// - Otherwise, return the *last* unescaped `]` that is plausibly followed by the start of a
///   workbook-scoped name token (`[workbook]Name...`). In formula text, the name may be followed by
///   operators (e.g. `=[workbook]Name+1`) or structured-ref suffixes (e.g. `[workbook]Table1[Col]`),
///   so this function does not require the remainder to be "token-complete".
pub fn find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token(
    src: &str,
    start: usize,
) -> Option<usize> {
    let mut best_name_end: Option<usize> = None;
    let sheet_end = scan_external_workbook_delimiter_end(src, start, |end| {
        // Prefer a sheet-qualified prefix (`[workbook]Sheet!` or `[workbook]Sheet1:Sheet3!`).
        let after_end = skip_ws_pos(src, end);
        if let Some(mut sheet_end) = scan_sheet_name_token_end(src, after_end) {
            sheet_end = skip_ws_pos(src, sheet_end);
            if sheet_end < src.len() && src[sheet_end..].starts_with(':') {
                sheet_end = skip_ws_pos(src, sheet_end + 1);
                let Some(second_end) = scan_sheet_name_token_end(src, sheet_end) else {
                    // This wasn't a valid sheet span; don't accept based on a bare `!` after `:`.
                    // Treat the delimiter as an interior `]` and continue scanning for a later one.
                    return WorkbookDelimiterDecision::Continue;
                };
                sheet_end = skip_ws_pos(src, second_end);
            }

            if sheet_end < src.len() && src[sheet_end..].starts_with('!') {
                return WorkbookDelimiterDecision::Accept;
            }
        }

        // Otherwise, treat this as a workbook-scoped name prefix `[workbook]Name`.
        if starts_with_workbook_scoped_name_token_in_formula_text(&src[end..]) {
            best_name_end = Some(end);
        }

        WorkbookDelimiterDecision::Continue
    });

    sheet_end.or(best_name_end)
}

fn find_external_workbook_delimiter_end(raw: &str, start: usize) -> Option<usize> {
    scan_external_workbook_delimiter_end(raw, start, |end| {
        if end >= raw.len() {
            return WorkbookDelimiterDecision::Reject;
        }

        let remainder = raw[end..].trim_start();
        if remainder.is_empty() {
            return WorkbookDelimiterDecision::Reject;
        }

        // Accept candidates that yield a plausible remainder (sheet spec or workbook-scoped name).
        //
        // When the bracketed segment starts at byte 0, we may be parsing the "canonical" form where
        // the workbook id itself contains unescaped `]` from bracketed path components (e.g.
        // `[C:\[foo]\Book.xlsx]Sheet1`). In that case, treat rejected delimiters as interior and
        // keep scanning for a later `]` that yields a plausible remainder.
        //
        // For bracketed segments that start later in the string (e.g. `C:\[foo]\[Book.xlsx]Sheet1`),
        // a rejected candidate is more likely to be a bracketed directory component. In that case,
        // don't "extend" the segment past the first `]`; let the outer scanner consider later
        // `[...]` segments instead.
        if is_valid_sheet_spec_remainder(remainder)
            || is_valid_workbook_scoped_name_remainder_in_sheet_spec(remainder)
        {
            return WorkbookDelimiterDecision::Accept;
        }
        if start == 0 {
            return WorkbookDelimiterDecision::Continue;
        }
        WorkbookDelimiterDecision::Reject
    })
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
/// - This operates on a *sheet-spec token* or workbook-scoped token, not on arbitrary full formula
///   text.
pub fn split_external_workbook_prefix(raw: &str) -> Option<(&str, &str)> {
    let end = find_external_workbook_prefix_span_in_sheet_spec(raw)?.end;
    let (prefix, remainder) = raw.split_at(end);
    (!remainder.is_empty()).then_some((prefix, remainder))
}

/// Find the bracketed external workbook segment within a raw sheet spec/workbook reference.
///
/// Returns the byte range spanning the selected `[workbook]` segment, including the brackets.
///
/// This is useful for consumers that need to remove the brackets while preserving any path prefix
/// that precedes the workbook segment.
///
/// Important: this operates on a *sheet spec token* (e.g. `C:\path\[Book.xlsx]Sheet1` or
/// `[Book.xlsx]Sheet1`) or a workbook-scoped token (e.g. `[Book.xlsx]MyName`). Do **not** run this
/// over arbitrary formula text: formulas can contain other `[...]` constructs (structured
/// references, field access) that are not external workbook prefixes.
pub fn find_external_workbook_prefix_span_in_sheet_spec(
    raw: &str,
) -> Option<std::ops::Range<usize>> {
    let bytes = raw.as_bytes();
    let mut i = 0usize;
    let mut best: Option<(usize, usize)> = None; // (open, end) where end is exclusive of the closing `]`

    while i < bytes.len() {
        if bytes[i] == b'[' {
            if let Some(end) = find_external_workbook_delimiter_end(raw, i) {
                best = match best {
                    None => Some((i, end)),
                    Some((best_start, best_end)) => {
                        if end > best_end || (end == best_end && i < best_start) {
                            Some((i, end))
                        } else {
                            Some((best_start, best_end))
                        }
                    }
                };

                // Skip the bracketed segment to avoid misclassifying `[` characters inside the
                // workbook identifier as the start of a new prefix.
                i = end;
                continue;
            }
        }

        // Advance by UTF-8 char boundaries so we don't accidentally interpret `[` / `]` bytes
        // inside multi-byte sequences as actual bracket characters.
        let ch = raw[i..].chars().next()?;
        i += ch.len_utf8();
    }

    best.map(|(start, end)| start..end)
}

/// Backward-compatible alias for [`find_external_workbook_prefix_span_in_sheet_spec`].
///
/// Prefer [`find_external_workbook_prefix_span_in_sheet_spec`]; this name is ambiguous and is easy
/// to misuse on full formula text.
pub fn find_external_workbook_prefix_span(raw: &str) -> Option<std::ops::Range<usize>> {
    find_external_workbook_prefix_span_in_sheet_spec(raw)
}

/// Parse an Excel-style path-qualified external workbook sheet key that has been lexed as a single
/// string (typically via quoted sheet syntax), e.g. `C:\path\[Book.xlsx]Sheet1`.
///
/// Returns `(workbook, sheet_part)` where `workbook` has no surrounding brackets and `sheet_part`
/// is everything after the closing `]` (it may include a `:` for 3D spans).
pub fn parse_path_qualified_external_sheet_key(raw: &str) -> Option<(String, String)> {
    if raw.starts_with('[') {
        return None;
    }

    let span = find_external_workbook_prefix_span_in_sheet_spec(raw)?;
    if span.start == 0 || span.start + 1 >= span.end {
        return None;
    }

    let book = &raw[span.start + 1..span.end - 1];
    let sheet_part = &raw[span.end..];
    if book.is_empty() || sheet_part.is_empty() {
        return None;
    }

    let prefix = &raw[..span.start];
    let mut workbook = String::new();
    if workbook
        .try_reserve_exact(prefix.len().saturating_add(book.len()))
        .is_err()
    {
        debug_assert!(
            false,
            "allocation failed (external sheet key workbook, len={})",
            raw.len()
        );
        return None;
    }
    workbook.push_str(prefix);
    workbook.push_str(book);
    Some((workbook, sheet_part.to_string()))
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

/// Format a workbook-only canonical external key: `"[Book]"`.
pub fn format_external_workbook_key(workbook: &str) -> String {
    let mut out = String::new();
    if out.try_reserve_exact(workbook.len().saturating_add(2)).is_err() {
        debug_assert!(
            false,
            "allocation failed (format external workbook key, len={})",
            workbook.len()
        );
    }
    out.push('[');
    out.push_str(workbook);
    out.push(']');
    out
}

/// Format a single-sheet canonical external key: `"[Book]Sheet"`.
pub fn format_external_key(workbook: &str, sheet: &str) -> String {
    let mut out = String::new();
    if out
        .try_reserve_exact(workbook.len().saturating_add(sheet.len()).saturating_add(2))
        .is_err()
    {
        debug_assert!(
            false,
            "allocation failed (format external key, workbook_len={}, sheet_len={})",
            workbook.len(),
            sheet.len()
        );
    }
    out.push('[');
    out.push_str(workbook);
    out.push(']');
    out.push_str(sheet);
    out
}

/// Format a 3D-span canonical external key: `"[Book]Start:End"`.
pub fn format_external_span_key(workbook: &str, start: &str, end: &str) -> String {
    let mut out = String::new();
    if out
        .try_reserve_exact(
            workbook
                .len()
                .saturating_add(start.len())
                .saturating_add(end.len())
                .saturating_add(3),
        )
        .is_err()
    {
        debug_assert!(
            false,
            "allocation failed (format external span key, workbook_len={}, start_len={}, end_len={})",
            workbook.len(),
            start.len(),
            end.len()
        );
    }
    out.push('[');
    out.push_str(workbook);
    out.push(']');
    out.push_str(start);
    out.push(':');
    out.push_str(end);
    out
}

/// Expand an external workbook 3D sheet span into per-sheet canonical external keys.
///
/// Given:
/// - a `workbook` identifier (no surrounding brackets),
/// - span endpoints `start` and `end` (sheet names),
/// - the external workbook's sheet names in tab order (sheet names only; no `[workbook]` prefix),
///
/// returns canonical external sheet keys like `"[Book.xlsx]Sheet2"` for each sheet in the span.
///
/// Notes:
/// - Endpoint matching uses Excel-like Unicode-aware, case-insensitive comparison via
///   [`crate::sheet_name_eq_case_insensitive`].
/// - If either endpoint is missing from `sheet_names`, returns `None`.
pub fn expand_external_sheet_span_from_order(
    workbook: &str,
    start: &str,
    end: &str,
    sheet_names: &[String],
) -> Option<Vec<String>> {
    let mut start_idx: Option<usize> = None;
    let mut end_idx: Option<usize> = None;
    for (idx, name) in sheet_names.iter().enumerate() {
        if start_idx.is_none() && crate::sheet_name_eq_case_insensitive(name, start) {
            start_idx = Some(idx);
        }
        if end_idx.is_none() && crate::sheet_name_eq_case_insensitive(name, end) {
            end_idx = Some(idx);
        }
        if start_idx.is_some() && end_idx.is_some() {
            break;
        }
    }

    let start_idx = start_idx?;
    let end_idx = end_idx?;
    let (lo, hi) = if start_idx <= end_idx {
        (start_idx, end_idx)
    } else {
        (end_idx, start_idx)
    };

    let count = hi - lo + 1;
    let mut out: Vec<String> = Vec::new();
    if out.try_reserve_exact(count).is_err() {
        debug_assert!(
            false,
            "allocation failed (expand external sheet span, count={count})"
        );
        return None;
    }
    for name in &sheet_names[lo..=hi] {
        out.push(format_external_key(workbook, name));
    }
    Some(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn bracketed_identifier_escape_roundtrips_with_unescape() {
        let cases = [
            "",
            "plain",
            "A]B",
            "A]]B",
            "trailing]",
            "many]]]]brackets",
        ];
        for raw in cases {
            let escaped = escape_bracketed_identifier_content(raw);
            let unescaped = unescape_bracketed_identifier_content(escaped.as_ref());
            assert_eq!(unescaped.as_ref(), raw, "roundtrip failed for raw={raw:?}");
        }
    }

    #[test]
    fn format_external_keys_roundtrip_with_parsers() {
        let workbook = "Book.xlsx";
        let sheet = "Sheet1";
        let span_start = "Sheet1";
        let span_end = "Sheet3";

        let single = format_external_key(workbook, sheet);
        assert_eq!(parse_external_key(&single), Some((workbook, sheet)));

        let workbook_only = format_external_workbook_key(workbook);
        assert_eq!(parse_external_workbook_key(&workbook_only), Some(workbook));

        let span = format_external_span_key(workbook, span_start, span_end);
        assert_eq!(
            parse_external_span_key(&span),
            Some((workbook, span_start, span_end))
        );
    }

    #[test]
    fn format_external_keys_roundtrip_with_paths_and_escapes() {
        let workbook_with_path = r"C:\dir\Book.xlsx";
        let workbook_with_escaped_rbracket = "Book]]Name.xlsx";

        let key = format_external_key(workbook_with_path, "Sheet 1");
        assert_eq!(
            parse_external_key(&key),
            Some((workbook_with_path, "Sheet 1"))
        );

        let key = format_external_key(workbook_with_escaped_rbracket, "Sheet1");
        assert_eq!(
            parse_external_key(&key),
            Some((workbook_with_escaped_rbracket, "Sheet1"))
        );

        let span = format_external_span_key(workbook_with_path, "A", "B");
        assert_eq!(
            parse_external_span_key(&span),
            Some((workbook_with_path, "A", "B"))
        );

        let wb_only = format_external_workbook_key(workbook_with_escaped_rbracket);
        assert_eq!(
            parse_external_workbook_key(&wb_only),
            Some(workbook_with_escaped_rbracket)
        );
    }

    #[test]
    fn find_external_workbook_prefix_end_parses_escaped_brackets() {
        let src = "[Book]]Name.xlsx]Sheet1";
        let end = find_external_workbook_prefix_end(src, 0).expect("end");
        assert_eq!(&src[..end], "[Book]]Name.xlsx]");
        assert_eq!(&src[end..], "Sheet1");
    }

    #[test]
    fn find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token_skips_bracketed_paths()
    {
        let src = r"[C:\[foo]\Book.xlsx]Sheet1";
        let end = find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token(src, 0)
            .expect("end");
        assert_eq!(&src[..end], r"[C:\[foo]\Book.xlsx]");
        assert_eq!(&src[end..], "Sheet1");
    }

    #[test]
    fn find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token_skips_bracketed_paths_without_separators(
    ) {
        // Some producers emit bracketed path components without escaping the inner `]`, even when
        // the path component isn't followed by a separator (e.g. `[foo]Book.xlsx`).
        let src = r"[C:\[foo]Book.xlsx]Sheet1!A1";
        let end = find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token(src, 0)
            .expect("end");
        assert_eq!(&src[..end], r"[C:\[foo]Book.xlsx]");
        assert_eq!(&src[end..], "Sheet1!A1");
    }

    #[test]
    fn find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token_handles_workbook_ids_with_lbracket(
    ) {
        let src = "[A1[Name.xlsx]Sheet1";
        let end = find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token(src, 0)
            .expect("end");
        assert_eq!(&src[..end], "[A1[Name.xlsx]");
        assert_eq!(&src[end..], "Sheet1");
    }

    #[test]
    fn find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token_accepts_operator_delimited_name(
    ) {
        // Formula-text disambiguation: workbook-scoped name refs can be followed by operators, e.g.
        // `=[Book.xlsx]MyName+1`. Ensure we still detect the workbook prefix end.
        let src = "[A1[Name.xlsx]MyName+1";
        let end = find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token(src, 0)
            .expect("end");
        assert_eq!(&src[..end], "[A1[Name.xlsx]");
        assert_eq!(&src[end..], "MyName+1");
    }

    #[test]
    fn find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token_rejects_structured_refs(
    ) {
        assert!(
            find_external_workbook_prefix_end_if_followed_by_sheet_or_name_token("[@Col2]", 0)
                .is_none()
        );
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
    fn split_external_workbook_prefix_accepts_canonical_keys_with_unescaped_brackets_without_separators(
    ) {
        assert_eq!(
            split_external_workbook_prefix(r"[C:\[foo]Book.xlsx]Sheet1"),
            Some((r"[C:\[foo]Book.xlsx]", "Sheet1"))
        );
    }

    #[test]
    fn find_external_workbook_prefix_span_prefers_outer_brackets() {
        let src = r"C:\path\[A1[Name.xlsx]Sheet1";
        assert_eq!(
            find_external_workbook_prefix_span_in_sheet_spec(src),
            Some(8..22)
        );
    }

    #[test]
    fn split_external_workbook_prefix_handles_workbook_only_external_structured_refs() {
        // Workbook-only external structured ref: `[Book]Table1[Col]`.
        // The workbook id itself can contain `[` characters; the structured-ref suffix includes
        // additional `[...]` segments that should NOT be treated as workbook prefixes.
        assert_eq!(
            split_external_workbook_prefix("[A1[Name.xlsx]Table1[Col2]"),
            Some(("[A1[Name.xlsx]", "Table1[Col2]"))
        );
    }

    #[test]
    fn parse_path_qualified_external_sheet_key_extracts_workbook_and_sheet() {
        let (workbook, sheet) =
            parse_path_qualified_external_sheet_key(r"C:\path\[Book.xlsx]Sheet1").expect("parse");
        assert_eq!(workbook, r"C:\path\Book.xlsx");
        assert_eq!(sheet, "Sheet1");
    }

    #[test]
    fn parse_path_qualified_external_sheet_key_supports_bracketed_path_components() {
        let (workbook, sheet) =
            parse_path_qualified_external_sheet_key(r"C:\[foo]\[Book.xlsx]Sheet1").expect("parse");
        assert_eq!(workbook, r"C:\[foo]\Book.xlsx");
        assert_eq!(sheet, "Sheet1");
    }

    #[test]
    fn parse_path_qualified_external_sheet_key_supports_workbook_names_containing_lbracket() {
        let (workbook, sheet) =
            parse_path_qualified_external_sheet_key(r"C:\path\[A1[Name.xlsx]Sheet1")
                .expect("parse");
        assert_eq!(workbook, r"C:\path\A1[Name.xlsx");
        assert_eq!(sheet, "Sheet1");
    }

    #[test]
    fn parse_path_qualified_external_sheet_key_supports_workbook_names_with_escaped_rbracket() {
        let (workbook, sheet) =
            parse_path_qualified_external_sheet_key(r"C:\path\[Book[Name]].xlsx]Sheet1")
                .expect("parse");
        assert_eq!(workbook, r"C:\path\Book[Name]].xlsx");
        assert_eq!(sheet, "Sheet1");
    }

    #[test]
    fn parse_path_qualified_external_sheet_key_rejects_canonical_keys() {
        assert!(parse_path_qualified_external_sheet_key("[Book.xlsx]Sheet1").is_none());
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

    #[test]
    fn expand_external_sheet_span_from_order_expands_in_tab_order() {
        let order = vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ];
        let expanded =
            expand_external_sheet_span_from_order("Book.xlsx", "Sheet1", "Sheet3", &order)
                .expect("expand");
        assert_eq!(
            expanded,
            vec![
                "[Book.xlsx]Sheet1".to_string(),
                "[Book.xlsx]Sheet2".to_string(),
                "[Book.xlsx]Sheet3".to_string()
            ]
        );
    }

    #[test]
    fn expand_external_sheet_span_from_order_accepts_reverse_endpoints() {
        let order = vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ];
        let expanded =
            expand_external_sheet_span_from_order("Book.xlsx", "Sheet3", "Sheet1", &order)
                .expect("expand");
        assert_eq!(
            expanded,
            vec![
                "[Book.xlsx]Sheet1".to_string(),
                "[Book.xlsx]Sheet2".to_string(),
                "[Book.xlsx]Sheet3".to_string()
            ]
        );
    }

    #[test]
    fn expand_external_sheet_span_from_order_returns_none_when_endpoints_missing() {
        let order = vec!["Sheet1".to_string(), "Sheet2".to_string()];
        assert!(
            expand_external_sheet_span_from_order("Book.xlsx", "Sheet1", "Sheet3", &order)
                .is_none()
        );
    }
}
