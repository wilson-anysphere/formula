use core::fmt;

use std::borrow::Cow;

use unicode_normalization::UnicodeNormalization;

/// Maximum worksheet name length enforced by Excel.
///
/// Excel stores sheet names as UTF-16 and enforces the 31-character limit in terms of UTF-16 code
/// units. That means characters outside the BMP (e.g. many emoji) count as 2.
pub const EXCEL_MAX_SHEET_NAME_LEN: usize = 31;

/// Errors returned when validating worksheet names.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SheetNameError {
    /// The name is empty or consists only of whitespace.
    EmptyName,
    /// The name exceeds [`EXCEL_MAX_SHEET_NAME_LEN`].
    TooLong,
    /// The name contains a character that Excel forbids in worksheet names.
    InvalidCharacter(char),
    /// Excel forbids worksheet names that begin or end with `'`.
    LeadingOrTrailingApostrophe,
    /// The name conflicts with an existing sheet (case-insensitive).
    DuplicateName,
}

impl fmt::Display for SheetNameError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            SheetNameError::EmptyName => f.write_str("sheet name cannot be blank"),
            SheetNameError::TooLong => write!(
                f,
                "sheet name cannot exceed {EXCEL_MAX_SHEET_NAME_LEN} characters"
            ),
            SheetNameError::InvalidCharacter(ch) => {
                write!(f, "sheet name contains invalid character `{ch}`")
            }
            SheetNameError::LeadingOrTrailingApostrophe => {
                f.write_str("sheet name cannot begin or end with an apostrophe")
            }
            SheetNameError::DuplicateName => f.write_str("sheet name already exists"),
        }
    }
}

impl std::error::Error for SheetNameError {}

/// Validate a worksheet name using Excel-compatible rules.
pub fn validate_sheet_name(name: &str) -> Result<(), SheetNameError> {
    if name.trim().is_empty() {
        return Err(SheetNameError::EmptyName);
    }

    if name.encode_utf16().count() > EXCEL_MAX_SHEET_NAME_LEN {
        return Err(SheetNameError::TooLong);
    }

    if let Some(ch) = name
        .chars()
        .find(|ch| matches!(ch, ':' | '\\' | '/' | '?' | '*' | '[' | ']'))
    {
        return Err(SheetNameError::InvalidCharacter(ch));
    }

    if name.starts_with('\'') || name.ends_with('\'') {
        return Err(SheetNameError::LeadingOrTrailingApostrophe);
    }

    Ok(())
}

/// Excel compares sheet names case-insensitively across Unicode.
///
/// We approximate Excel's behavior by normalizing both names with Unicode NFKC (compatibility
/// normalization) and then applying Unicode uppercasing. This is deterministic and locale-independent.
pub fn sheet_name_eq_case_insensitive(a: &str, b: &str) -> bool {
    if a == b {
        return true;
    }
    if a.is_ascii() && b.is_ascii() {
        return a.eq_ignore_ascii_case(b);
    }
    a.nfkc()
        .flat_map(|c| c.to_uppercase())
        .eq(b.nfkc().flat_map(|c| c.to_uppercase()))
}

/// Returns a canonical "case folded" representation of a sheet name that matches
/// [`sheet_name_eq_case_insensitive`].
///
/// This is useful when building hash map keys for sheet-name lookups that need to behave like Excel
/// (e.g. treating `Straße` and `STRASSE` as equal).
pub fn sheet_name_casefold(name: &str) -> String {
    if name.is_ascii() {
        return name.to_ascii_uppercase();
    }
    name.nfkc().flat_map(|c| c.to_uppercase()).collect()
}

fn truncate_to_utf16_units(name: &str, max_units: usize) -> String {
    if name.encode_utf16().count() <= max_units {
        return name.to_string();
    }

    let mut out = String::new();
    let mut units = 0usize;
    for ch in name.chars() {
        let ch_units = ch.len_utf16();
        if units + ch_units > max_units {
            break;
        }
        out.push(ch);
        units += ch_units;
    }
    out
}

/// Sanitize a sheet name so that it always satisfies [`validate_sheet_name`].
///
/// This helper is intended for inputs like file names (CSV/Parquet imports) that might contain
/// characters Excel disallows in worksheet names.
///
/// Notes:
/// - This does **not** guarantee uniqueness within a workbook (e.g. two different inputs could both
///   sanitize to `"Sheet1"`).
pub fn sanitize_sheet_name(name: &str) -> String {
    let mut sanitized = name.trim().to_string();

    sanitized.retain(|ch| !matches!(ch, ':' | '\\' | '/' | '?' | '*' | '[' | ']'));

    sanitized = sanitized.trim().trim_matches('\'').trim().to_string();

    sanitized = truncate_to_utf16_units(&sanitized, EXCEL_MAX_SHEET_NAME_LEN);

    // Truncation can create a leading/trailing apostrophe (e.g. if an internal apostrophe becomes
    // the final character), so strip again.
    sanitized = sanitized.trim().trim_matches('\'').trim().to_string();

    if sanitized.is_empty() {
        return "Sheet1".to_string();
    }

    debug_assert!(
        validate_sheet_name(&sanitized).is_ok(),
        "sanitize_sheet_name produced invalid name: {sanitized:?}"
    );

    sanitized
}

/// Escape `'` for Excel single-quoted identifiers by doubling it (`'` → `''`).
///
/// Excel uses single quotes to quote sheet references and other identifier-like tokens. Within a
/// quoted identifier, a literal `'` is represented as `''`.
pub fn escape_excel_single_quotes(s: &str) -> Cow<'_, str> {
    if !s.contains('\'') {
        return Cow::Borrowed(s);
    }
    let mut out = String::with_capacity(s.len().saturating_add(4));
    push_escaped_excel_single_quotes(&mut out, s);
    Cow::Owned(out)
}

/// Unescape Excel single-quoted identifier content **leniently**.
///
/// Excel represents a literal `'` within a single-quoted identifier by doubling it (`''`).
///
/// This helper performs the inverse mapping (`''` -> `'`) while leaving any unmatched `'`
/// characters as-is. This matches the historical behavior of `s.replace(\"''\", \"'\")` used in
/// various import paths.
pub fn unescape_excel_single_quotes_lenient(inner: &str) -> Cow<'_, str> {
    if !inner.contains("''") {
        return Cow::Borrowed(inner);
    }

    let mut out = String::with_capacity(inner.len());
    let mut chars = inner.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '\'' && chars.peek() == Some(&'\'') {
            chars.next();
            out.push('\'');
        } else {
            out.push(ch);
        }
    }
    Cow::Owned(out)
}

/// Unquote an Excel single-quoted identifier (`'...'`) **leniently**.
///
/// Returns `None` if `raw` is not wrapped in single quotes; otherwise returns the unescaped inner
/// content as described by [`unescape_excel_single_quotes_lenient`].
pub fn unquote_excel_single_quoted_identifier_lenient(raw: &str) -> Option<Cow<'_, str>> {
    let inner = raw.strip_prefix('\'')?.strip_suffix('\'')?;
    Some(unescape_excel_single_quotes_lenient(inner))
}

/// Unquote a worksheet name **leniently**.
///
/// This mirrors common importer behavior:
/// - trims whitespace
/// - if wrapped in single quotes, unescapes doubled quotes (`''` -> `'`)
/// - otherwise returns the trimmed name as-is
///
/// Note: this always returns an owned `String` because many callers need an owned sheet name.
pub fn unquote_sheet_name_lenient(name: &str) -> String {
    let trimmed = name.trim();
    unquote_excel_single_quoted_identifier_lenient(trimmed)
        .map(|inner| inner.into_owned())
        .unwrap_or_else(|| trimmed.to_string())
}

/// Append `s` to `out`, escaping `'` as `''` (Excel single-quoted identifier rules).
pub fn push_escaped_excel_single_quotes(out: &mut String, s: &str) {
    for ch in s.chars() {
        if ch == '\'' {
            out.push('\'');
            out.push('\'');
        } else {
            out.push(ch);
        }
    }
}

/// Append a single-quoted Excel identifier to `out`, escaping internal `'` as `''`.
pub fn push_excel_single_quoted_identifier(out: &mut String, s: &str) {
    out.push('\'');
    push_escaped_excel_single_quotes(out, s);
    out.push('\'');
}

/// Append a worksheet name suitable for embedding in an A1-style sheet reference.
///
/// This emits either `Sheet1` or a single-quoted/escaped form like `'Sheet 1'` depending on
/// [`sheet_name_needs_quotes_a1`].
pub fn push_sheet_name_a1(out: &mut String, sheet: &str) {
    if sheet_name_needs_quotes_a1(sheet) {
        push_excel_single_quoted_identifier(out, sheet);
    } else {
        out.push_str(sheet);
    }
}

/// Append a worksheet name suitable for embedding in an R1C1-style sheet reference.
///
/// This emits either `Sheet1` or a single-quoted/escaped form like `'Sheet 1'` depending on
/// [`sheet_name_needs_quotes_r1c1`].
pub fn push_sheet_name_r1c1(out: &mut String, sheet: &str) {
    if sheet_name_needs_quotes_r1c1(sheet) {
        push_excel_single_quoted_identifier(out, sheet);
    } else {
        out.push_str(sheet);
    }
}

/// Append a 3D sheet range name (`Sheet1:Sheet3`), quoting/escaping as needed for A1-style formulas.
///
/// Excel represents 3D sheet ranges as:
/// - `Sheet1:Sheet3!A1` for simple sheet identifiers
/// - `'Sheet 1:Sheet3'!A1` when either side requires quoting.
///
/// Note: quoting each side independently (`'Sheet 1':Sheet3!A1`) is not a valid 3D sheet range.
pub fn push_sheet_range_name_a1(out: &mut String, start: &str, end: &str) {
    if !sheet_name_needs_quotes_a1(start) && !sheet_name_needs_quotes_a1(end) {
        out.push_str(start);
        out.push(':');
        out.push_str(end);
        return;
    }

    out.push('\'');
    push_escaped_excel_single_quotes(out, start);
    out.push(':');
    push_escaped_excel_single_quotes(out, end);
    out.push('\'');
}

/// Append a 3D sheet range name (`Sheet1:Sheet3`), quoting/escaping as needed for R1C1-style formulas.
///
/// See [`push_sheet_range_name_a1`] for the quoting rules; this variant uses
/// [`sheet_name_needs_quotes_r1c1`].
pub fn push_sheet_range_name_r1c1(out: &mut String, start: &str, end: &str) {
    if !sheet_name_needs_quotes_r1c1(start) && !sheet_name_needs_quotes_r1c1(end) {
        out.push_str(start);
        out.push(':');
        out.push_str(end);
        return;
    }

    out.push('\'');
    push_escaped_excel_single_quotes(out, start);
    out.push(':');
    push_escaped_excel_single_quotes(out, end);
    out.push('\'');
}

/// Returns `true` if `sheet` must be single-quoted to be safely embedded as a sheet reference in
/// an A1-style formula (e.g. `Sheet1!A1`).
///
/// This is intentionally conservative: quoting is always accepted by Excel, while the unquoted
/// form is only valid for a subset of identifier-like sheet names.
pub fn sheet_name_needs_quotes_a1(sheet: &str) -> bool {
    if sheet.is_empty() {
        return true;
    }
    if sheet
        .chars()
        .any(|c| c.is_whitespace() || matches!(c, '!' | '\''))
    {
        return true;
    }

    if sheet.eq_ignore_ascii_case("TRUE") || sheet.eq_ignore_ascii_case("FALSE") {
        return true;
    }

    // Sheet names that would be tokenized as an A1 cell reference prefix must be quoted to avoid
    // ambiguous parses (e.g. `'A1'!B2`, `'A1.Price'!B2`).
    if starts_like_a1_cell_ref(sheet) {
        return true;
    }

    // Excel also requires quoting sheet names that look like R1C1 references, even when the
    // surrounding formula uses A1-style references (e.g. `'R1C1'!A1`).
    if starts_like_r1c1_ref(sheet) {
        return true;
    }

    !is_valid_sheet_ident(sheet)
}

/// Returns `true` if `sheet` must be single-quoted to be safely embedded as a sheet reference in
/// an R1C1-style formula (e.g. `Sheet1!R1C1`).
pub fn sheet_name_needs_quotes_r1c1(sheet: &str) -> bool {
    if sheet_name_needs_quotes_a1(sheet) {
        return true;
    }

    // In R1C1 mode, sheet names like `R1C1` / `RC` / `R1` can be tokenized as references.
    starts_like_r1c1_ref(sheet)
}

/// Returns `true` if `s` starts with an A1-style cell reference (e.g. `A1`, `$B$2`).
///
/// This is intended for lexer disambiguation (sheet/name identifiers vs cell refs) rather than
/// strict Excel bounds checking.
pub fn starts_like_a1_cell_ref(s: &str) -> bool {
    let mut chars = s.chars().peekable();
    if chars.peek() == Some(&'$') {
        chars.next();
    }

    let mut col_len = 0usize;
    while let Some(&ch) = chars.peek() {
        if ch.is_ascii_alphabetic() {
            col_len += 1;
            if col_len > 3 {
                return false;
            }
            chars.next();
        } else {
            break;
        }
    }
    if col_len == 0 {
        return false;
    }

    if chars.peek() == Some(&'$') {
        chars.next();
    }

    let mut row: u32 = 0;
    let mut row_len = 0usize;
    let mut overflow = false;
    while let Some(&ch) = chars.peek() {
        if ch.is_ascii_digit() {
            row_len += 1;
            row = match row
                .checked_mul(10)
                .and_then(|v| v.checked_add((ch as u8 - b'0') as u32))
            {
                Some(v) => v,
                None => {
                    overflow = true;
                    0
                }
            };
            chars.next();
        } else {
            break;
        }
    }

    if row_len == 0 || overflow || row == 0 {
        return false;
    }

    // Match the lexer guard: treat as a cell ref only if the next character does *not* continue
    // an identifier (except `.` which is used for field access like `A1.Price`).
    !matches!(chars.peek(), Some(c) if (is_sheet_ident_cont_char(*c) && *c != '.') || *c == '(')
}

/// Returns `true` if `s` starts with an R1C1-style reference token (`RC`, `R1C2`, `R1`, `C2`).
pub fn starts_like_r1c1_ref(s: &str) -> bool {
    starts_like_r1c1_cell_ref(s) || starts_like_r1c1_row_ref(s) || starts_like_r1c1_col_ref(s)
}

fn starts_like_r1c1_cell_ref(s: &str) -> bool {
    let mut chars = s.chars().peekable();
    let Some(first) = chars.next() else {
        return false;
    };
    if !matches!(first, 'R' | 'r') {
        return false;
    }

    if matches!(chars.peek(), Some(c) if c.is_ascii_digit()) && !consume_u32_nonzero(&mut chars) {
        return false;
    }

    let Some(ch) = chars.next() else {
        return false;
    };
    if !matches!(ch, 'C' | 'c') {
        return false;
    }

    if matches!(chars.peek(), Some(c) if c.is_ascii_digit()) && !consume_u32_nonzero(&mut chars) {
        return false;
    }

    true
}

fn starts_like_r1c1_row_ref(s: &str) -> bool {
    let mut chars = s.chars().peekable();
    let Some(first) = chars.next() else {
        return false;
    };
    if !matches!(first, 'R' | 'r') {
        return false;
    }

    if matches!(chars.peek(), Some(c) if c.is_ascii_digit()) && !consume_u32_nonzero(&mut chars) {
        return false;
    }

    // Matches the lexer guard: treat as a row ref only if the next character does *not* continue
    // an identifier.
    !matches!(chars.peek(), Some(c) if is_sheet_ident_cont_char(*c) || *c == '(')
}

fn starts_like_r1c1_col_ref(s: &str) -> bool {
    let mut chars = s.chars().peekable();
    let Some(first) = chars.next() else {
        return false;
    };
    if !matches!(first, 'C' | 'c') {
        return false;
    }

    if matches!(chars.peek(), Some(c) if c.is_ascii_digit()) && !consume_u32_nonzero(&mut chars) {
        return false;
    }

    !matches!(chars.peek(), Some(c) if is_sheet_ident_cont_char(*c) || *c == '(')
}

fn consume_u32_nonzero<I>(chars: &mut std::iter::Peekable<I>) -> bool
where
    I: Iterator<Item = char>,
{
    let mut value: u32 = 0;
    let mut len = 0usize;
    while matches!(chars.peek(), Some(c) if c.is_ascii_digit()) {
        let ch = chars.next().expect("peeked");
        len += 1;
        value = match value
            .checked_mul(10)
            .and_then(|v| v.checked_add((ch as u8 - b'0') as u32))
        {
            Some(v) => v,
            None => return false,
        };
    }
    len > 0 && value != 0
}

fn is_valid_sheet_ident(ident: &str) -> bool {
    let mut chars = ident.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !matches!(first, '$' | '_' | '\\' | 'A'..='Z' | 'a'..='z') {
        return false;
    }
    chars.all(is_sheet_ident_cont_char)
}

fn is_sheet_ident_cont_char(c: char) -> bool {
    matches!(
        c,
        '$' | '_' | '\\' | '.' | 'A'..='Z' | 'a'..='z' | '0'..='9'
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn escape_excel_single_quotes_doubles_apostrophes() {
        assert_eq!(escape_excel_single_quotes("Sheet1").as_ref(), "Sheet1");
        assert_eq!(escape_excel_single_quotes("O'Brien").as_ref(), "O''Brien");
        assert_eq!(
            escape_excel_single_quotes("a'b'c").as_ref(),
            "a''b''c"
        );
    }

    #[test]
    fn push_excel_single_quoted_identifier_wraps_and_escapes() {
        let mut out = String::new();
        push_excel_single_quoted_identifier(&mut out, "O'Brien");
        assert_eq!(out, "'O''Brien'");
    }

    #[test]
    fn push_sheet_name_a1_quotes_only_when_needed() {
        let mut out = String::new();
        push_sheet_name_a1(&mut out, "Sheet1");
        assert_eq!(out, "Sheet1");

        let mut out = String::new();
        push_sheet_name_a1(&mut out, "Sheet 1");
        assert_eq!(out, "'Sheet 1'");

        let mut out = String::new();
        push_sheet_name_a1(&mut out, "O'Brien");
        assert_eq!(out, "'O''Brien'");
    }

    #[test]
    fn push_sheet_name_r1c1_quotes_only_when_needed() {
        let mut out = String::new();
        push_sheet_name_r1c1(&mut out, "Sheet1");
        assert_eq!(out, "Sheet1");

        let mut out = String::new();
        push_sheet_name_r1c1(&mut out, "Sheet 1");
        assert_eq!(out, "'Sheet 1'");

        let mut out = String::new();
        push_sheet_name_r1c1(&mut out, "O'Brien");
        assert_eq!(out, "'O''Brien'");
    }

    #[test]
    fn push_sheet_range_name_a1_quotes_when_needed() {
        let mut out = String::new();
        push_sheet_range_name_a1(&mut out, "Sheet1", "Sheet3");
        assert_eq!(out, "Sheet1:Sheet3");

        let mut out = String::new();
        push_sheet_range_name_a1(&mut out, "Sheet 1", "Sheet3");
        assert_eq!(out, "'Sheet 1:Sheet3'");

        let mut out = String::new();
        push_sheet_range_name_a1(&mut out, "O'Brien", "Sheet3");
        assert_eq!(out, "'O''Brien:Sheet3'");
    }

    #[test]
    fn push_sheet_range_name_r1c1_quotes_when_needed() {
        let mut out = String::new();
        push_sheet_range_name_r1c1(&mut out, "Sheet1", "Sheet3");
        assert_eq!(out, "Sheet1:Sheet3");

        let mut out = String::new();
        push_sheet_range_name_r1c1(&mut out, "Sheet1", "RC");
        assert_eq!(out, "'Sheet1:RC'");
    }

    #[test]
    fn unescape_excel_single_quotes_lenient_matches_replace_semantics() {
        assert_eq!(
            unescape_excel_single_quotes_lenient("O''Brien").as_ref(),
            "O'Brien"
        );
        // `replace("''", "'")` leaves unmatched quotes as-is.
        assert_eq!(
            unescape_excel_single_quotes_lenient("a'''b").as_ref(),
            "a''b"
        );
        assert_eq!(
            unescape_excel_single_quotes_lenient("a'b").as_ref(),
            "a'b"
        );
    }

    #[test]
    fn unquote_excel_single_quoted_identifier_lenient_strips_and_unescapes() {
        assert_eq!(
            unquote_excel_single_quoted_identifier_lenient("'O''Brien'")
                .expect("quoted")
                .as_ref(),
            "O'Brien"
        );
        assert!(unquote_excel_single_quoted_identifier_lenient("Sheet1").is_none());
    }

    #[test]
    fn sheet_name_needs_quotes_a1_matches_lexer_disambiguation_rules() {
        assert!(sheet_name_needs_quotes_a1(""));
        assert!(sheet_name_needs_quotes_a1("Sheet 1"));
        assert!(sheet_name_needs_quotes_a1("O'Brien"));
        assert!(sheet_name_needs_quotes_a1("Sheet!"));

        assert!(sheet_name_needs_quotes_a1("TRUE"));
        assert!(sheet_name_needs_quotes_a1("FALSE"));
        assert!(sheet_name_needs_quotes_a1("A1"));
        assert!(sheet_name_needs_quotes_a1("$B$2"));
        assert!(sheet_name_needs_quotes_a1("A1.Price"));
        assert!(!sheet_name_needs_quotes_a1("A1B"));

        assert!(!sheet_name_needs_quotes_a1("Sheet1"));
        assert!(!sheet_name_needs_quotes_a1("_Sheet1"));
        assert!(sheet_name_needs_quotes_a1("R1C1"));
    }

    #[test]
    fn sheet_name_needs_quotes_r1c1_quotes_r1c1_like_prefixes() {
        assert!(sheet_name_needs_quotes_r1c1("R1C1"));
        assert!(sheet_name_needs_quotes_r1c1("RC"));
        assert!(sheet_name_needs_quotes_r1c1("R1"));
        assert!(sheet_name_needs_quotes_r1c1("C2"));

        assert!(!sheet_name_needs_quotes_r1c1("Sheet1"));
    }
}
