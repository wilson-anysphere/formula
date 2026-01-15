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
}
