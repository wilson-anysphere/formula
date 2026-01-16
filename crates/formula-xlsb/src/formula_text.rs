use crate::Error;

/// Escape a raw string value for use inside an Excel formula string literal.
///
/// Excel escapes a quote (`"`) inside a string literal by doubling it (`""`).
///
/// This function does *not* add the outer quotes; it only escapes the inner
/// content.
pub fn escape_excel_string_literal(s: &str) -> String {
    formula_model::escape_excel_double_quotes(s)
}

/// Unescape the content of an Excel formula string literal.
///
/// This function expects the *inner* string literal content (without the outer
/// quotes). Any quote characters must be escaped as doubled quotes (`""`).
pub fn unescape_excel_string_literal(s: &str) -> Result<String, Error> {
    if !s.contains('"') {
        return Ok(s.to_string());
    }
    formula_model::unescape_excel_double_quotes(s).ok_or_else(|| {
        Error::InvalidExcelStringLiteral(
            "found an unescaped '\"' inside a string literal".to_string(),
        )
    })
}
