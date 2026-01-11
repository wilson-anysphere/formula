use crate::Error;

/// Escape a raw string value for use inside an Excel formula string literal.
///
/// Excel escapes a quote (`"`) inside a string literal by doubling it (`""`).
///
/// This function does *not* add the outer quotes; it only escapes the inner
/// content.
pub fn escape_excel_string_literal(s: &str) -> String {
    if !s.contains('"') {
        return s.to_string();
    }

    // Each quote becomes two quotes, so reserve one extra byte per quote
    // (in UTF-8, `"` is 1 byte).
    let quote_count = s.chars().filter(|&ch| ch == '"').count();
    let mut out = String::with_capacity(s.len() + quote_count);

    for ch in s.chars() {
        if ch == '"' {
            out.push('"');
            out.push('"');
        } else {
            out.push(ch);
        }
    }

    out
}

/// Unescape the content of an Excel formula string literal.
///
/// This function expects the *inner* string literal content (without the outer
/// quotes). Any quote characters must be escaped as doubled quotes (`""`).
pub fn unescape_excel_string_literal(s: &str) -> Result<String, Error> {
    if !s.contains('"') {
        return Ok(s.to_string());
    }

    let mut out = String::with_capacity(s.len());
    let mut chars = s.chars().peekable();

    while let Some(ch) = chars.next() {
        if ch != '"' {
            out.push(ch);
            continue;
        }

        match chars.peek().copied() {
            Some('"') => {
                // `""` => `"`
                chars.next();
                out.push('"');
            }
            _ => {
                return Err(Error::InvalidExcelStringLiteral(
                    "found an unescaped '\"' inside a string literal".to_string(),
                ));
            }
        }
    }

    Ok(out)
}
