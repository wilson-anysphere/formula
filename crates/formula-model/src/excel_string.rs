/// Helpers for Excel formula string literal escaping.
///
/// Excel uses double quotes (`"`) to delimit string literals in formulas, and escapes a literal
/// quote by doubling it (`""`).
#[inline]
pub fn push_escaped_excel_double_quote_char(out: &mut String, ch: char) {
    if ch == '"' {
        out.push('"');
        out.push('"');
    } else {
        out.push(ch);
    }
}

/// Push `value` into `out`, escaping embedded quotes as `""`.
pub fn push_escaped_excel_double_quotes(out: &mut String, value: &str) {
    for ch in value.chars() {
        push_escaped_excel_double_quote_char(out, ch);
    }
}

/// Push a complete Excel formula string literal (`"..."`) into `out`.
pub fn push_excel_double_quoted_string_literal(out: &mut String, value: &str) {
    out.push('"');
    push_escaped_excel_double_quotes(out, value);
    out.push('"');
}

/// Escape embedded quotes in `value` as `""`.
#[must_use]
pub fn escape_excel_double_quotes(value: &str) -> String {
    if !value.contains('"') {
        return value.to_string();
    }

    // Each quote becomes two quotes, so reserve one extra byte per quote.
    let quote_count = value.chars().filter(|&ch| ch == '"').count();
    let mut out = String::with_capacity(value.len().saturating_add(quote_count));
    push_escaped_excel_double_quotes(&mut out, value);
    out
}

/// Unescape `inner` content of an Excel formula string literal.
///
/// `inner` must not include the outer quotes. Any quote characters must be escaped as doubled
/// quotes (`""`).
pub fn unescape_excel_double_quotes(inner: &str) -> Option<String> {
    if !inner.contains('"') {
        return Some(inner.to_string());
    }

    let mut out = String::with_capacity(inner.len());
    let mut chars = inner.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch != '"' {
            out.push(ch);
            continue;
        }

        match chars.peek().copied() {
            Some('"') => {
                chars.next();
                out.push('"');
            }
            _ => return None,
        }
    }
    Some(out)
}

/// Unescape a full Excel formula string literal (`"..."`).
pub fn unescape_excel_double_quoted_string_literal(raw: &str) -> Option<String> {
    let inner = raw.strip_prefix('"')?.strip_suffix('"')?;
    unescape_excel_double_quotes(inner)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn escape_excel_double_quotes_doubles_quotes() {
        assert_eq!(escape_excel_double_quotes(r#"a"b"#), r#"a""b"#);
        assert_eq!(escape_excel_double_quotes("plain"), "plain");
    }

    #[test]
    fn push_excel_double_quoted_string_literal_wraps_and_escapes() {
        let mut out = String::new();
        push_excel_double_quoted_string_literal(&mut out, r#"a"b"#);
        assert_eq!(out, r#""a""b""#);
    }

    #[test]
    fn push_escaped_excel_double_quote_char_doubles_quote_char_only() {
        let mut out = String::new();
        push_escaped_excel_double_quote_char(&mut out, '"');
        push_escaped_excel_double_quote_char(&mut out, 'x');
        assert_eq!(out, "\"\"x");
    }

    #[test]
    fn unescape_excel_double_quotes_rejects_unescaped_quote() {
        assert_eq!(unescape_excel_double_quotes(r#"a"b"#), None);
    }

    #[test]
    fn unescape_excel_double_quotes_accepts_doubled_quotes() {
        assert_eq!(unescape_excel_double_quotes(r#"a""b"#), Some(r#"a"b"#.to_string()));
        assert_eq!(unescape_excel_double_quotes("plain"), Some("plain".to_string()));
        assert_eq!(unescape_excel_double_quotes(""), Some("".to_string()));
    }

    #[test]
    fn unescape_excel_double_quoted_string_literal_strips_outer_quotes() {
        assert_eq!(unescape_excel_double_quoted_string_literal(r#""a""b""#), Some(r#"a"b"#.to_string()));
        assert_eq!(unescape_excel_double_quoted_string_literal(r#""""#), Some("".to_string()));
        assert_eq!(unescape_excel_double_quoted_string_literal(r#""a"b""#), None);
        assert_eq!(unescape_excel_double_quoted_string_literal("nope"), None);
    }
}

