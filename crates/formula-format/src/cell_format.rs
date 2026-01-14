use crate::{builtin_format_code, FormatCode, BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX};

/// Return Excel-compatible `CELL("format")` classification code for an Excel number format string.
///
/// This is a **classification** helper: it does not attempt to fully parse/render the format code.
///
/// The return value follows Excel's `CELL("format")` conventions for common numeric formats:
/// - `"G"` for General
/// - `"F<n>"` for fixed/number formats (`n` = decimal places)
/// - `"C<n>"` for currency formats (`n` = decimal places)
/// - `"P<n>"` for percent formats (`n` = decimal places)
/// - `"S<n>"` for scientific formats (`n` = decimal places)
/// - `"D"` for date/time formats (best-effort; Excel uses `D1`..`D9`)
/// - `"@"` for text formats
///
/// Currency detection accounts for:
/// - common currency symbols (`$`, `€`, `£`, `¥`) outside quotes/escapes
/// - OOXML bracket currency tokens like `[$€-407]` (but *not* locale-only tokens like `[$-409]`).
pub fn cell_format_code(format_code: Option<&str>) -> String {
    let code = format_code.unwrap_or("General");
    let code = if code.trim().is_empty() { "General" } else { code };
    let code = resolve_builtin_placeholder(code).unwrap_or(code);

    // Parse into sections so we can correctly choose the "positive" section when conditions are
    // present. When parsing fails, fall back to General classification.
    let parsed = FormatCode::parse(code).unwrap_or_else(|_| FormatCode::general());
    let positive = parsed.select_section_for_number(1.0);
    let pattern = positive.pattern;

    if pattern.trim().eq_ignore_ascii_case("general") {
        return "G".to_string();
    }

    // If the selected positive section looks like a date/time format, Excel returns a `D*` code.
    // We return `"D"` as a best-effort marker for now.
    if crate::datetime::looks_like_datetime(pattern) {
        return "D".to_string();
    }

    if crate::number::pattern_is_text(pattern) {
        return "@".to_string();
    }

    let decimals = count_decimal_places(pattern).min(9);

    let kind = if is_currency_format(pattern) {
        'C'
    } else if is_percent_format(pattern) {
        'P'
    } else if is_scientific_format(pattern) {
        'S'
    } else {
        'F'
    };

    format!("{kind}{decimals}")
}

fn resolve_builtin_placeholder(code: &str) -> Option<&'static str> {
    let id = code
        .strip_prefix(BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX)?
        .trim()
        .parse::<u16>()
        .ok()?;
    builtin_format_code(id)
}

fn count_decimal_places(pattern: &str) -> usize {
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut after_decimal = false;
    let mut count = 0usize;

    for ch in pattern.chars() {
        if escape {
            escape = false;
            continue;
        }

        if in_quotes {
            if ch == '"' {
                in_quotes = false;
            }
            continue;
        }

        if in_brackets {
            if ch == ']' {
                in_brackets = false;
            }
            continue;
        }

        match ch {
            '"' => in_quotes = true,
            '\\' => escape = true,
            '[' => in_brackets = true,
            '.' => {
                after_decimal = true;
                count = 0;
            }
            '0' | '#' | '?' if after_decimal => count += 1,
            _ if after_decimal => break,
            _ => {}
        }
    }

    if after_decimal { count } else { 0 }
}

fn is_percent_format(pattern: &str) -> bool {
    scan_outside_quotes(pattern, |ch| ch == '%')
}

fn is_scientific_format(pattern: &str) -> bool {
    scan_outside_quotes(pattern, |ch| ch == 'E' || ch == 'e')
}

fn is_currency_format(pattern: &str) -> bool {
    // Detect explicit currency symbols outside quotes/escapes, OR bracket currency tokens like
    // `[$€-407]`. Locale-only tokens like `[$-409]` should *not* be treated as currency.
    scan_outside_quotes(pattern, |ch| matches!(ch, '$' | '€' | '£' | '¥'))
        || contains_bracket_currency_token(pattern)
}

fn contains_bracket_currency_token(pattern: &str) -> bool {
    let mut in_quotes = false;
    let mut escape = false;
    let mut chars = pattern.chars().peekable();

    while let Some(ch) = chars.next() {
        if escape {
            escape = false;
            continue;
        }

        if in_quotes {
            if ch == '"' {
                in_quotes = false;
            }
            continue;
        }

        match ch {
            '"' => in_quotes = true,
            '\\' => escape = true,
            '[' => {
                let mut content = String::new();
                let mut closed = false;
                while let Some(c) = chars.next() {
                    if c == ']' {
                        closed = true;
                        break;
                    }
                    content.push(c);
                }
                if !closed {
                    // No closing bracket: treat as literal and stop probing this token.
                    continue;
                }
                if bracket_is_currency(&content) {
                    return true;
                }
            }
            _ => {}
        }
    }

    false
}

fn bracket_is_currency(content: &str) -> bool {
    let content = content.trim();
    let Some(after) = content.strip_prefix('$') else {
        return false;
    };
    // We only treat `[$<sym>-<lcid>]` as currency when `<sym>` is non-empty.
    let Some((symbol, _lcid)) = after.split_once('-') else {
        return false;
    };
    !symbol.is_empty()
}

fn scan_outside_quotes(pattern: &str, pred: impl Fn(char) -> bool) -> bool {
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;

    for ch in pattern.chars() {
        if escape {
            escape = false;
            continue;
        }

        if in_quotes {
            if ch == '"' {
                in_quotes = false;
            }
            continue;
        }

        if in_brackets {
            if ch == ']' {
                in_brackets = false;
            }
            continue;
        }

        match ch {
            '"' => in_quotes = true,
            '\\' => escape = true,
            '[' => in_brackets = true,
            _ if pred(ch) => return true,
            _ => {}
        }
    }

    false
}
