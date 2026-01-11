use crate::error::{ExcelError, ExcelResult};

/// VALUE(text)
///
/// Implements numeric parsing with the common US-style separators (`,` thousands,
/// `.` decimal). Excel's full VALUE function also parses dates/times based on
/// locale; that will be added when the calculation engine has locale context.
pub fn value(text: &str) -> ExcelResult<f64> {
    numbervalue(text, Some('.'), Some(','))
}

/// NUMBERVALUE(number_text, [decimal_separator], [group_separator])
pub fn numbervalue(
    number_text: &str,
    decimal_separator: Option<char>,
    group_separator: Option<char>,
) -> ExcelResult<f64> {
    let decimal_separator = decimal_separator.unwrap_or('.');
    let group_separator = group_separator.unwrap_or(',');

    if decimal_separator == group_separator {
        return Err(ExcelError::Value);
    }

    parse_number_with_separators(number_text, decimal_separator, Some(group_separator))
}

fn parse_number_with_separators(
    text: &str,
    decimal_separator: char,
    group_separator: Option<char>,
) -> ExcelResult<f64> {
    let mut s = text.trim();
    if s.is_empty() {
        return Err(ExcelError::Value);
    }

    let mut negative = false;
    if s.starts_with('(') && s.ends_with(')') {
        negative = true;
        s = s[1..s.len() - 1].trim();
    }

    if let Some(rest) = s.strip_prefix('-') {
        negative = !negative;
        s = rest.trim_start();
    } else if let Some(rest) = s.strip_prefix('+') {
        s = rest.trim_start();
    }

    // Strip a small set of common currency symbols.
    s = s
        .trim_start_matches(|c: char| matches!(c, '$' | '€' | '£' | '¥'))
        .trim();

    let mut percent = false;
    if let Some(rest) = s.strip_suffix('%') {
        percent = true;
        s = rest.trim_end();
    }

    let mut normalized = String::with_capacity(s.len());
    for c in s.chars() {
        if let Some(group) = group_separator {
            if c == group {
                continue;
            }
        }
        if c == decimal_separator {
            normalized.push('.');
        } else if c.is_whitespace() {
            continue;
        } else {
            normalized.push(c);
        }
    }

    let mut number = normalized.parse::<f64>().map_err(|_| ExcelError::Value)?;
    if percent {
        number /= 100.0;
    }
    if negative {
        number = -number;
    }
    if number.is_finite() {
        Ok(number)
    } else {
        Err(ExcelError::Num)
    }
}
