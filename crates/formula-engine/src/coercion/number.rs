use crate::error::{ExcelError, ExcelResult};

pub(crate) fn parse_number_strict(
    text: &str,
    decimal_sep: char,
    group_sep: Option<char>,
) -> ExcelResult<f64> {
    if text.trim().is_empty() {
        return Err(ExcelError::Value);
    }

    // Reuse the coercion parser for all non-empty inputs so the rules stay consistent between
    // VALUE/NUMBERVALUE and implicit numeric coercion.
    parse_number_coercion(text, decimal_sep, group_sep)
}

pub(crate) fn parse_number_coercion(
    text: &str,
    decimal_sep: char,
    group_sep: Option<char>,
) -> ExcelResult<f64> {
    let s = text.trim();
    if s.is_empty() {
        return Ok(0.0);
    }

    parse_number_nonempty(s, decimal_sep, group_sep)
}

fn parse_number_nonempty(
    mut s: &str,
    decimal_sep: char,
    group_sep: Option<char>,
) -> ExcelResult<f64> {
    if group_sep == Some(decimal_sep) {
        return Err(ExcelError::Value);
    }

    // Accounting formats in Excel use parentheses to signal negation.
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

    // Strip a small set of common currency symbols. Excel is more permissive here, but the core
    // engine only needs the most frequently used characters for now.
    let without_currency = s.trim_start_matches(|c: char| matches!(c, '$' | '€' | '£' | '¥'));
    if without_currency.len() != s.len() {
        s = without_currency.trim_start();
    }

    // Excel accepts multiple trailing percent signs (e.g. "10%%" == 0.001).
    let mut percent_count = 0u32;
    loop {
        let trimmed = s.trim_end();
        if let Some(rest) = trimmed.strip_suffix('%') {
            percent_count += 1;
            s = rest;
            continue;
        }
        s = trimmed;
        break;
    }

    let mut normalized = String::with_capacity(s.len());
    for c in s.chars() {
        if c.is_whitespace() {
            continue;
        }
        if let Some(group) = group_sep {
            if c == group {
                continue;
            }
        }
        if c == decimal_sep {
            normalized.push('.');
        } else {
            normalized.push(c);
        }
    }

    if normalized.is_empty() {
        return Err(ExcelError::Value);
    }

    // Excel does not accept textual "INF"/"NaN" inputs even though Rust's float parser does.
    if normalized.eq_ignore_ascii_case("inf")
        || normalized.eq_ignore_ascii_case("infinity")
        || normalized.eq_ignore_ascii_case("nan")
    {
        return Err(ExcelError::Value);
    }

    let mut number = normalized.parse::<f64>().map_err(|_| ExcelError::Value)?;
    if percent_count != 0 {
        let exponent = i32::try_from(percent_count).unwrap_or(i32::MAX);
        number /= 100_f64.powi(exponent);
    }
    if negative {
        number = -number;
    }

    // Normalize -0 to 0 for Excel parity.
    if number == 0.0 {
        number = 0.0;
    }

    if number.is_finite() {
        Ok(number)
    } else {
        Err(ExcelError::Num)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strict_and_coercion_differ_on_empty_input() {
        assert_eq!(
            parse_number_strict("", '.', Some(',')).unwrap_err(),
            ExcelError::Value
        );
        assert_eq!(parse_number_coercion("", '.', Some(',')).unwrap(), 0.0);
        assert_eq!(parse_number_coercion(" \t ", '.', Some(',')).unwrap(), 0.0);
    }
}
