use crate::error::{ExcelError, ExcelResult};

/// Locale configuration for parsing numbers from text.
///
/// This is intentionally separate from [`crate::LocaleConfig`]: formula lexing has to avoid
/// ambiguous thousands separators (e.g. `,` in en-US formulas), while numeric coercion wants
/// to accept those separators.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NumberLocale {
    pub decimal_separator: char,
    pub group_separator: Option<char>,
}

impl NumberLocale {
    pub const fn en_us() -> Self {
        Self {
            decimal_separator: '.',
            group_separator: Some(','),
        }
    }

    pub const fn new(decimal_separator: char, group_separator: Option<char>) -> Self {
        Self {
            decimal_separator,
            group_separator,
        }
    }
}

pub(crate) fn parse_number(text: &str, locale: NumberLocale) -> ExcelResult<f64> {
    parse_number_with_separators(text, locale.decimal_separator, locale.group_separator)
}

/// Parse a number from text using Excel-like coercion rules.
///
/// This parser is used both by the VALUE/NUMBERVALUE helpers and by implicit coercions
/// (e.g. `"1"+1`). It intentionally targets a pragmatic subset of Excel behavior that
/// matches common spreadsheet inputs:
/// - Leading/trailing whitespace is ignored.
/// - Leading `+` / `-` is supported.
/// - Parentheses negatives: `(123)` -> `-123`.
/// - Common currency symbols (`$ € £ ¥`) are ignored.
/// - Trailing percent: `12%` -> `0.12`.
/// - Group separators are accepted in the integer part when they follow 3-digit grouping.
/// - Scientific notation is supported and validated so group separators don't corrupt exponents.
///
/// Returns:
/// - `Ok(f64)` for finite values (including `0` for empty/whitespace-only inputs).
/// - `Err(ExcelError::Num)` for overflow / infinities.
/// - `Err(ExcelError::Value)` for invalid text.
pub(crate) fn parse_number_with_separators(
    text: &str,
    decimal_separator: char,
    group_separator: Option<char>,
) -> ExcelResult<f64> {
    let mut s = text.trim();
    if s.is_empty() {
        // Excel coerces empty/whitespace-only text to 0 in numeric contexts (`--""`).
        return Ok(0.0);
    }

    let mut negative = false;

    // Parentheses negatives.
    if let Some(inner) = s.strip_prefix('(').and_then(|rest| rest.strip_suffix(')')) {
        negative = true;
        s = inner.trim();
    }

    // Explicit leading sign.
    if let Some(rest) = s.strip_prefix('-') {
        negative = !negative;
        s = rest.trim_start();
    } else if let Some(rest) = s.strip_prefix('+') {
        s = rest.trim_start();
    }

    // Strip a small set of common currency symbols at the start/end.
    s = s
        .trim_start_matches(is_currency_symbol)
        .trim_start()
        .trim_end_matches(is_currency_symbol)
        .trim_end();

    // Percent.
    let mut percent = false;
    if let Some(rest) = s.strip_suffix('%') {
        percent = true;
        s = rest.trim_end();
    }

    if s.is_empty() {
        // If we stripped meaningful symbols and there is nothing left, this isn't a number.
        return Err(ExcelError::Value);
    }

    // Split into mantissa/exponent.
    let (mantissa, exponent) = split_exponent(s)?;

    // Normalize the mantissa with locale separators.
    let normalized_mantissa = normalize_mantissa(mantissa, decimal_separator, group_separator)?;

    let mut normalized = normalized_mantissa;
    if let Some(exponent) = exponent {
        normalized.push('e');
        normalized.push_str(exponent);
    }

    let mut number = normalized.parse::<f64>().map_err(|_| ExcelError::Value)?;
    if percent {
        number /= 100.0;
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

fn split_exponent(s: &str) -> ExcelResult<(&str, Option<&str>)> {
    let mut exp_idx: Option<usize> = None;
    for (idx, ch) in s.char_indices() {
        if ch == 'e' || ch == 'E' {
            if exp_idx.is_some() {
                return Err(ExcelError::Value);
            }
            exp_idx = Some(idx);
        }
    }

    let Some(exp_idx) = exp_idx else {
        return Ok((s, None));
    };

    let mantissa = &s[..exp_idx];
    let exponent = &s[exp_idx + 1..];
    if mantissa.is_empty() || exponent.is_empty() {
        return Err(ExcelError::Value);
    }

    // Validate exponent: `[+-]?\d+`
    let mut chars = exponent.chars();
    let first = chars.next().ok_or(ExcelError::Value)?;
    let mut saw_digit = false;
    match first {
        '+' | '-' => {}
        d if d.is_ascii_digit() => {
            saw_digit = true;
        }
        _ => return Err(ExcelError::Value),
    }
    for ch in chars {
        if ch.is_ascii_digit() {
            saw_digit = true;
        } else {
            return Err(ExcelError::Value);
        }
    }
    if !saw_digit {
        return Err(ExcelError::Value);
    }

    Ok((mantissa, Some(exponent)))
}

fn normalize_mantissa(
    mantissa: &str,
    decimal_separator: char,
    group_separator: Option<char>,
) -> ExcelResult<String> {
    let mut decimal_idx: Option<usize> = None;
    for (idx, ch) in mantissa.char_indices() {
        if ch == decimal_separator {
            if decimal_idx.is_some() {
                return Err(ExcelError::Value);
            }
            decimal_idx = Some(idx);
        }
        if ch == ' ' || ch == '\t' || ch == '\n' || ch == '\r' {
            // Excel trims outer whitespace but does not treat internal whitespace as a digit
            // separator for canonical locales.
            return Err(ExcelError::Value);
        }
    }

    let (int_part, frac_part, has_decimal) = match decimal_idx {
        Some(idx) => {
            let after = &mantissa[idx + decimal_separator.len_utf8()..];
            (&mantissa[..idx], after, true)
        }
        None => (mantissa, "", false),
    };

    if let Some(group) = group_separator {
        if frac_part.contains(group) {
            return Err(ExcelError::Value);
        }
        validate_grouping(int_part, group)?;
    }

    // Validate digits and build the normalized form.
    let mut out = String::with_capacity(mantissa.len());
    let mut saw_digit = false;

    if int_part.is_empty() {
        out.push('0');
    } else {
        for ch in int_part.chars() {
            if Some(ch) == group_separator {
                continue;
            }
            if !ch.is_ascii_digit() {
                return Err(ExcelError::Value);
            }
            saw_digit = true;
            out.push(ch);
        }
    }

    if has_decimal {
        if !frac_part.is_empty() {
            out.push('.');
            for ch in frac_part.chars() {
                if !ch.is_ascii_digit() {
                    return Err(ExcelError::Value);
                }
                saw_digit = true;
                out.push(ch);
            }
        } else if !saw_digit {
            // Input was just "." / "," (depending on locale).
            return Err(ExcelError::Value);
        }
    }

    // If we never saw a digit (e.g. mantissa is empty or only grouping separators), reject.
    if !saw_digit && frac_part.is_empty() {
        return Err(ExcelError::Value);
    }

    Ok(out)
}

fn validate_grouping(int_part: &str, group_separator: char) -> ExcelResult<()> {
    if !int_part.contains(group_separator) {
        // Fast path: no grouping separators. Validate that the remainder is digits only.
        if int_part.is_empty() {
            return Ok(());
        }
        if int_part.chars().all(|c| c.is_ascii_digit()) {
            return Ok(());
        }
        return Err(ExcelError::Value);
    }

    let mut iter = int_part.split(group_separator);
    let first = iter.next().unwrap_or("");
    if first.is_empty() {
        return Err(ExcelError::Value);
    }
    if !first.chars().all(|c| c.is_ascii_digit()) {
        return Err(ExcelError::Value);
    }
    let first_len = first.chars().count();
    if !(1..=3).contains(&first_len) {
        return Err(ExcelError::Value);
    }

    for group in iter {
        if group.is_empty() {
            return Err(ExcelError::Value);
        }
        if group.chars().count() != 3 {
            return Err(ExcelError::Value);
        }
        if !group.chars().all(|c| c.is_ascii_digit()) {
            return Err(ExcelError::Value);
        }
    }

    Ok(())
}

fn is_currency_symbol(ch: char) -> bool {
    matches!(ch, '$' | '€' | '£' | '¥')
}
