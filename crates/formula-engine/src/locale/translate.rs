use crate::eval::FormulaParseError;

use super::FormulaLocale;

/// Convert a locale-specific formula into the canonical form we persist/evaluate.
///
/// Canonical form uses:
/// - English function names (e.g. `SUM`)
/// - `,` as argument separator
/// - `.` as decimal separator
///
/// The input may include an optional leading `=`, which is preserved in the output.
pub fn canonicalize_formula(formula: &str, locale: &FormulaLocale) -> Result<String, FormulaParseError> {
    translate_formula(formula, locale, Direction::ToCanonical)
}

/// Convert a canonical (English) formula into its locale-specific display form.
///
/// The input may include an optional leading `=`, which is preserved in the output.
pub fn localize_formula(formula: &str, locale: &FormulaLocale) -> Result<String, FormulaParseError> {
    translate_formula(formula, locale, Direction::ToLocalized)
}

#[derive(Debug, Clone, Copy)]
enum Direction {
    ToCanonical,
    ToLocalized,
}

fn translate_formula(
    formula: &str,
    locale: &FormulaLocale,
    dir: Direction,
) -> Result<String, FormulaParseError> {
    let trimmed = formula.trim_start();
    let mut pos = 0usize;
    let mut out = String::with_capacity(trimmed.len());

    if trimmed.as_bytes().first() == Some(&b'=') {
        out.push('=');
        pos += 1;
    }

    while pos < trimmed.len() {
        let Some(ch) = trimmed[pos..].chars().next() else {
            break;
        };

        if ch.is_whitespace() {
            out.push(ch);
            pos += ch.len_utf8();
            continue;
        }

        match ch {
            '"' => {
                pos = copy_quoted_segment(trimmed, pos, '"', &mut out)?;
            }
            '\'' => {
                pos = copy_quoted_segment(trimmed, pos, '\'', &mut out)?;
            }
            _ if is_number_start(trimmed, pos, locale, dir) => {
                let (next, rendered) = scan_number(trimmed, pos, locale, dir)?;
                out.push_str(&rendered);
                pos = next;
            }
            _ if is_ident_start(ch) => {
                let (next, ident) = scan_identifier(trimmed, pos);
                let is_fn = is_function_name(trimmed, next);
                if is_fn {
                    match dir {
                        Direction::ToCanonical => out.push_str(&locale.canonical_function_name(&ident)),
                        Direction::ToLocalized => {
                            out.push_str(&locale.localized_function_name(&ident))
                        }
                    }
                } else {
                    out.push_str(&ident);
                }
                pos = next;
            }
            _ => {
                match dir {
                    Direction::ToCanonical => {
                        if ch == locale.argument_separator {
                            out.push(',');
                        } else {
                            out.push(ch);
                        }
                    }
                    Direction::ToLocalized => {
                        if ch == ',' {
                            out.push(locale.argument_separator);
                        } else {
                            out.push(ch);
                        }
                    }
                }
                pos += ch.len_utf8();
            }
        }
    }

    Ok(out)
}

fn is_ident_start(ch: char) -> bool {
    ch.is_ascii_alphabetic() || ch == '_' || ch == '$'
}

fn scan_identifier(input: &str, start: usize) -> (usize, String) {
    let mut pos = start;
    while let Some(ch) = input[pos..].chars().next() {
        if ch.is_ascii_alphanumeric() || ch == '_' || ch == '.' || ch == '$' {
            pos += ch.len_utf8();
        } else {
            break;
        }
    }
    (pos, input[start..pos].to_string())
}

fn is_function_name(input: &str, mut pos: usize) -> bool {
    while let Some(ch) = input[pos..].chars().next() {
        if ch.is_whitespace() {
            pos += ch.len_utf8();
            continue;
        }
        return ch == '(';
    }
    false
}

fn is_number_start(input: &str, pos: usize, locale: &FormulaLocale, dir: Direction) -> bool {
    let ch = input[pos..].chars().next().unwrap();
    if ch.is_ascii_digit() {
        return true;
    }

    let decimal = match dir {
        Direction::ToCanonical => locale.decimal_separator,
        Direction::ToLocalized => '.',
    };

    if ch != decimal {
        return false;
    }

    let next_pos = pos + ch.len_utf8();
    input
        .get(next_pos..)
        .and_then(|s| s.chars().next())
        .is_some_and(|next| next.is_ascii_digit())
}

fn scan_number(
    input: &str,
    start: usize,
    locale: &FormulaLocale,
    dir: Direction,
) -> Result<(usize, String), FormulaParseError> {
    let decimal_in = match dir {
        Direction::ToCanonical => locale.decimal_separator,
        Direction::ToLocalized => '.',
    };

    let decimal_out = match dir {
        Direction::ToCanonical => '.',
        Direction::ToLocalized => locale.decimal_separator,
    };

    let thousands = match dir {
        Direction::ToCanonical => locale.numeric_thousands_separator(),
        // Canonical formulas do not contain grouping separators.
        Direction::ToLocalized => None,
    };

    let mut pos = start;
    let mut out = String::new();
    let mut seen_decimal = false;

    while let Some(ch) = input.get(pos..).and_then(|s| s.chars().next()) {
        match ch {
            '0'..='9' => {
                out.push(ch);
                pos += 1;
            }
            'e' | 'E' => {
                out.push(ch);
                pos += 1;
                if matches!(input.get(pos..).and_then(|s| s.chars().next()), Some('+') | Some('-')) {
                    let sign = input[pos..].chars().next().unwrap();
                    out.push(sign);
                    pos += 1;
                }
                // Consume exponent digits.
                while matches!(input.get(pos..).and_then(|s| s.chars().next()), Some(d) if d.is_ascii_digit())
                {
                    out.push(input[pos..].chars().next().unwrap());
                    pos += 1;
                }
            }
            _ if Some(ch) == thousands => {
                pos += ch.len_utf8();
            }
            _ if ch == decimal_in && !seen_decimal => {
                out.push(decimal_out);
                seen_decimal = true;
                pos += ch.len_utf8();
            }
            _ => break,
        }
    }

    Ok((pos, out))
}

fn copy_quoted_segment(
    input: &str,
    start: usize,
    quote: char,
    out: &mut String,
) -> Result<usize, FormulaParseError> {
    // Copy opening quote.
    out.push(quote);
    let mut pos = start + quote.len_utf8();

    loop {
        let Some(ch) = input.get(pos..).and_then(|s| s.chars().next()) else {
            return Err(FormulaParseError::UnexpectedEof);
        };

        out.push(ch);
        pos += ch.len_utf8();

        if ch != quote {
            continue;
        }

        // Escaped quote ("" or '') stays within the quoted segment.
        if input.get(pos..).and_then(|s| s.chars().next()) == Some(quote) {
            out.push(quote);
            pos += quote.len_utf8();
            continue;
        }

        break;
    }

    Ok(pos)
}

