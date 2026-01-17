use num_complex::Complex64;
use std::borrow::Cow;

use crate::functions::FunctionContext;
use crate::value::{ErrorKind, NumberLocale, Value};
use crate::LocaleConfig;

#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) struct ParsedComplex {
    pub value: Complex64,
    pub suffix: char,
}

fn parse_component(text: &str, locale: NumberLocale) -> Result<f64, ErrorKind> {
    // Excel's complex-number functions surface invalid parsing as #NUM!, even when the
    // underlying numeric coercion would normally yield #VALUE!.
    //
    // Use `LocaleConfig::parse_number` rather than `value::parse_number` so complex strings can
    // accept both:
    // - the workbook decimal separator (e.g. `1,5` in `de-DE`)
    // - the canonical `.` decimal separator (e.g. `1.5`), mirroring criteria parsing behavior.
    let mut cfg = LocaleConfig::en_us();
    cfg.decimal_separator = locale.decimal_separator;
    cfg.thousands_separator = locale.group_separator;
    let n = cfg.parse_number(text).ok_or(ErrorKind::Num)?;
    if n.is_finite() {
        Ok(n)
    } else {
        Err(ErrorKind::Num)
    }
}

fn strip_whitespace_if_needed(s: &str) -> Cow<'_, str> {
    let mut it = s.char_indices();
    while let Some((idx, ch)) = it.next() {
        if !ch.is_whitespace() {
            continue;
        }

        let mut out = String::with_capacity(s.len());
        out.push_str(&s[..idx]);
        for (_, ch) in it {
            if !ch.is_whitespace() {
                out.push(ch);
            }
        }
        return Cow::Owned(out);
    }
    Cow::Borrowed(s)
}

/// Parse an Excel-style complex number string (engineering functions).
///
/// Supported forms:
/// - `3+4i`, `3-4i`
/// - `4i`, `-4j`
/// - `i`, `-i`, `+i`
/// - `3` (pure real)
///
/// Whitespace is ignored.
pub(crate) fn parse_complex(text: &str, locale: NumberLocale) -> Result<ParsedComplex, ErrorKind> {
    let normalized = strip_whitespace_if_needed(text);
    let normalized = normalized.as_ref();
    if normalized.is_empty() {
        return Ok(ParsedComplex {
            value: Complex64::new(0.0, 0.0),
            suffix: 'i',
        });
    }

    let (suffix, core) = match normalized.chars().last() {
        Some('i') | Some('I') => Some('i'),
        Some('j') | Some('J') => Some('j'),
        _ => None,
    }
    .map(|suffix| {
        let last = normalized
            .chars()
            .last()
            .expect("non-empty normalized has a last char");
        let core_end = normalized.len().saturating_sub(last.len_utf8());
        (suffix, &normalized[..core_end])
    })
    .unwrap_or(('i', normalized));

    if normalized.len() != core.len() {
        // Excel only allows the imaginary unit suffix in the final position.
        if core.chars().any(|c| matches!(c, 'i' | 'I' | 'j' | 'J')) {
            return Err(ErrorKind::Num);
        }

        if core.is_empty() || core == "+" {
            return Ok(ParsedComplex {
                value: Complex64::new(0.0, 1.0),
                suffix,
            });
        }
        if core == "-" {
            return Ok(ParsedComplex {
                value: Complex64::new(0.0, -1.0),
                suffix,
            });
        }

        // "a+i"/"a-i" (implicit coefficient of 1).
        if let Some(last) = core.chars().last() {
            if last == '+' || last == '-' {
                let re_str = &core[..core.len() - last.len_utf8()];
                let re = parse_component(re_str, locale)?;
                let im = if last == '+' { 1.0 } else { -1.0 };
                return Ok(ParsedComplex {
                    value: Complex64::new(re, im),
                    suffix,
                });
            }
        }

        // Pure imaginary coefficient ("4i", "-1.5i").
        if let Ok(im) = parse_component(core, locale) {
            return Ok(ParsedComplex {
                value: Complex64::new(0.0, im),
                suffix,
            });
        }

        // Split into `real` and `imag` parts using the final '+'/'-' that is not part of an exponent.
        let mut split_idx: Option<usize> = None;
        for (idx, ch) in core.char_indices().rev() {
            if idx == 0 {
                continue;
            }
            if ch != '+' && ch != '-' {
                continue;
            }
            let prev = core[..idx].chars().last().unwrap_or('\0');
            if prev == 'e' || prev == 'E' {
                continue;
            }
            split_idx = Some(idx);
            break;
        }

        let split_idx = split_idx.ok_or(ErrorKind::Num)?;
        let (re_str, im_str) = core.split_at(split_idx);
        let re = parse_component(re_str, locale)?;
        let im = parse_component(im_str, locale)?;
        Ok(ParsedComplex {
            value: Complex64::new(re, im),
            suffix,
        })
    } else {
        // Reject stray i/j characters elsewhere.
        if normalized.chars().any(|c| matches!(c, 'i' | 'I' | 'j' | 'J')) {
            return Err(ErrorKind::Num);
        }

        let re = parse_component(normalized, locale)?;
        Ok(ParsedComplex {
            value: Complex64::new(re, 0.0),
            suffix: 'i',
        })
    }
}

pub(crate) fn format_complex(
    mut z: Complex64,
    suffix: char,
    ctx: &dyn FunctionContext,
) -> Result<String, ErrorKind> {
    if !z.re.is_finite() || !z.im.is_finite() {
        return Err(ErrorKind::Num);
    }

    // Normalize -0 to 0 for Excel parity.
    if z.re == 0.0 {
        z.re = 0.0;
    }
    if z.im == 0.0 {
        z.im = 0.0;
    }

    if z.im == 0.0 {
        return Value::Number(z.re).coerce_to_string_with_ctx(ctx);
    }

    let suffix = match suffix {
        'j' | 'J' => 'j',
        _ => 'i',
    };

    if z.re == 0.0 {
        return Ok(format_imag_only(z.im, suffix, ctx)?);
    }

    let re_str = Value::Number(z.re).coerce_to_string_with_ctx(ctx)?;
    let (sign, abs_im) = if z.im.is_sign_negative() {
        ('-', -z.im)
    } else {
        ('+', z.im)
    };

    let im_coeff = if abs_im == 1.0 {
        String::new()
    } else {
        Value::Number(abs_im).coerce_to_string_with_ctx(ctx)?
    };

    Ok(format!("{re_str}{sign}{im_coeff}{suffix}"))
}

fn format_imag_only(im: f64, suffix: char, ctx: &dyn FunctionContext) -> Result<String, ErrorKind> {
    if im == 1.0 {
        return Ok(suffix.to_string());
    }
    if im == -1.0 {
        return Ok(format!("-{suffix}"));
    }
    let coeff = Value::Number(im).coerce_to_string_with_ctx(ctx)?;
    Ok(format!("{coeff}{suffix}"))
}
