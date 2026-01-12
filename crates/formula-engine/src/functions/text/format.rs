use crate::error::{ExcelError, ExcelResult};
use crate::date::ExcelDateSystem;
use crate::{ErrorKind, Value};
use formula_format::{DateSystem, FormatOptions, Locale, Value as FmtValue};

/// DOLLAR(number, [decimals])
pub fn dollar(number: f64, decimals: Option<i32>) -> ExcelResult<String> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let decimals = decimals.unwrap_or(2);
    if decimals < -127 || decimals > 127 {
        return Err(ExcelError::Num);
    }

    let negative = number.is_sign_negative();
    let abs = number.abs();

    let rounded = round_to_decimals(abs, decimals);
    let formatted = format_fixed_with_grouping(rounded, decimals.max(0) as usize);

    let with_symbol = format!("${formatted}");
    Ok(if negative {
        format!("({with_symbol})")
    } else {
        with_symbol
    })
}

fn round_to_decimals(value: f64, decimals: i32) -> f64 {
    if decimals >= 0 {
        let factor = 10_f64.powi(decimals);
        (value * factor).round() / factor
    } else {
        let factor = 10_f64.powi(-decimals);
        (value / factor).round() * factor
    }
}

fn format_fixed_with_grouping(value: f64, decimals: usize) -> String {
    let mut s = format!("{:.*}", decimals, value);
    if let Some(dot) = s.find('.') {
        let int_part = &s[..dot];
        let frac_part = &s[dot + 1..];
        let grouped = group_thousands(int_part);
        s = if decimals == 0 {
            grouped
        } else {
            format!("{grouped}.{frac_part}")
        };
    } else {
        s = group_thousands(&s);
    }
    s
}

fn group_thousands(int_part: &str) -> String {
    let bytes = int_part.as_bytes();
    let len = bytes.len();
    if len <= 3 {
        return int_part.to_string();
    }

    let mut out = String::with_capacity(len + len / 3);
    let mut first_group = len % 3;
    if first_group == 0 {
        first_group = 3;
    }
    out.push_str(&int_part[..first_group]);
    let mut idx = first_group;
    while idx < len {
        out.push(',');
        out.push_str(&int_part[idx..idx + 3]);
        idx += 3;
    }
    out
}

/// TEXT(value, format_text)
///
/// Excel's TEXT function uses Excel's number format code language (the same one
/// used for cell display formatting). Delegate formatting to `formula-format` so
/// we support:
/// - multi-section formats: `pos;neg;zero;text`
/// - conditions: `[<1]...`
/// - date/time formats: `m/d/yyyy`, etc.
/// - locale overrides: `[$-409]...`
/// - text placeholders (`@`), fill/underscore tokens, and more
pub fn text(value: &Value, format_text: &str, date_system: ExcelDateSystem) -> Result<String, ErrorKind> {
    // Match existing semantics: propagate errors/spills and treat non-finite
    // numbers as Excel #NUM! (Excel doesn't have NaN/Inf cell values).
    let fmt_value = match value {
        Value::Error(e) => return Err(*e),
        Value::Spill { .. } => return Err(ErrorKind::Spill),
        Value::Reference(_) | Value::ReferenceUnion(_) | Value::Lambda(_) => return Err(ErrorKind::Value),
        Value::Array(arr) => return text(&arr.top_left(), format_text, date_system),
        Value::Number(n) => {
            if !n.is_finite() {
                return Err(ErrorKind::Num);
            }
            FmtValue::Number(*n)
        }
        Value::Text(s) => FmtValue::Text(s.as_str()),
        Value::Bool(b) => FmtValue::Bool(*b),
        Value::Blank => FmtValue::Blank,
    };

    let options = FormatOptions {
        locale: Locale::en_us(),
        date_system: match date_system {
            // `formula-format` always uses the Lotus 1-2-3 leap-year bug behavior
            // for the 1900 date system (Excel compatibility).
            ExcelDateSystem::Excel1900 { .. } => DateSystem::Excel1900,
            ExcelDateSystem::Excel1904 => DateSystem::Excel1904,
        },
    };

    Ok(formula_format::format_value(fmt_value, Some(format_text), &options).text)
}
