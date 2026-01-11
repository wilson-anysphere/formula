use crate::error::{ExcelError, ExcelResult};
use crate::{ErrorKind, Value};

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
/// Excel's full formatting language is extensive; this implementation supports a
/// high-coverage subset used by the P0 corpus:
/// - Integer and fixed-decimal patterns like `0`, `0.00`
/// - Thousands grouping with `#,##0` / `#,##0.00`
/// - Percent formatting like `0%` / `0.00%`
/// - Currency patterns starting with `$`
pub fn text(value: &Value, format_text: &str) -> Result<String, ErrorKind> {
    match value {
        Value::Error(e) => Err(*e),
        Value::Text(s) => Ok(s.clone()),
        Value::Blank => Ok(String::new()),
        Value::Bool(b) => Ok(if *b {
            "TRUE".to_string()
        } else {
            "FALSE".to_string()
        }),
        Value::Number(n) => format_number_with_pattern(*n, format_text),
        Value::Reference(_) | Value::ReferenceUnion(_) => Err(ErrorKind::Value),
        Value::Array(arr) => text(&arr.top_left(), format_text),
        Value::Lambda(_) => Err(ErrorKind::Value),
        Value::Spill { .. } => Err(ErrorKind::Spill),
    }
}

fn format_number_with_pattern(number: f64, format_text: &str) -> Result<String, ErrorKind> {
    if !number.is_finite() {
        return Err(ErrorKind::Num);
    }

    let fmt = format_text.trim();
    if fmt.is_empty() || fmt.eq_ignore_ascii_case("GENERAL") || fmt == "@" {
        return Ok(number.to_string());
    }

    let mut prefix = "";
    let mut body = fmt;
    if let Some(rest) = fmt.strip_prefix('$') {
        prefix = "$";
        body = rest;
    }

    let percent = body.ends_with('%');
    if percent {
        body = body.trim_end_matches('%');
    }

    let decimals = body
        .split_once('.')
        .map(|(_, frac)| frac.chars().filter(|c| *c == '0' || *c == '#').count())
        .unwrap_or(0);

    let grouping = body.contains(',');

    let mut value = number;
    if percent {
        value *= 100.0;
    }

    let formatted = if grouping {
        format_fixed_with_grouping(value.abs(), decimals)
    } else {
        format!("{:.*}", decimals, value.abs())
    };

    let mut out = String::new();
    if number.is_sign_negative() {
        out.push('-');
    }
    out.push_str(prefix);
    out.push_str(&formatted);
    if percent {
        out.push('%');
    }
    Ok(out)
}
