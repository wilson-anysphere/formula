use crate::date::ExcelDateSystem;
use crate::error::{ExcelError, ExcelResult};
use crate::locale::ValueLocaleConfig;
use crate::value::format_number_general_with_options;
use crate::{ErrorKind, Value};
use formula_format::{DateSystem, FormatOptions, Value as FmtValue};

/// DOLLAR(number, [decimals])
pub fn dollar(
    number: f64,
    decimals: Option<i32>,
    value_locale: ValueLocaleConfig,
) -> ExcelResult<String> {
    if !number.is_finite() {
        return Err(ExcelError::Num);
    }
    let decimals = decimals.unwrap_or(2);
    if decimals < -127 || decimals > 127 {
        return Err(ExcelError::Num);
    }

    let abs = number.abs();

    let rounded = round_to_decimals(abs, decimals);
    let signed = if number.is_sign_negative() {
        -rounded
    } else {
        rounded
    };

    let frac_digits = decimals.max(0) as usize;
    let mut pattern = String::from("$#,##0");
    if frac_digits > 0 {
        pattern.push('.');
        pattern.extend(std::iter::repeat('0').take(frac_digits));
    }
    let format_code = format!("{pattern};({pattern})");

    let options = FormatOptions {
        locale: value_locale.separators,
        // Date system doesn't affect numeric DOLLAR formatting; use Excel 1900 by default.
        date_system: DateSystem::Excel1900,
    };

    Ok(formula_format::format_value(FmtValue::Number(signed), Some(&format_code), &options).text)
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
pub fn text(
    value: &Value,
    format_text: &str,
    date_system: ExcelDateSystem,
    value_locale: ValueLocaleConfig,
) -> Result<String, ErrorKind> {
    fn value_to_display_string(
        value: &Value,
        options: &FormatOptions,
        date_system: ExcelDateSystem,
    ) -> Result<String, ErrorKind> {
        match value {
            Value::Blank => Ok(String::new()),
            Value::Number(n) => Ok(format_number_general_with_options(
                *n,
                options.locale,
                date_system,
            )),
            Value::Text(s) => Ok(s.clone()),
            Value::Entity(v) => Ok(v.display.clone()),
            Value::Record(v) => {
                if let Some(display_field) = v.display_field.as_deref() {
                    if let Some(value) = v.get_field_case_insensitive(display_field) {
                        return value_to_display_string(&value, options, date_system);
                    }
                }
                Ok(v.display.clone())
            }
            Value::Bool(b) => Ok(if *b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }),
            Value::Error(e) => Err(*e),
            Value::Spill { .. } => Err(ErrorKind::Spill),
            Value::Reference(_) | Value::ReferenceUnion(_) | Value::Lambda(_) => {
                Err(ErrorKind::Value)
            }
            Value::Array(arr) => value_to_display_string(&arr.top_left(), options, date_system),
        }
    }

    let options = FormatOptions {
        locale: value_locale.separators,
        date_system: match date_system {
            // `formula-format` always uses the Lotus 1-2-3 leap-year bug behavior
            // for the 1900 date system (Excel compatibility).
            ExcelDateSystem::Excel1900 { .. } => DateSystem::Excel1900,
            ExcelDateSystem::Excel1904 => DateSystem::Excel1904,
        },
    };

    // Match existing semantics: propagate errors/spills and treat non-finite
    // numbers as Excel #NUM! (Excel doesn't have NaN/Inf cell values).
    let fmt_value = match value {
        Value::Error(e) => return Err(*e),
        Value::Spill { .. } => return Err(ErrorKind::Spill),
        Value::Reference(_) | Value::ReferenceUnion(_) | Value::Lambda(_) => {
            return Err(ErrorKind::Value)
        }
        Value::Array(arr) => return text(&arr.top_left(), format_text, date_system, value_locale),
        Value::Number(n) => {
            if !n.is_finite() {
                return Err(ErrorKind::Num);
            }
            FmtValue::Number(*n)
        }
        Value::Text(s) => FmtValue::Text(s.as_str()),
        Value::Entity(v) => FmtValue::Text(v.display.as_str()),
        Value::Record(v) => {
            let display = if let Some(display_field) = v.display_field.as_deref() {
                if let Some(value) = v.get_field_case_insensitive(display_field) {
                    value_to_display_string(&value, &options, date_system)?
                } else {
                    v.display.clone()
                }
            } else {
                v.display.clone()
            };

            return Ok(formula_format::format_value(
                FmtValue::Text(display.as_str()),
                Some(format_text),
                &options,
            )
            .text);
        }
        Value::Bool(b) => FmtValue::Bool(*b),
        Value::Blank => FmtValue::Blank,
    };

    Ok(formula_format::format_value(fmt_value, Some(format_text), &options).text)
}
