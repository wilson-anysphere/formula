use crate::date::ExcelDateSystem;
use crate::locale::ValueLocaleConfig;
use crate::{ErrorKind, Value};
use formula_format::{DateSystem, FormatOptions, Value as FmtValue};

/// TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
pub fn textjoin(
    delimiter: &str,
    ignore_empty: bool,
    values: &[Value],
    date_system: ExcelDateSystem,
    value_locale: ValueLocaleConfig,
) -> Result<String, ErrorKind> {
    let options = FormatOptions {
        locale: value_locale.separators,
        date_system: match date_system {
            // `formula-format` always uses the Lotus 1-2-3 leap-year bug behavior
            // for the 1900 date system (Excel compatibility).
            ExcelDateSystem::Excel1900 { .. } => DateSystem::Excel1900,
            ExcelDateSystem::Excel1904 => DateSystem::Excel1904,
        },
    };

    let mut out = String::new();
    let mut first = true;

    for value in values {
        let piece = match value_to_text(value, &options) {
            Ok(piece) => piece,
            Err(e) => return Err(e),
        };

        if ignore_empty && piece.is_empty() {
            continue;
        }

        if !first {
            out.push_str(delimiter);
        }
        first = false;
        out.push_str(&piece);
    }

    Ok(out)
}

fn value_to_text(value: &Value, options: &FormatOptions) -> Result<String, ErrorKind> {
    match value {
        Value::Blank => Ok(String::new()),
        Value::Number(n) => Ok(formula_format::format_value(FmtValue::Number(*n), None, options).text),
        Value::Text(s) => Ok(s.clone()),
        Value::Bool(b) => Ok(if *b { "TRUE".to_string() } else { "FALSE".to_string() }),
        Value::Entity(entity) => Ok(entity.display.clone()),
        Value::Record(record) => {
            if let Some(display_field) = record.display_field.as_deref() {
                if let Some(value) = record.get_field_case_insensitive(display_field) {
                    return value_to_text(&value, options);
                }
            }
            Ok(record.display.clone())
        }
        Value::Error(e) => Err(*e),
        other => {
            if matches!(other, Value::Spill { .. }) {
                return Err(ErrorKind::Spill);
            }
            if matches!(other, Value::Reference(_) | Value::ReferenceUnion(_) | Value::Lambda(_)) {
                return Err(ErrorKind::Value);
            }
            if let Value::Array(arr) = other {
                return value_to_text(&arr.top_left(), options);
            }

            // Other rich scalar values: treat as their display string.
            Ok(other.to_string())
        }
    }
}
