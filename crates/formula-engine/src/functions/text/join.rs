use crate::date::ExcelDateSystem;
use crate::locale::ValueLocaleConfig;
use crate::value::format_number_general_with_options;
use crate::{ErrorKind, Value};
use formula_format::Locale;

/// TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
pub fn textjoin(
    delimiter: &str,
    ignore_empty: bool,
    values: &[Value],
    date_system: ExcelDateSystem,
    value_locale: ValueLocaleConfig,
) -> Result<String, ErrorKind> {
    let locale = value_locale.separators;

    let mut out = String::new();
    let mut first = true;

    for value in values {
        let piece = match value_to_text(value, locale, date_system) {
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

fn value_to_text(
    value: &Value,
    locale: Locale,
    date_system: ExcelDateSystem,
) -> Result<String, ErrorKind> {
    match value {
        Value::Blank => Ok(String::new()),
        Value::Number(n) => Ok(format_number_general_with_options(*n, locale, date_system)),
        Value::Text(s) => Ok(s.clone()),
        Value::Bool(b) => Ok(if *b {
            "TRUE".to_string()
        } else {
            "FALSE".to_string()
        }),
        Value::Entity(entity) => Ok(entity.display.clone()),
        Value::Record(record) => {
            if let Some(display_field) = record.display_field.as_deref() {
                if let Some(value) = record.get_field_case_insensitive(display_field) {
                    return value_to_text(&value, locale, date_system);
                }
            }
            Ok(record.display.clone())
        }
        Value::Error(e) => Err(*e),
        other => {
            if matches!(other, Value::Spill { .. }) {
                return Err(ErrorKind::Spill);
            }
            if matches!(
                other,
                Value::Reference(_) | Value::ReferenceUnion(_) | Value::Lambda(_)
            ) {
                return Err(ErrorKind::Value);
            }
            if let Value::Array(arr) = other {
                return value_to_text(&arr.top_left(), locale, date_system);
            }

            // Other rich scalar values: treat as their display string.
            Ok(other.to_string())
        }
    }
}
