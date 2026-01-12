use formula_format::{FormatOptions, Value as FmtValue};

use crate::Value;

/// Format an evaluated [`Value`] into a user-visible string using an Excel
/// number format code.
///
/// The formatting rules are implemented in [`formula_format`]. This helper
/// exists to keep UI callers from needing to understand the engine's internal
/// value representation.
pub fn format_value_for_display(
    value: &Value,
    format_code: Option<&str>,
    options: &FormatOptions,
) -> formula_format::FormattedValue {
    fn to_fmt_value(value: &Value) -> FmtValue<'_> {
        match value {
            Value::Number(n) => FmtValue::Number(*n),
            Value::Text(s) => FmtValue::Text(s.as_str()),
            Value::Entity(v) => FmtValue::Text(v.display.as_str()),
            Value::Record(v) => FmtValue::Text(v.display.as_str()),
            Value::Bool(b) => FmtValue::Bool(*b),
            Value::Blank => FmtValue::Blank,
            Value::Error(e) => FmtValue::Error(e.as_code()),
            Value::Reference(_) | Value::ReferenceUnion(_) | Value::Record(_) => {
                FmtValue::Error("#VALUE!")
            }
            Value::Array(arr) => to_fmt_value(arr.get(0, 0).unwrap_or(&Value::Blank)),
            Value::Lambda(_) => FmtValue::Error("#CALC!"),
            Value::Spill { .. } => FmtValue::Error("#SPILL!"),
        }
    }

    let fmt_value = to_fmt_value(value);

    formula_format::format_value(fmt_value, format_code, options)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::value::{EntityValue, RecordValue};

    #[test]
    fn formats_entity_as_text_using_display_string() {
        let value = Value::Entity(EntityValue::new("Apple Inc."));

        let formatted = format_value_for_display(&value, None, &FormatOptions::default());
        assert_eq!(formatted.text, "Apple Inc.");
    }

    #[test]
    fn formats_record_as_text_using_display_string() {
        let value = Value::Record(RecordValue::new("Apple Inc."));

        let formatted = format_value_for_display(&value, None, &FormatOptions::default());
        assert_eq!(formatted.text, "Apple Inc.");
    }

    #[test]
    fn entity_and_record_to_string_use_display_string() {
        let entity = Value::Entity(EntityValue::new("Apple Inc."));
        assert_eq!(entity.to_string(), "Apple Inc.");

        let record = Value::Record(RecordValue::new("Apple Inc."));
        assert_eq!(record.to_string(), "Apple Inc.");
    }

    #[test]
    fn coerce_to_string_for_entity_and_record_returns_display_string() {
        let entity = Value::Entity(EntityValue::new("Apple Inc."));
        assert_eq!(entity.coerce_to_string().unwrap(), "Apple Inc.");

        let record = Value::Record(RecordValue::new("Apple Inc."));
        assert_eq!(record.coerce_to_string().unwrap(), "Apple Inc.");
    }
}
