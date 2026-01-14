use std::borrow::Cow;

use formula_format::{FormatOptions, Value as FmtValue};

use crate::value::RecordValue;
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
    enum DisplayValue<'a> {
        Number(f64),
        Text(Cow<'a, str>),
        Bool(bool),
        Blank,
        Error(&'static str),
    }

    fn value_to_display_string(
        value: Value,
        options: &FormatOptions,
    ) -> Result<String, &'static str> {
        match value {
            Value::Blank => Ok(String::new()),
            Value::Number(n) => {
                let fmt_value = FmtValue::Number(n);
                Ok(formula_format::format_value(fmt_value, None, options).text)
            }
            Value::Text(s) => Ok(s),
            Value::Entity(v) => Ok(v.display),
            Value::Record(v) => record_to_display_text(&v, options).map(|cow| cow.into_owned()),
            Value::Bool(b) => Ok(if b {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }),
            Value::Error(e) => Err(e.as_code()),
            Value::Reference(_) | Value::ReferenceUnion(_) => Err("#VALUE!"),
            Value::Array(arr) => value_to_display_string(arr.top_left(), options),
            Value::Lambda(_) => Err("#CALC!"),
            Value::Spill { .. } => Err("#SPILL!"),
        }
    }

    fn record_to_display_text<'a>(
        record: &'a RecordValue,
        options: &FormatOptions,
    ) -> Result<Cow<'a, str>, &'static str> {
        if let Some(display_field) = record.display_field.as_deref() {
            if let Some(value) = record.get_field_case_insensitive(display_field) {
                return value_to_display_string(value, options).map(Cow::Owned);
            }
        }

        Ok(Cow::Borrowed(record.display.as_str()))
    }

    fn to_display_value<'a>(value: &'a Value, options: &FormatOptions) -> DisplayValue<'a> {
        match value {
            Value::Number(n) => DisplayValue::Number(*n),
            Value::Text(s) => DisplayValue::Text(Cow::Borrowed(s.as_str())),
            Value::Entity(v) => DisplayValue::Text(Cow::Borrowed(v.display.as_str())),
            Value::Record(v) => match record_to_display_text(v, options) {
                Ok(text) => DisplayValue::Text(text),
                Err(err) => DisplayValue::Error(err),
            },
            Value::Bool(b) => DisplayValue::Bool(*b),
            Value::Blank => DisplayValue::Blank,
            Value::Error(e) => DisplayValue::Error(e.as_code()),
            Value::Reference(_) | Value::ReferenceUnion(_) => DisplayValue::Error("#VALUE!"),
            Value::Array(arr) => to_display_value(arr.get(0, 0).unwrap_or(&Value::Blank), options),
            Value::Lambda(_) => DisplayValue::Error("#CALC!"),
            Value::Spill { .. } => DisplayValue::Error("#SPILL!"),
        }
    }

    let display_value = to_display_value(value, options);
    match display_value {
        DisplayValue::Number(n) => {
            let fmt_value = FmtValue::Number(n);
            formula_format::format_value(fmt_value, format_code, options)
        }
        DisplayValue::Text(text) => {
            // Records are treated as text for formatting purposes so numeric format codes
            // don't reinterpret their display strings.
            formula_format::format_value(FmtValue::Text(text.as_ref()), format_code, options)
        }
        DisplayValue::Bool(b) => {
            formula_format::format_value(FmtValue::Bool(b), format_code, options)
        }
        DisplayValue::Blank => formula_format::format_value(FmtValue::Blank, format_code, options),
        DisplayValue::Error(err) => {
            formula_format::format_value(FmtValue::Error(err), format_code, options)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::value::{EntityValue, RecordValue};
    use formula_format::Locale;

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
    fn formats_record_using_display_field_when_present() {
        let mut record = RecordValue::new("Fallback");
        record.display_field = Some("Name".to_string());
        record
            .fields
            .insert("Name".to_string(), Value::Text("Apple".to_string()));

        let value = Value::Record(record);
        let formatted = format_value_for_display(&value, None, &FormatOptions::default());
        assert_eq!(formatted.text, "Apple");
    }

    #[test]
    fn formats_record_display_field_number_using_locale_options() {
        let mut record = RecordValue::new("Fallback");
        record.display_field = Some("Value".to_string());
        record
            .fields
            .insert("Value".to_string(), Value::Number(1.5));

        let value = Value::Record(record);
        let options = FormatOptions {
            locale: Locale::de_de(),
            ..FormatOptions::default()
        };
        let formatted = format_value_for_display(&value, None, &options);
        assert_eq!(formatted.text, "1,5");
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
