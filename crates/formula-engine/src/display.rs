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
            Value::Bool(b) => FmtValue::Bool(*b),
            Value::Blank => FmtValue::Blank,
            Value::Error(e) => FmtValue::Error(e.as_code()),
            Value::Array(arr) => to_fmt_value(arr.get(0, 0).unwrap_or(&Value::Blank)),
            Value::Lambda(_) => FmtValue::Error("#CALC!"),
            Value::Spill { .. } => FmtValue::Error("#SPILL!"),
        }
    }

    let fmt_value = to_fmt_value(value);

    formula_format::format_value(fmt_value, format_code, options)
}
