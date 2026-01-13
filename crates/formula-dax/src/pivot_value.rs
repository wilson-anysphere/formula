use crate::Value;
use formula_model::pivots::PivotValue;

/// Converts a DAX [`Value`] into the canonical pivot scalar [`PivotValue`].
pub fn dax_value_to_pivot_value(v: &Value) -> PivotValue {
    match v {
        Value::Blank => PivotValue::Blank,
        Value::Number(n) => PivotValue::Number(n.0),
        Value::Boolean(b) => PivotValue::Bool(*b),
        Value::Text(s) => PivotValue::Text(s.as_ref().to_string()),
    }
}

/// Converts a canonical pivot scalar [`PivotValue`] into a DAX [`Value`].
///
/// DAX currently lacks a dedicated date scalar type in `formula-dax`. Until it does, pivot dates
/// are converted to `Value::Text` using the ISO `YYYY-MM-DD` string form produced by
/// `chrono::NaiveDate::to_string()`.
pub fn pivot_value_to_dax_value(v: &PivotValue) -> Value {
    match v {
        PivotValue::Blank => Value::Blank,
        PivotValue::Number(n) => Value::from(*n),
        PivotValue::Bool(b) => Value::from(*b),
        PivotValue::Text(s) => Value::from(s.clone()),
        PivotValue::Date(d) => Value::from(d.to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_dax_value_to_pivot_value() {
        assert_eq!(dax_value_to_pivot_value(&Value::Blank), PivotValue::Blank);
        assert_eq!(
            dax_value_to_pivot_value(&Value::from(1.25)),
            PivotValue::Number(1.25)
        );
        assert_eq!(
            dax_value_to_pivot_value(&Value::from(true)),
            PivotValue::Bool(true)
        );
        assert_eq!(
            dax_value_to_pivot_value(&Value::from("hello")),
            PivotValue::Text("hello".to_string())
        );
    }

    #[test]
    fn test_pivot_value_to_dax_value() {
        assert_eq!(pivot_value_to_dax_value(&PivotValue::Blank), Value::Blank);
        assert_eq!(
            pivot_value_to_dax_value(&PivotValue::Number(2.5)),
            Value::from(2.5)
        );
        assert_eq!(
            pivot_value_to_dax_value(&PivotValue::Bool(false)),
            Value::from(false)
        );
        assert_eq!(
            pivot_value_to_dax_value(&PivotValue::Text("hi".to_string())),
            Value::from("hi")
        );
    }
}
