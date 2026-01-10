use crate::Value;

/// Returns TRUE if the value is a blank cell.
pub fn isblank(value: &Value) -> bool {
    matches!(value, Value::Blank)
}

/// Returns TRUE if the value is any Excel error.
pub fn iserror(value: &Value) -> bool {
    matches!(value, Value::Error(_))
}

/// Returns TRUE if the value is a number.
pub fn isnumber(value: &Value) -> bool {
    matches!(value, Value::Number(n) if n.is_finite())
}

/// Returns TRUE if the value is text.
pub fn istext(value: &Value) -> bool {
    matches!(value, Value::Text(_))
}

/// Returns the numeric type code used by Excel's TYPE function.
///
/// Excel uses:
/// - 1: number
/// - 2: text
/// - 4: logical
/// - 16: error
/// - 64: array
pub fn r#type(value: &Value) -> i32 {
    match value {
        Value::Number(_) | Value::Blank => 1,
        Value::Text(_) => 2,
        Value::Bool(_) => 4,
        Value::Error(_) => 16,
    }
}
