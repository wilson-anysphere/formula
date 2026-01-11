use crate::{ErrorKind, Value};

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

/// Returns TRUE if the value is a boolean.
pub fn islogical(value: &Value) -> bool {
    matches!(value, Value::Bool(_))
}

/// Returns TRUE if the value is the `#N/A` error.
pub fn isna(value: &Value) -> bool {
    matches!(value, Value::Error(ErrorKind::NA))
}

/// Returns TRUE if the value is any error except `#N/A`.
pub fn iserr(value: &Value) -> bool {
    matches!(value, Value::Error(e) if *e != ErrorKind::NA)
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
        Value::Error(_) | Value::Lambda(_) => 16,
        Value::Array(_) | Value::Spill { .. } => 64,
    }
}

/// Returns the numeric error type codes used by Excel's ERROR.TYPE function.
///
/// Excel historically defined codes 1-8 for the classic errors. Modern Excel adds new
/// error values for dynamic arrays; this engine follows the extended mapping documented in
/// `docs/01-formula-engine.md`:
///
/// - 1: `#NULL!`
/// - 2: `#DIV/0!`
/// - 3: `#VALUE!`
/// - 4: `#REF!`
/// - 5: `#NAME?`
/// - 6: `#NUM!`
/// - 7: `#N/A`
/// - 9: `#SPILL!`
/// - 10: `#CALC!`
pub fn error_type_code(kind: ErrorKind) -> i32 {
    match kind {
        ErrorKind::Null => 1,
        ErrorKind::Div0 => 2,
        ErrorKind::Value => 3,
        ErrorKind::Ref => 4,
        ErrorKind::Name => 5,
        ErrorKind::Num => 6,
        ErrorKind::NA => 7,
        ErrorKind::Spill => 9,
        ErrorKind::Calc => 10,
    }
}

/// Returns the ERROR.TYPE numeric code for a value, or `None` if the value is not an error.
pub fn error_type(value: &Value) -> Option<i32> {
    match value {
        Value::Error(kind) => Some(error_type_code(*kind)),
        _ => None,
    }
}
