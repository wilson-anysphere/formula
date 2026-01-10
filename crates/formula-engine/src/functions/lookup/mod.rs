use crate::functions::text;
use crate::{ErrorKind, Value};

fn values_equal_for_lookup(lookup_value: &Value, candidate: &Value) -> bool {
    match (lookup_value, candidate) {
        (Value::Number(a), Value::Number(b)) => a == b,
        (Value::Text(a), Value::Text(b)) => a.eq_ignore_ascii_case(b),
        (Value::Bool(a), Value::Bool(b)) => a == b,
        (Value::Error(a), Value::Error(b)) => a == b,
        (Value::Number(a), Value::Text(b)) | (Value::Text(b), Value::Number(a)) => {
            text::value(b).is_ok_and(|parsed| parsed == *a)
        }
        (Value::Bool(a), Value::Number(b)) | (Value::Number(b), Value::Bool(a)) => {
            (*b == 0.0 && !a) || (*b == 1.0 && *a)
        }
        (Value::Blank, Value::Blank) => true,
        _ => false,
    }
}

/// XMATCH(lookup_value, lookup_array)
///
/// This implements the most common mode: exact match, searching first-to-last.
/// Returns a 1-based index on success, or `#N/A` when no match is found.
pub fn xmatch(lookup_value: &Value, lookup_array: &[Value]) -> Result<i32, ErrorKind> {
    for (idx, candidate) in lookup_array.iter().enumerate() {
        if values_equal_for_lookup(lookup_value, candidate) {
            return Ok(i32::try_from(idx + 1).unwrap_or(i32::MAX));
        }
    }
    Err(ErrorKind::NA)
}

/// XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])
///
/// Implements the most common mode: exact match, searching first-to-last.
pub fn xlookup(
    lookup_value: &Value,
    lookup_array: &[Value],
    return_array: &[Value],
    if_not_found: Option<Value>,
) -> Result<Value, ErrorKind> {
    if lookup_array.len() != return_array.len() {
        return Err(ErrorKind::Value);
    }

    match xmatch(lookup_value, lookup_array) {
        Ok(pos) => {
            let idx = usize::try_from(pos - 1).map_err(|_| ErrorKind::Value)?;
            return_array.get(idx).cloned().ok_or(ErrorKind::Value)
        }
        Err(ErrorKind::NA) => if_not_found.ok_or(ErrorKind::NA),
        Err(e) => Err(e),
    }
}
