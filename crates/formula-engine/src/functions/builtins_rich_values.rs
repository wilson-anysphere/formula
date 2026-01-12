use std::cmp::Ordering;
use std::collections::HashMap;

use crate::eval::CompiledExpr;
use crate::functions::array_lift;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec, ThreadSafety, ValueType, Volatility};
use crate::value::{casefold, cmp_case_insensitive, Array, ErrorKind, Value};

pub(crate) static FIELDACCESS_SPEC: FunctionSpec = FunctionSpec {
    name: "_FIELDACCESS",
    min_args: 2,
    max_args: 2,
    volatility: Volatility::NonVolatile,
    thread_safety: ThreadSafety::ThreadSafe,
    array_support: ArraySupport::SupportsArrays,
    return_type: ValueType::Any,
    arg_types: &[ValueType::Any, ValueType::Any],
    implementation: fieldaccess_fn,
};

fn fieldaccess_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let base = array_lift::eval_arg(ctx, &args[0]);

    // The canonical lowering for the field access operator always passes a text literal, but
    // allow direct `_FIELDACCESS` calls to supply any scalar value and coerce via Excel's
    // normal "to text" semantics.
    let field_value = eval_scalar_arg(ctx, &args[1]);
    let field = match &field_value {
        Value::Text(s) => s.clone(),
        _ => match field_value.coerce_to_string_with_ctx(ctx) {
            Ok(s) => s,
            Err(e) => return Value::Error(e),
        },
    };

    match base {
        Value::Error(e) => Value::Error(e),
        Value::Array(arr) => {
            let mut out = Vec::with_capacity(arr.values.len());
            for elem in &arr.values {
                out.push(fieldaccess_scalar(elem, &field));
            }
            Value::Array(Array::new(arr.rows, arr.cols, out))
        }
        other => fieldaccess_scalar(&other, &field),
    }
}

fn fieldaccess_scalar(base: &Value, field: &str) -> Value {
    match base {
        Value::Error(e) => Value::Error(*e),
        Value::Entity(entity) => match map_get_case_insensitive(&entity.fields, field) {
            Some(v) => v.clone(),
            None => Value::Error(ErrorKind::Field),
        },
        Value::Record(record) => match map_get_case_insensitive(&record.fields, field) {
            Some(v) => v.clone(),
            None => Value::Error(ErrorKind::Field),
        },
        // Field access on a non-rich value is a type error.
        _ => Value::Error(ErrorKind::Value),
    }
}

fn map_get_case_insensitive<'a>(map: &'a HashMap<String, Value>, key: &str) -> Option<&'a Value> {
    if let Some(v) = map.get(key) {
        return Some(v);
    }

    // Fast-path lookups for callers (or builders) that store case-folded keys.
    let folded = casefold(key);
    if let Some(v) = map.get(&folded) {
        return Some(v);
    }

    map.iter()
        .find(|(k, _)| cmp_case_insensitive(k, key) == Ordering::Equal)
        .map(|(_, v)| v)
}
