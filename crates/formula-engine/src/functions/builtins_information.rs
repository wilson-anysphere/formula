use crate::eval::{CellAddr, CompiledExpr};
use crate::functions::{ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{Array, ErrorKind, Value};

use crate::functions::information;

fn map_array<F>(arr: &Array, f: F) -> Value
where
    F: Fn(&Value) -> Value + Copy,
{
    let total = arr.values.len();
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        debug_assert!(
            false,
            "information builtin exceeds materialization limit (cells={total})"
        );
        return Value::Error(ErrorKind::Spill);
    }

    let mut out: Vec<Value> = Vec::new();
    if out.try_reserve_exact(total).is_err() {
        debug_assert!(
            false,
            "information builtin allocation failed (cells={total})"
        );
        return Value::Error(ErrorKind::Num);
    }
    for v in arr.iter() {
        out.push(f(v));
    }
    Value::Array(Array::new(arr.rows, arr.cols, out))
}

fn map_value<F>(value: &Value, f: F) -> Value
where
    F: Fn(&Value) -> Value + Copy,
{
    match value {
        Value::Array(arr) => map_array(arr, f),
        other => f(other),
    }
}

fn map_arg<F>(ctx: &dyn FunctionContext, expr: &CompiledExpr, f: F) -> Value
where
    F: Fn(&Value) -> Value + Copy,
{
    match ctx.eval_arg(expr) {
        ArgValue::Scalar(v) => map_value(&v, f),
        ArgValue::Reference(r) => {
            let r = r.normalized();
            ctx.record_reference(&r);
            if r.is_single_cell() {
                let v = ctx.get_cell_value(&r.sheet_id, r.start);
                map_value(&v, f)
            } else {
                let rows = (r.end.row - r.start.row + 1) as usize;
                let cols = (r.end.col - r.start.col + 1) as usize;
                let total = match rows.checked_mul(cols) {
                    Some(v) => v,
                    None => return Value::Error(ErrorKind::Spill),
                };
                if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
                    return Value::Error(ErrorKind::Spill);
                }
                let mut values = Vec::new();
                if values.try_reserve_exact(total).is_err() {
                    return Value::Error(ErrorKind::Num);
                }
                for row in r.start.row..=r.end.row {
                    for col in r.start.col..=r.end.col {
                        let v = ctx.get_cell_value(&r.sheet_id, CellAddr { row, col });
                        values.push(f(&v));
                    }
                }
                Value::Array(Array::new(rows, cols, values))
            }
        }
        // Discontiguous unions cannot be represented as a single rectangular array.
        ArgValue::ReferenceUnion(_) => Value::Error(ErrorKind::Value),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "ISBLANK",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: isblank_fn,
    }
}

fn isblank_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    map_arg(ctx, &args[0], |v| Value::Bool(information::isblank(v)))
}

inventory::submit! {
    FunctionSpec {
        name: "ISNUMBER",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: isnumber_fn,
    }
}

fn isnumber_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    map_arg(ctx, &args[0], |v| Value::Bool(information::isnumber(v)))
}

inventory::submit! {
    FunctionSpec {
        name: "ISTEXT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: istext_fn,
    }
}

fn istext_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    map_arg(ctx, &args[0], |v| Value::Bool(information::istext(v)))
}

inventory::submit! {
    FunctionSpec {
        name: "ISREF",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: isref_fn,
    }
}

fn isref_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    fn is_ref_value(value: &Value) -> bool {
        matches!(value, Value::Reference(_) | Value::ReferenceUnion(_))
    }

    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(_) | ArgValue::ReferenceUnion(_) => Value::Bool(true),
        ArgValue::Scalar(v) => match v {
            Value::Array(arr) => map_array(&arr, |v| Value::Bool(is_ref_value(v))),
            other => Value::Bool(is_ref_value(&other)),
        },
    }
}

inventory::submit! {
    FunctionSpec {
        name: "ISNONTEXT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: isnontext_fn,
    }
}

fn isnontext_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    map_arg(ctx, &args[0], |v| Value::Bool(!information::istext(v)))
}

inventory::submit! {
    FunctionSpec {
        name: "ISLOGICAL",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: islogical_fn,
    }
}

fn islogical_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    map_arg(ctx, &args[0], |v| Value::Bool(information::islogical(v)))
}

inventory::submit! {
    FunctionSpec {
        name: "ISNA",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: isna_fn,
    }
}

fn isna_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    map_arg(ctx, &args[0], |v| Value::Bool(information::isna(v)))
}

inventory::submit! {
    FunctionSpec {
        name: "ISERR",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: iserr_fn,
    }
}

fn iserr_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    map_arg(ctx, &args[0], |v| Value::Bool(information::iserr(v)))
}

inventory::submit! {
    FunctionSpec {
        name: "TYPE",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: type_fn,
    }
}

fn type_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let code = match ctx.eval_arg(&args[0]) {
        ArgValue::Scalar(v) => information::r#type(&v),
        ArgValue::Reference(r) => {
            let r = r.normalized();
            ctx.record_reference(&r);
            if r.is_single_cell() {
                information::r#type(&ctx.get_cell_value(&r.sheet_id, r.start))
            } else {
                64
            }
        }
        ArgValue::ReferenceUnion(_) => return Value::Error(ErrorKind::Value),
    };
    Value::Number(code as f64)
}

inventory::submit! {
    FunctionSpec {
        name: "ERROR.TYPE",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: error_type_fn,
    }
}

fn error_type_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    map_arg(ctx, &args[0], |v| match information::error_type(v) {
        Some(code) => Value::Number(code as f64),
        None => Value::Error(ErrorKind::NA),
    })
}

inventory::submit! {
    FunctionSpec {
        name: "N",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: n_fn,
    }
}

fn n_value(v: &Value) -> Value {
    match v {
        Value::Error(e) => Value::Error(*e),
        Value::Number(n) => Value::Number(*n),
        Value::Bool(b) => Value::Number(if *b { 1.0 } else { 0.0 }),
        Value::Blank | Value::Text(_) | Value::Entity(_) | Value::Record(_) => Value::Number(0.0),
        Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Array(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => Value::Error(ErrorKind::Value),
    }
}

fn n_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    map_arg(ctx, &args[0], n_value)
}

inventory::submit! {
    FunctionSpec {
        name: "T",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Any],
        implementation: t_fn,
    }
}

fn t_value(v: &Value) -> Value {
    fn text_like_value(value: &Value) -> Option<String> {
        match value {
            Value::Text(s) => Some(s.clone()),
            Value::Entity(entity) => Some(entity.display.clone()),
            Value::Record(record) => Some(record_display(record)),
            _ => None,
        }
    }

    fn record_display(record: &crate::value::RecordValue) -> String {
        if let Some(display_field) = record.display_field.as_deref() {
            if let Some(value) = record.get_field_case_insensitive(display_field) {
                if let Some(text) = text_like_value(&value) {
                    return text;
                }
            }
        }
        record.display.clone()
    }

    match v {
        Value::Error(e) => Value::Error(*e),
        Value::Text(s) => Value::Text(s.clone()),
        Value::Entity(entity) => Value::Text(entity.display.clone()),
        Value::Record(record) => Value::Text(record_display(record)),
        Value::Number(_)
        | Value::Bool(_)
        | Value::Blank
        | Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Array(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => Value::Text(String::new()),
    }
}

fn t_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    map_arg(ctx, &args[0], t_value)
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn map_value_maps_arrays_without_changing_shape() {
        let value = Value::Array(Array::new(
            2,
            2,
            vec![
                Value::Number(1.0),
                Value::Blank,
                Value::Number(3.0),
                Value::Number(4.0),
            ],
        ));

        let out = map_value(&value, |v| Value::Bool(matches!(v, Value::Number(_))));
        assert_eq!(
            out,
            Value::Array(Array::new(
                2,
                2,
                vec![
                    Value::Bool(true),
                    Value::Bool(false),
                    Value::Bool(true),
                    Value::Bool(true),
                ]
            ))
        );
    }
}
