use crate::eval::CompiledExpr;
use crate::eval::MAX_MATERIALIZED_ARRAY_CELLS;
use crate::functions::array_lift;
use crate::functions::{
    eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec, ThreadSafety, ValueType,
    Volatility,
};
use crate::value::{Array, ErrorKind, Value};

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
            let total = arr.values.len();
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut out: Vec<Value> = Vec::new();
            if out.try_reserve_exact(total).is_err() {
                debug_assert!(false, "_FIELDACCESS array allocation failed (cells={total})");
                return Value::Error(ErrorKind::Num);
            }
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
        Value::Entity(entity) => entity
            .get_field_case_insensitive(field)
            .unwrap_or(Value::Error(ErrorKind::Field)),
        Value::Record(record) => record
            .get_field_case_insensitive(field)
            .unwrap_or(Value::Error(ErrorKind::Field)),
        // Field access on a non-rich value is a type error.
        _ => Value::Error(ErrorKind::Value),
    }
}
