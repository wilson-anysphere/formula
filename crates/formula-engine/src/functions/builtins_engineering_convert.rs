use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

inventory::submit! {
    FunctionSpec {
        name: "CONVERT",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Text, ValueType::Text],
        implementation: convert_fn,
    }
}

fn convert_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = match eval_scalar_arg(ctx, &args[0]).coerce_to_number_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    // Excel surfaces non-finite numeric inputs as #NUM!.
    if !number.is_finite() {
        return Value::Error(ErrorKind::Num);
    }

    let from_unit = match eval_scalar_arg(ctx, &args[1]).coerce_to_string_with_ctx(ctx) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let to_unit = match eval_scalar_arg(ctx, &args[2]).coerce_to_string_with_ctx(ctx) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };

    match crate::functions::engineering::convert(number, &from_unit, &to_unit) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(e),
    }
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
