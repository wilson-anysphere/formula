use super::builtins_helpers::excel_result_number;
use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::Value;

inventory::submit! {
    FunctionSpec {
        name: "PDURATION",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: pduration_fn,
    }
}

fn pduration_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rate = match eval_scalar_arg(ctx, &args[0]).coerce_to_number_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pv = match eval_scalar_arg(ctx, &args[1]).coerce_to_number_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let fv = match eval_scalar_arg(ctx, &args[2]).coerce_to_number_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::pduration(rate, pv, fv))
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function (indirectly via
// `financial::__force_link`) ensures the module (and its `inventory::submit!` entries) are
// retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
