use super::builtins_helpers::excel_result_number;
use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::Value;

inventory::submit! {
    FunctionSpec {
        name: "ISPMT",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: ispmt_fn,
    }
}

fn ispmt_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rate = match eval_scalar_arg(ctx, &args[0]).coerce_to_number_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let per = match eval_scalar_arg(ctx, &args[1]).coerce_to_number_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let nper = match eval_scalar_arg(ctx, &args[2]).coerce_to_number_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pv = match eval_scalar_arg(ctx, &args[3]).coerce_to_number_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::ispmt(rate, per, nper, pv))
}

inventory::submit! {
    FunctionSpec {
        name: "DOLLARDE",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: dollarde_fn,
    }
}

fn dollarde_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let fractional_dollar = match eval_scalar_arg(ctx, &args[0]).coerce_to_number_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let fraction = match eval_scalar_arg(ctx, &args[1]).coerce_to_number_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::dollarde(fractional_dollar, fraction))
}

inventory::submit! {
    FunctionSpec {
        name: "DOLLARFR",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: dollarfr_fn,
    }
}

fn dollarfr_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let decimal_dollar = match eval_scalar_arg(ctx, &args[0]).coerce_to_number_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let fraction = match eval_scalar_arg(ctx, &args[1]).coerce_to_number_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::dollarfr(decimal_dollar, fraction))
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function (indirectly via
// `financial::__force_link`) ensures the module (and its `inventory::submit!` entries) are
// retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
