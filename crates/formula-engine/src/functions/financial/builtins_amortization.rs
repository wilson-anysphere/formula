use super::builtins_helpers::{eval_finite_number_arg, excel_result_number};
use crate::eval::CompiledExpr;
use crate::functions::{ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::Value;

inventory::submit! {
    FunctionSpec {
        name: "CUMIPMT",
        min_args: 6,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: cumipmt_fn,
    }
}

fn cumipmt_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rate = match eval_finite_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let nper = match eval_finite_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pv = match eval_finite_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let start_period = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let end_period = match eval_finite_number_arg(ctx, &args[4]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let typ = match eval_finite_number_arg(ctx, &args[5]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::amortization::cumipmt(
        rate,
        nper,
        pv,
        start_period,
        end_period,
        typ,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "CUMPRINC",
        min_args: 6,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: cumprinc_fn,
    }
}

fn cumprinc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rate = match eval_finite_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let nper = match eval_finite_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pv = match eval_finite_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let start_period = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let end_period = match eval_finite_number_arg(ctx, &args[4]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let typ = match eval_finite_number_arg(ctx, &args[5]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::amortization::cumprinc(
        rate,
        nper,
        pv,
        start_period,
        end_period,
        typ,
    ))
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function (indirectly via
// `financial::__force_link`) ensures the module (and its `inventory::submit!` entries) are
// retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
