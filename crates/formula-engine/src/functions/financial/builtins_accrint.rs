use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::Value;

use super::builtins_helpers::{
    basis_from_optional_arg, coerce_to_bool_finite, coerce_to_finite_number, coerce_to_i32_trunc,
    datevalue_from_value, excel_result_number,
};

inventory::submit! {
    FunctionSpec {
        name: "ACCRINTM",
        min_args: 4,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: accrintm_fn,
    }
}

fn accrintm_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let issue = eval_scalar_arg(ctx, &args[0]);
    let settlement = eval_scalar_arg(ctx, &args[1]);
    let rate = eval_scalar_arg(ctx, &args[2]);
    let par = eval_scalar_arg(ctx, &args[3]);
    let basis = match basis_from_optional_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let issue = match datevalue_from_value(ctx, &issue, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let rate = match coerce_to_finite_number(ctx, &rate) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let par = match coerce_to_finite_number(ctx, &par) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::accrintm(issue, settlement, rate, par, basis, system))
}

inventory::submit! {
    FunctionSpec {
        name: "ACCRINT",
        min_args: 6,
        max_args: 8,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any,
            ValueType::Any,
            ValueType::Any,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Bool,
        ],
        implementation: accrint_fn,
    }
}

fn accrint_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let issue = eval_scalar_arg(ctx, &args[0]);
    let first_interest = eval_scalar_arg(ctx, &args[1]);
    let settlement = eval_scalar_arg(ctx, &args[2]);
    let rate = eval_scalar_arg(ctx, &args[3]);
    let par = eval_scalar_arg(ctx, &args[4]);
    let frequency = eval_scalar_arg(ctx, &args[5]);

    let basis = match basis_from_optional_arg(ctx, args.get(6)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let calc_method = if args.len() == 8 {
        eval_scalar_arg(ctx, &args[7])
    } else {
        Value::Blank
    };

    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let issue = match datevalue_from_value(ctx, &issue, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let first_interest = match datevalue_from_value(ctx, &first_interest, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let rate = match coerce_to_finite_number(ctx, &rate) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let par = match coerce_to_finite_number(ctx, &par) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match coerce_to_i32_trunc(ctx, &frequency) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let calc_method = if matches!(calc_method, Value::Blank) {
        false
    } else {
        match coerce_to_bool_finite(ctx, &calc_method) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    };

    excel_result_number(super::accrint(
        issue,
        first_interest,
        settlement,
        rate,
        par,
        frequency,
        basis,
        calc_method,
        system,
    ))
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function (indirectly via
// `financial::__force_link`) ensures the module (and its `inventory::submit!` entries) are
// retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
