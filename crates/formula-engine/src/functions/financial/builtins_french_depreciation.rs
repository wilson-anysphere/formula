use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, FunctionContext, FunctionSpec};
use crate::functions::{ArraySupport, ThreadSafety, ValueType, Volatility};
use crate::value::Value;

use super::builtins_helpers::{
    coerce_to_finite_number, coerce_to_i32_trunc, datevalue_from_value, excel_error_kind,
};

inventory::submit! {
    FunctionSpec {
        name: "AMORLINC",
        min_args: 6,
        max_args: 7,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Number, // cost
            ValueType::Any,    // date_purchased
            ValueType::Any,    // first_period
            ValueType::Number, // salvage
            ValueType::Number, // period
            ValueType::Number, // rate
            ValueType::Number, // basis
        ],
        implementation: amorlinc_fn,
    }
}

fn amorlinc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let cost = match coerce_to_finite_number(ctx, &eval_scalar_arg(ctx, &args[0])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let date_purchased =
        match datevalue_from_value(ctx, &eval_scalar_arg(ctx, &args[1]), system, now_utc) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
    let first_period =
        match datevalue_from_value(ctx, &eval_scalar_arg(ctx, &args[2]), system, now_utc) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };

    let salvage = match coerce_to_finite_number(ctx, &eval_scalar_arg(ctx, &args[3])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let period = match coerce_to_finite_number(ctx, &eval_scalar_arg(ctx, &args[4])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(ctx, &eval_scalar_arg(ctx, &args[5])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let basis_value = if args.len() == 7 {
        eval_scalar_arg(ctx, &args[6])
    } else {
        Value::Blank
    };
    let basis = if matches!(basis_value, Value::Blank) {
        None
    } else {
        match coerce_to_i32_trunc(ctx, &basis_value) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    };

    match super::amorlinc(
        cost,
        date_purchased,
        first_period,
        salvage,
        period,
        rate,
        basis,
        system,
    ) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "AMORDEGRC",
        min_args: 6,
        max_args: 7,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Number, // cost
            ValueType::Any,    // date_purchased
            ValueType::Any,    // first_period
            ValueType::Number, // salvage
            ValueType::Number, // period
            ValueType::Number, // rate
            ValueType::Number, // basis
        ],
        implementation: amordegrec_fn,
    }
}

fn amordegrec_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let cost = match coerce_to_finite_number(ctx, &eval_scalar_arg(ctx, &args[0])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let date_purchased =
        match datevalue_from_value(ctx, &eval_scalar_arg(ctx, &args[1]), system, now_utc) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
    let first_period =
        match datevalue_from_value(ctx, &eval_scalar_arg(ctx, &args[2]), system, now_utc) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };

    let salvage = match coerce_to_finite_number(ctx, &eval_scalar_arg(ctx, &args[3])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let period = match coerce_to_finite_number(ctx, &eval_scalar_arg(ctx, &args[4])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(ctx, &eval_scalar_arg(ctx, &args[5])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let basis_value = if args.len() == 7 {
        eval_scalar_arg(ctx, &args[6])
    } else {
        Value::Blank
    };
    let basis = if matches!(basis_value, Value::Blank) {
        None
    } else {
        match coerce_to_i32_trunc(ctx, &basis_value) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    };

    match super::amordegrec(
        cost,
        date_purchased,
        first_period,
        salvage,
        period,
        rate,
        basis,
        system,
    ) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}
