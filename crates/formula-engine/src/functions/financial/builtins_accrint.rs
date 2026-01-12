use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

use super::builtins_helpers::{coerce_to_finite_number, datevalue_from_value, excel_error_kind};

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
    let basis = if args.len() == 5 {
        eval_scalar_arg(ctx, &args[4])
    } else {
        Value::Blank
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

    let basis = if matches!(basis, Value::Blank) {
        0
    } else {
        match coerce_to_i32_trunc(ctx, &basis) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    };
    let basis = match super::coupon_schedule::validate_basis(basis) {
        Ok(v) => v,
        Err(e) => return Value::Error(excel_error_kind(e)),
    };

    match super::accrintm(issue, settlement, rate, par, basis, system) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
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

    let basis = if args.len() >= 7 {
        eval_scalar_arg(ctx, &args[6])
    } else {
        Value::Blank
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
    let frequency = match super::coupon_schedule::validate_frequency(frequency) {
        Ok(v) => v,
        Err(e) => return Value::Error(excel_error_kind(e)),
    };

    let basis = if matches!(basis, Value::Blank) {
        0
    } else {
        match coerce_to_i32_trunc(ctx, &basis) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    };
    let basis = match super::coupon_schedule::validate_basis(basis) {
        Ok(v) => v,
        Err(e) => return Value::Error(excel_error_kind(e)),
    };

    let calc_method = if matches!(calc_method, Value::Blank) {
        false
    } else {
        match coerce_to_bool_finite(ctx, &calc_method) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    };

    match super::accrint(issue, first_interest, settlement, rate, par, frequency, basis, calc_method, system) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

fn coerce_number_to_i32_trunc(n: f64) -> Result<i32, ErrorKind> {
    let t = n.trunc();
    if t < (i32::MIN as f64) || t > (i32::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(t as i32)
}

fn coerce_to_i32_trunc(ctx: &dyn FunctionContext, v: &Value) -> Result<i32, ErrorKind> {
    let n = coerce_to_finite_number(ctx, v)?;
    coerce_number_to_i32_trunc(n)
}

fn coerce_to_bool_finite(ctx: &dyn FunctionContext, v: &Value) -> Result<bool, ErrorKind> {
    match v {
        Value::Number(n) => {
            if !n.is_finite() {
                return Err(ErrorKind::Num);
            }
            Ok(*n != 0.0)
        }
        _ => v.coerce_to_bool_with_ctx(ctx),
    }
}
