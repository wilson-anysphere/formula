use chrono::{DateTime, Utc};

use crate::coercion::ValueLocaleConfig;
use crate::date::ExcelDateSystem;
use crate::error::ExcelError;
use crate::eval::CompiledExpr;
use crate::functions::date_time;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

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
    let cfg = ctx.value_locale();

    let issue = match datevalue_from_value(ctx, &issue, system, cfg, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let settlement = match datevalue_from_value(ctx, &settlement, system, cfg, now_utc) {
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
    let cfg = ctx.value_locale();

    let issue = match datevalue_from_value(ctx, &issue, system, cfg, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let first_interest = match datevalue_from_value(ctx, &first_interest, system, cfg, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let settlement = match datevalue_from_value(ctx, &settlement, system, cfg, now_utc) {
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

    let basis = if matches!(basis, Value::Blank) {
        0
    } else {
        match coerce_to_i32_trunc(ctx, &basis) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    };

    let calc_method = if matches!(calc_method, Value::Blank) {
        true
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

fn excel_error_kind(err: ExcelError) -> ErrorKind {
    match err {
        ExcelError::Div0 => ErrorKind::Div0,
        ExcelError::Num => ErrorKind::Num,
        ExcelError::Value => ErrorKind::Value,
    }
}

fn coerce_to_finite_number(ctx: &dyn FunctionContext, v: &Value) -> Result<f64, ErrorKind> {
    let n = v.coerce_to_number_with_ctx(ctx)?;
    if !n.is_finite() {
        return Err(ErrorKind::Num);
    }
    Ok(n)
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

fn datevalue_from_value(
    ctx: &dyn FunctionContext,
    value: &Value,
    system: ExcelDateSystem,
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
) -> Result<i32, ErrorKind> {
    match value {
        Value::Text(s) => date_time::datevalue(s, cfg, now_utc, system).map_err(excel_error_kind),
        _ => {
            let n = value.coerce_to_number_with_ctx(ctx)?;
            if !n.is_finite() {
                return Err(ErrorKind::Num);
            }
            let serial = n.floor();
            if serial < (i32::MIN as f64) || serial > (i32::MAX as f64) {
                return Err(ErrorKind::Num);
            }
            Ok(serial as i32)
        }
    }
}

