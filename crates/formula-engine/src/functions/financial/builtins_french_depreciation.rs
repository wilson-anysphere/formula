use chrono::{DateTime, Utc};

use crate::coercion::ValueLocaleConfig;
use crate::date::ExcelDateSystem;
use crate::error::ExcelError;
use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, FunctionContext, FunctionSpec};
use crate::functions::{ArraySupport, ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

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
    let cfg = ctx.value_locale();

    let date_purchased = match datevalue_from_value(
        ctx,
        &eval_scalar_arg(ctx, &args[1]),
        system,
        cfg,
        now_utc,
    ) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let first_period = match datevalue_from_value(
        ctx,
        &eval_scalar_arg(ctx, &args[2]),
        system,
        cfg,
        now_utc,
    ) {
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
    let cfg = ctx.value_locale();

    let date_purchased = match datevalue_from_value(
        ctx,
        &eval_scalar_arg(ctx, &args[1]),
        system,
        cfg,
        now_utc,
    ) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let first_period = match datevalue_from_value(
        ctx,
        &eval_scalar_arg(ctx, &args[2]),
        system,
        cfg,
        now_utc,
    ) {
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

fn coerce_to_i32_trunc(ctx: &dyn FunctionContext, v: &Value) -> Result<i32, ErrorKind> {
    let n = coerce_to_finite_number(ctx, v)?;
    let t = n.trunc();
    if t < (i32::MIN as f64) || t > (i32::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(t as i32)
}

fn datevalue_from_value(
    ctx: &dyn FunctionContext,
    value: &Value,
    system: ExcelDateSystem,
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
) -> Result<i32, ErrorKind> {
    // Keep behavior aligned with `builtins_date_time.rs`:
    // - Text: DATEVALUE parsing with locale
    // - Numeric: floor to serial
    // - Non-finite numeric: #NUM!
    // - Unparseable text: #VALUE!
    match value {
        Value::Text(s) => crate::functions::date_time::datevalue(s, cfg, now_utc, system).map_err(excel_error_kind),
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

