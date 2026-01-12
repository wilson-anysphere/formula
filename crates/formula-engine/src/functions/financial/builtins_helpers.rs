use chrono::{DateTime, Utc};

use crate::date::ExcelDateSystem;
use crate::error::{ExcelError, ExcelResult};
use crate::eval::CompiledExpr;
use crate::functions::date_time;
use crate::functions::{eval_scalar_arg, FunctionContext};
use crate::value::{ErrorKind, Value};

pub(super) fn excel_error_kind(e: ExcelError) -> ErrorKind {
    match e {
        ExcelError::Div0 => ErrorKind::Div0,
        ExcelError::Value => ErrorKind::Value,
        ExcelError::Num => ErrorKind::Num,
    }
}

pub(super) fn excel_result_number(res: ExcelResult<f64>) -> Value {
    match res {
        Ok(n) => Value::Number(n),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

pub(super) fn excel_result_serial(res: ExcelResult<i32>) -> Value {
    match res {
        Ok(n) => Value::Number(n as f64),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

pub(super) fn coerce_to_finite_number(
    ctx: &dyn FunctionContext,
    v: &Value,
) -> Result<f64, ErrorKind> {
    let n = v.coerce_to_number_with_ctx(ctx)?;
    if !n.is_finite() {
        return Err(ErrorKind::Num);
    }
    Ok(n)
}

pub(super) fn coerce_number_to_i32_trunc(n: f64) -> Result<i32, ErrorKind> {
    let t = n.trunc();
    if t < (i32::MIN as f64) || t > (i32::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(t as i32)
}

pub(super) fn coerce_to_i32_trunc(ctx: &dyn FunctionContext, v: &Value) -> Result<i32, ErrorKind> {
    let n = coerce_to_finite_number(ctx, v)?;
    coerce_number_to_i32_trunc(n)
}

pub(super) fn coerce_to_bool_finite(
    ctx: &dyn FunctionContext,
    v: &Value,
) -> Result<bool, ErrorKind> {
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

pub(super) fn eval_date_arg(
    ctx: &dyn FunctionContext,
    expr: &CompiledExpr,
) -> Result<i32, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    match v {
        Value::Error(e) => Err(e),
        other => datevalue_from_value(ctx, &other, ctx.date_system(), ctx.now_utc()),
    }
}

pub(super) fn eval_finite_number_arg(
    ctx: &dyn FunctionContext,
    expr: &CompiledExpr,
) -> Result<f64, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    match v {
        Value::Error(e) => Err(e),
        other => coerce_to_finite_number(ctx, &other),
    }
}

pub(super) fn eval_i32_trunc_arg(
    ctx: &dyn FunctionContext,
    expr: &CompiledExpr,
) -> Result<i32, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    match v {
        Value::Error(e) => Err(e),
        other => coerce_to_i32_trunc(ctx, &other),
    }
}
pub(super) fn basis_from_optional_arg(
    ctx: &dyn FunctionContext,
    arg: Option<&CompiledExpr>,
) -> Result<i32, ErrorKind> {
    let Some(arg) = arg else {
        return Ok(0);
    };
    let v = eval_scalar_arg(ctx, arg);
    if matches!(v, Value::Blank) {
        return Ok(0);
    }
    let basis = coerce_to_i32_trunc(ctx, &v)?;
    super::coupon_schedule::validate_basis(basis).map_err(excel_error_kind)
}

pub(super) fn datevalue_from_value(
    ctx: &dyn FunctionContext,
    value: &Value,
    system: ExcelDateSystem,
    now_utc: DateTime<Utc>,
) -> Result<i32, ErrorKind> {
    match value {
        Value::Text(s) => {
            date_time::datevalue(s, ctx.value_locale(), now_utc, system).map_err(excel_error_kind)
        }
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
