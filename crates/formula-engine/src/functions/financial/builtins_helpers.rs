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

pub(super) fn coerce_to_finite_number(ctx: &dyn FunctionContext, v: &Value) -> Result<f64, ErrorKind> {
    let n = v.coerce_to_number_with_ctx(ctx)?;
    if !n.is_finite() {
        return Err(ErrorKind::Num);
    }
    Ok(n)
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
    let n = coerce_to_finite_number(ctx, &v)?;
    let t = n.trunc();
    if t < (i32::MIN as f64) || t > (i32::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    let basis = t as i32;
    super::coupon_schedule::validate_basis(basis).map_err(excel_error_kind)
}

pub(super) fn datevalue_from_value(
    ctx: &dyn FunctionContext,
    value: &Value,
    system: ExcelDateSystem,
    now_utc: DateTime<Utc>,
) -> Result<i32, ErrorKind> {
    match value {
        Value::Text(s) => date_time::datevalue(s, ctx.value_locale(), now_utc, system).map_err(excel_error_kind),
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
