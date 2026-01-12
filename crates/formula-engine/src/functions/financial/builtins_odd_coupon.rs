use chrono::{DateTime, Utc};

use crate::date::ExcelDateSystem;
use crate::error::{ExcelError, ExcelResult};
use crate::eval::CompiledExpr;
use crate::functions::date_time;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

fn excel_error_kind(e: ExcelError) -> ErrorKind {
    match e {
        ExcelError::Div0 => ErrorKind::Div0,
        ExcelError::Value => ErrorKind::Value,
        ExcelError::Num => ErrorKind::Num,
    }
}

fn excel_result_number(res: ExcelResult<f64>) -> Value {
    match res {
        Ok(n) => Value::Number(n),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

fn datevalue_from_value(
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

fn coerce_to_finite_number(ctx: &dyn FunctionContext, v: &Value) -> Result<f64, ErrorKind> {
    let n = v.coerce_to_number_with_ctx(ctx)?;
    if !n.is_finite() {
        return Err(ErrorKind::Num);
    }
    Ok(n)
}

fn basis_from_optional_arg(
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
    if !(0..=4).contains(&basis) {
        return Err(ErrorKind::Num);
    }
    Ok(basis)
}

fn frequency_from_value(ctx: &dyn FunctionContext, v: &Value) -> Result<i32, ErrorKind> {
    let n = coerce_to_finite_number(ctx, v)?;
    let t = n.trunc();
    if t < (i32::MIN as f64) || t > (i32::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    let frequency = t as i32;
    if !(frequency == 1 || frequency == 2 || frequency == 4) {
        return Err(ErrorKind::Num);
    }
    Ok(frequency)
}

inventory::submit! {
    FunctionSpec {
        name: "ODDFPRICE",
        min_args: 8,
        max_args: 9,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any, // settlement
            ValueType::Any, // maturity
            ValueType::Any, // issue
            ValueType::Any, // first_coupon
            ValueType::Number, // rate
            ValueType::Number, // yld
            ValueType::Number, // redemption
            ValueType::Number, // frequency
            ValueType::Number, // basis
        ],
        implementation: oddfprice_fn,
    }
}

fn oddfprice_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let issue = eval_scalar_arg(ctx, &args[2]);
    let first_coupon = eval_scalar_arg(ctx, &args[3]);
    let rate = eval_scalar_arg(ctx, &args[4]);
    let yld = eval_scalar_arg(ctx, &args[5]);
    let redemption = eval_scalar_arg(ctx, &args[6]);
    let frequency = eval_scalar_arg(ctx, &args[7]);
    let basis = match basis_from_optional_arg(ctx, args.get(8)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let issue = match datevalue_from_value(ctx, &issue, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let first_coupon = match datevalue_from_value(ctx, &first_coupon, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(ctx, &rate) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match coerce_to_finite_number(ctx, &yld) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(ctx, &redemption) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match frequency_from_value(ctx, &frequency) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "ODDFYIELD",
        min_args: 8,
        max_args: 9,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any, // settlement
            ValueType::Any, // maturity
            ValueType::Any, // issue
            ValueType::Any, // first_coupon
            ValueType::Number, // rate
            ValueType::Number, // pr
            ValueType::Number, // redemption
            ValueType::Number, // frequency
            ValueType::Number, // basis
        ],
        implementation: oddfyield_fn,
    }
}

fn oddfyield_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let issue = eval_scalar_arg(ctx, &args[2]);
    let first_coupon = eval_scalar_arg(ctx, &args[3]);
    let rate = eval_scalar_arg(ctx, &args[4]);
    let pr = eval_scalar_arg(ctx, &args[5]);
    let redemption = eval_scalar_arg(ctx, &args[6]);
    let frequency = eval_scalar_arg(ctx, &args[7]);
    let basis = match basis_from_optional_arg(ctx, args.get(8)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let issue = match datevalue_from_value(ctx, &issue, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let first_coupon = match datevalue_from_value(ctx, &first_coupon, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(ctx, &rate) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match coerce_to_finite_number(ctx, &pr) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(ctx, &redemption) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match frequency_from_value(ctx, &frequency) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        pr,
        redemption,
        frequency,
        basis,
        system,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "ODDLPRICE",
        min_args: 7,
        max_args: 8,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any, // settlement
            ValueType::Any, // maturity
            ValueType::Any, // last_interest
            ValueType::Number, // rate
            ValueType::Number, // yld
            ValueType::Number, // redemption
            ValueType::Number, // frequency
            ValueType::Number, // basis
        ],
        implementation: oddlprice_fn,
    }
}

fn oddlprice_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let last_interest = eval_scalar_arg(ctx, &args[2]);
    let rate = eval_scalar_arg(ctx, &args[3]);
    let yld = eval_scalar_arg(ctx, &args[4]);
    let redemption = eval_scalar_arg(ctx, &args[5]);
    let frequency = eval_scalar_arg(ctx, &args[6]);
    let basis = match basis_from_optional_arg(ctx, args.get(7)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let last_interest = match datevalue_from_value(ctx, &last_interest, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(ctx, &rate) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match coerce_to_finite_number(ctx, &yld) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(ctx, &redemption) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match frequency_from_value(ctx, &frequency) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "ODDLYIELD",
        min_args: 7,
        max_args: 8,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any, // settlement
            ValueType::Any, // maturity
            ValueType::Any, // last_interest
            ValueType::Number, // rate
            ValueType::Number, // pr
            ValueType::Number, // redemption
            ValueType::Number, // frequency
            ValueType::Number, // basis
        ],
        implementation: oddlyield_fn,
    }
}

fn oddlyield_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let last_interest = eval_scalar_arg(ctx, &args[2]);
    let rate = eval_scalar_arg(ctx, &args[3]);
    let pr = eval_scalar_arg(ctx, &args[4]);
    let redemption = eval_scalar_arg(ctx, &args[5]);
    let frequency = eval_scalar_arg(ctx, &args[6]);
    let basis = match basis_from_optional_arg(ctx, args.get(7)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let last_interest = match datevalue_from_value(ctx, &last_interest, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(ctx, &rate) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match coerce_to_finite_number(ctx, &pr) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(ctx, &redemption) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match frequency_from_value(ctx, &frequency) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::oddlyield(
        settlement,
        maturity,
        last_interest,
        rate,
        pr,
        redemption,
        frequency,
        basis,
        system,
    ))
}
