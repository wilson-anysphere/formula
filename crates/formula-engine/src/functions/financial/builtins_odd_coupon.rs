use crate::error::{ExcelError, ExcelResult};
use crate::eval::CompiledExpr;
use crate::functions::{ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

fn excel_result_number(res: ExcelResult<f64>) -> Value {
    match res {
        Ok(n) => Value::Number(n),
        Err(e) => Value::Error(match e {
            ExcelError::Div0 => ErrorKind::Div0,
            ExcelError::Value => ErrorKind::Value,
            ExcelError::Num => ErrorKind::Num,
        }),
    }
}

fn eval_finite_number_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<f64, ErrorKind> {
    let v = ctx.eval_scalar(expr);
    match v {
        Value::Error(e) => Err(e),
        other => {
            let n = other.coerce_to_number_with_ctx(ctx)?;
            if n.is_finite() {
                Ok(n)
            } else {
                Err(ErrorKind::Num)
            }
        }
    }
}

fn eval_i32_trunc(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<i32, ErrorKind> {
    let n = eval_finite_number_arg(ctx, expr)?;
    if n < (i32::MIN as f64) || n > (i32::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(n.trunc() as i32)
}

fn eval_optional_i32_trunc(
    ctx: &dyn FunctionContext,
    expr: Option<&CompiledExpr>,
) -> Result<Option<i32>, ErrorKind> {
    match expr {
        Some(e) => Ok(Some(eval_i32_trunc(ctx, e)?)),
        None => Ok(None),
    }
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
            ValueType::Any,   // settlement
            ValueType::Any,   // maturity
            ValueType::Any,   // issue
            ValueType::Any,   // first_coupon
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
    let system = ctx.date_system();

    let settlement = match eval_i32_trunc(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_i32_trunc(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let issue = match eval_i32_trunc(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let first_coupon = match eval_i32_trunc(ctx, &args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let rate = match eval_finite_number_arg(ctx, &args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match eval_finite_number_arg(ctx, &args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match eval_finite_number_arg(ctx, &args[6]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc(ctx, &args[7]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match eval_optional_i32_trunc(ctx, args.get(8)) {
        Ok(v) => v.unwrap_or(0),
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
            ValueType::Any,   // settlement
            ValueType::Any,   // maturity
            ValueType::Any,   // issue
            ValueType::Any,   // first_coupon
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
    let system = ctx.date_system();

    let settlement = match eval_i32_trunc(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_i32_trunc(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let issue = match eval_i32_trunc(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let first_coupon = match eval_i32_trunc(ctx, &args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let rate = match eval_finite_number_arg(ctx, &args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match eval_finite_number_arg(ctx, &args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match eval_finite_number_arg(ctx, &args[6]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc(ctx, &args[7]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match eval_optional_i32_trunc(ctx, args.get(8)) {
        Ok(v) => v.unwrap_or(0),
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
            ValueType::Any,   // settlement
            ValueType::Any,   // maturity
            ValueType::Any,   // last_interest
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
    let system = ctx.date_system();

    let settlement = match eval_i32_trunc(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_i32_trunc(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let last_interest = match eval_i32_trunc(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let rate = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match eval_finite_number_arg(ctx, &args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match eval_finite_number_arg(ctx, &args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc(ctx, &args[6]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match eval_optional_i32_trunc(ctx, args.get(7)) {
        Ok(v) => v.unwrap_or(0),
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
            ValueType::Any,   // settlement
            ValueType::Any,   // maturity
            ValueType::Any,   // last_interest
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
    let system = ctx.date_system();

    let settlement = match eval_i32_trunc(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_i32_trunc(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let last_interest = match eval_i32_trunc(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let rate = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match eval_finite_number_arg(ctx, &args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match eval_finite_number_arg(ctx, &args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc(ctx, &args[6]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match eval_optional_i32_trunc(ctx, args.get(7)) {
        Ok(v) => v.unwrap_or(0),
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
