use crate::error::ExcelResult;
use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

use super::builtins_helpers::{
    coerce_to_finite_number, datevalue_from_value, excel_error_kind, excel_result_number,
};

fn excel_result_serial(res: ExcelResult<i32>) -> Value {
    match res {
        Ok(n) => Value::Number(n as f64),
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

fn eval_date_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<i32, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    match v {
        Value::Error(e) => Err(e),
        other => datevalue_from_value(ctx, &other, ctx.date_system(), ctx.now_utc()),
    }
}

fn eval_finite_number_arg(
    ctx: &dyn FunctionContext,
    expr: &CompiledExpr,
) -> Result<f64, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    match v {
        Value::Error(e) => Err(e),
        other => coerce_to_finite_number(ctx, &other),
    }
}

fn eval_i32_trunc_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<i32, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    match v {
        Value::Error(e) => Err(e),
        other => coerce_to_i32_trunc(ctx, &other),
    }
}

// ---------------------------------------------------------------------
// COUP* schedule functions
// ---------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "COUPDAYBS",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: coupdaybs_fn,
    }
}

fn coupdaybs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let basis = match args.get(3) {
        Some(expr) => match eval_i32_trunc_arg(ctx, expr) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(super::coupdaybs(
        settlement,
        maturity,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "COUPDAYS",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: coupdays_fn,
    }
}

fn coupdays_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let basis = match args.get(3) {
        Some(expr) => match eval_i32_trunc_arg(ctx, expr) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(super::coupdays(
        settlement,
        maturity,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "COUPDAYSNC",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: coupdaysnc_fn,
    }
}

fn coupdaysnc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let basis = match args.get(3) {
        Some(expr) => match eval_i32_trunc_arg(ctx, expr) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(super::coupdaysnc(
        settlement,
        maturity,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "COUPNCD",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: coupncd_fn,
    }
}

fn coupncd_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let basis = match args.get(3) {
        Some(expr) => match eval_i32_trunc_arg(ctx, expr) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_serial(super::coupncd(
        settlement,
        maturity,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "COUPNUM",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: coupnum_fn,
    }
}

fn coupnum_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let basis = match args.get(3) {
        Some(expr) => match eval_i32_trunc_arg(ctx, expr) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(super::coupnum(
        settlement,
        maturity,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "COUPPCD",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: couppcd_fn,
    }
}

fn couppcd_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let basis = match args.get(3) {
        Some(expr) => match eval_i32_trunc_arg(ctx, expr) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_serial(super::couppcd(
        settlement,
        maturity,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

// ---------------------------------------------------------------------
// PRICE
// ---------------------------------------------------------------------
inventory::submit! {
    FunctionSpec {
        name: "PRICE",
        min_args: 6,
        max_args: 7,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any,
            ValueType::Any,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: price_fn,
    }
}

fn price_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match eval_finite_number_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match eval_finite_number_arg(ctx, &args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let basis = match args.get(6) {
        Some(expr) => match eval_i32_trunc_arg(ctx, expr) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(super::price(
        settlement,
        maturity,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

// ---------------------------------------------------------------------
// YIELD
// ---------------------------------------------------------------------
inventory::submit! {
    FunctionSpec {
        name: "YIELD",
        min_args: 6,
        max_args: 7,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any,
            ValueType::Any,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: yield_fn,
    }
}

fn yield_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match eval_finite_number_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match eval_finite_number_arg(ctx, &args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let basis = match args.get(6) {
        Some(expr) => match eval_i32_trunc_arg(ctx, expr) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(super::yield_rate(
        settlement,
        maturity,
        rate,
        pr,
        redemption,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

// ---------------------------------------------------------------------
// DURATION
// ---------------------------------------------------------------------
inventory::submit! {
    FunctionSpec {
        name: "DURATION",
        min_args: 5,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any,
            ValueType::Any,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: duration_fn,
    }
}

fn duration_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let coupon = match eval_finite_number_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let basis = match args.get(5) {
        Some(expr) => match eval_i32_trunc_arg(ctx, expr) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(super::duration(
        settlement,
        maturity,
        coupon,
        yld,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

// ---------------------------------------------------------------------
// MDURATION
// ---------------------------------------------------------------------
inventory::submit! {
    FunctionSpec {
        name: "MDURATION",
        min_args: 5,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any,
            ValueType::Any,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: mduration_fn,
    }
}

fn mduration_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let coupon = match eval_finite_number_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let basis = match args.get(5) {
        Some(expr) => match eval_i32_trunc_arg(ctx, expr) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        },
        None => 0,
    };

    excel_result_number(super::mduration(
        settlement,
        maturity,
        coupon,
        yld,
        frequency,
        basis,
        ctx.date_system(),
    ))
}
