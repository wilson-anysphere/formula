use crate::eval::CompiledExpr;
use crate::functions::engineering::special;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

fn eval_number_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<f64, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    match v.coerce_to_number_with_ctx(ctx) {
        Ok(n) if n.is_finite() => Ok(n),
        Ok(_) => Err(ErrorKind::Num),
        Err(e) => Err(e),
    }
}

fn eval_bessel_order(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<i32, ErrorKind> {
    let n = eval_number_arg(ctx, expr)?;
    if n.is_nan() || n.is_infinite() {
        return Err(ErrorKind::Num);
    }
    let n = n.trunc();
    if n < 0.0 || n > (i32::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(n as i32)
}

inventory::submit! {
    FunctionSpec {
        name: "ERF",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: erf_fn,
    }
}

fn erf_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let lower = match eval_number_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let lower_erf = match special::erf(lower) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    if args.len() == 1 {
        return Value::Number(lower_erf);
    }

    let upper = match eval_number_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let upper_erf = match special::erf(upper) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let out = upper_erf - lower_erf;
    if out.is_finite() {
        Value::Number(out)
    } else {
        Value::Error(ErrorKind::Num)
    }
}

inventory::submit! {
    FunctionSpec {
        name: "ERFC",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: erfc_fn,
    }
}

fn erfc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = match eval_number_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match special::erfc(x) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "BESSELJ",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: besselj_fn,
    }
}

fn besselj_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = match eval_number_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let n = match eval_bessel_order(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    match special::besselj(x, n) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "BESSELY",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: bessely_fn,
    }
}

fn bessely_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = match eval_number_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let n = match eval_bessel_order(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    match special::bessely(x, n) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "BESSELI",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: besseli_fn,
    }
}

fn besseli_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = match eval_number_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let n = match eval_bessel_order(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    match special::besseli(x, n) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "BESSELK",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: besselk_fn,
    }
}

fn besselk_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = match eval_number_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let n = match eval_bessel_order(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    match special::besselk(x, n) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}
