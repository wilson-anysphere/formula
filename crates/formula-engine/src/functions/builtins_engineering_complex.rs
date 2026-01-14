use num_complex::Complex64;

use crate::eval::CompiledExpr;
use crate::functions::engineering::complex::{format_complex, parse_complex, ParsedComplex};
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

const VAR_ARGS: usize = 255;

fn eval_complex_arg(
    ctx: &dyn FunctionContext,
    expr: &CompiledExpr,
) -> Result<ParsedComplex, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    match v {
        Value::Error(e) => Err(e),
        Value::Number(n) => {
            if !n.is_finite() {
                return Err(ErrorKind::Num);
            }
            Ok(ParsedComplex {
                value: Complex64::new(n, 0.0),
                suffix: 'i',
            })
        }
        Value::Bool(b) => Ok(ParsedComplex {
            value: Complex64::new(if b { 1.0 } else { 0.0 }, 0.0),
            suffix: 'i',
        }),
        Value::Blank => Ok(ParsedComplex {
            value: Complex64::new(0.0, 0.0),
            suffix: 'i',
        }),
        Value::Text(s) => parse_complex(&s, ctx.number_locale()),
        Value::Entity(v) => parse_complex(&v.display, ctx.number_locale()),
        Value::Record(v) => parse_complex(&v.display, ctx.number_locale()),
        Value::Array(_)
        | Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => Err(ErrorKind::Value),
    }
}

fn checked_complex(z: Complex64) -> Result<Complex64, ErrorKind> {
    if z.re.is_finite() && z.im.is_finite() {
        Ok(z)
    } else {
        Err(ErrorKind::Num)
    }
}

inventory::submit! {
    FunctionSpec {
        name: "COMPLEX",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Text],
        implementation: complex_fn,
    }
}

fn complex_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let re = match eval_scalar_arg(ctx, &args[0]).coerce_to_number_with_ctx(ctx) {
        Ok(v) if v.is_finite() => v,
        Ok(_) => return Value::Error(ErrorKind::Num),
        Err(e) => return Value::Error(e),
    };
    let im = match eval_scalar_arg(ctx, &args[1]).coerce_to_number_with_ctx(ctx) {
        Ok(v) if v.is_finite() => v,
        Ok(_) => return Value::Error(ErrorKind::Num),
        Err(e) => return Value::Error(e),
    };

    let suffix = if args.len() == 3 {
        let raw = match eval_scalar_arg(ctx, &args[2]).coerce_to_string_with_ctx(ctx) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let trimmed = raw.trim();
        if trimmed.is_empty() || trimmed.eq_ignore_ascii_case("i") {
            'i'
        } else if trimmed.eq_ignore_ascii_case("j") {
            'j'
        } else {
            return Value::Error(ErrorKind::Value);
        }
    } else {
        'i'
    };

    match format_complex(Complex64::new(re, im), suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMABS",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: imabs_fn,
    }
}

fn imabs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let z = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v.value,
        Err(e) => return Value::Error(e),
    };
    let out = z.norm();
    if out.is_finite() {
        Value::Number(out)
    } else {
        Value::Error(ErrorKind::Num)
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMAGINARY",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: imaginary_fn,
    }
}

fn imaginary_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let z = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v.value,
        Err(e) => return Value::Error(e),
    };
    Value::Number(z.im)
}

inventory::submit! {
    FunctionSpec {
        name: "IMREAL",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: imreal_fn,
    }
}

fn imreal_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let z = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v.value,
        Err(e) => return Value::Error(e),
    };
    Value::Number(z.re)
}

inventory::submit! {
    FunctionSpec {
        name: "IMARGUMENT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: imargument_fn,
    }
}

fn imargument_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let z = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v.value,
        Err(e) => return Value::Error(e),
    };
    if z.re == 0.0 && z.im == 0.0 {
        return Value::Error(ErrorKind::Div0);
    }
    Value::Number(z.im.atan2(z.re))
}

inventory::submit! {
    FunctionSpec {
        name: "IMCONJUGATE",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: imconjugate_fn,
    }
}

fn imconjugate_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let ParsedComplex { value, suffix } = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match format_complex(value.conj(), suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMSUM",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: imsum_fn,
    }
}

fn imsum_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let first = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut acc = first.value;
    for arg in &args[1..] {
        let z = match eval_complex_arg(ctx, arg) {
            Ok(v) => v.value,
            Err(e) => return Value::Error(e),
        };
        acc += z;
    }
    match format_complex(acc, first.suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMPRODUCT",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: improduct_fn,
    }
}

fn improduct_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let first = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut acc = first.value;
    for arg in &args[1..] {
        let z = match eval_complex_arg(ctx, arg) {
            Ok(v) => v.value,
            Err(e) => return Value::Error(e),
        };
        acc *= z;
    }
    match format_complex(acc, first.suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMSUB",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Text],
        implementation: imsub_fn,
    }
}

fn imsub_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let a = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let b = match eval_complex_arg(ctx, &args[1]) {
        Ok(v) => v.value,
        Err(e) => return Value::Error(e),
    };
    let out = match checked_complex(a.value - b) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match format_complex(out, a.suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMDIV",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Text],
        implementation: imdiv_fn,
    }
}

fn imdiv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let a = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let b = match eval_complex_arg(ctx, &args[1]) {
        Ok(v) => v.value,
        Err(e) => return Value::Error(e),
    };
    if b.re == 0.0 && b.im == 0.0 {
        return Value::Error(ErrorKind::Div0);
    }
    let out = match checked_complex(a.value / b) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match format_complex(out, a.suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMPOWER",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: impower_fn,
    }
}

fn impower_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let a = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let power = match eval_scalar_arg(ctx, &args[1]).coerce_to_number_with_ctx(ctx) {
        Ok(v) if v.is_finite() => v,
        Ok(_) => return Value::Error(ErrorKind::Num),
        Err(e) => return Value::Error(e),
    };

    if a.value.re == 0.0 && a.value.im == 0.0 && power < 0.0 {
        return Value::Error(ErrorKind::Div0);
    }

    // Excel returns exact real results for common integer powers (e.g. IMPOWER("i",2) == "-1").
    // `powf` goes through exp/ln and can introduce small rounding artifacts in the imaginary part,
    // so prefer integer exponentiation when possible.
    let out_raw =
        if power.fract() == 0.0 && power >= (i32::MIN as f64) && power <= (i32::MAX as f64) {
            a.value.powi(power as i32)
        } else {
            a.value.powf(power)
        };

    let out = match checked_complex(out_raw) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match format_complex(out, a.suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMSQRT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: imsqrt_fn,
    }
}

fn imsqrt_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let a = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let out = match checked_complex(a.value.sqrt()) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match format_complex(out, a.suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMLN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: imln_fn,
    }
}

fn imln_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let a = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let out = match checked_complex(a.value.ln()) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match format_complex(out, a.suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMLOG2",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: imlog2_fn,
    }
}

fn imlog2_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let a = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let out = match checked_complex(a.value.ln() / std::f64::consts::LN_2) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match format_complex(out, a.suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMLOG10",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: imlog10_fn,
    }
}

fn imlog10_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let a = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let out = match checked_complex(a.value.ln() / std::f64::consts::LN_10) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match format_complex(out, a.suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMSIN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: imsin_fn,
    }
}

fn imsin_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let a = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let out = match checked_complex(a.value.sin()) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match format_complex(out, a.suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMCOS",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: imcos_fn,
    }
}

fn imcos_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let a = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let out = match checked_complex(a.value.cos()) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match format_complex(out, a.suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IMEXP",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: imexp_fn,
    }
}

fn imexp_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let a = match eval_complex_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let out = match checked_complex(a.value.exp()) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match format_complex(out, a.suffix, ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}
