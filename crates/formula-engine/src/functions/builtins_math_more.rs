use std::collections::HashSet;

use crate::error::ExcelError;
use crate::eval::CompiledExpr;
use crate::functions::array_lift;
use crate::functions::{ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

fn excel_error_kind(err: ExcelError) -> ErrorKind {
    match err {
        ExcelError::Div0 => ErrorKind::Div0,
        ExcelError::Value => ErrorKind::Value,
        ExcelError::Num => ErrorKind::Num,
    }
}

fn lift1_number(
    ctx: &dyn FunctionContext,
    expr: &CompiledExpr,
    f: impl Fn(f64) -> Result<f64, ExcelError>,
) -> Value {
    let value = array_lift::eval_arg(ctx, expr);
    array_lift::lift1(value, |v| {
        let n = v.coerce_to_number_with_ctx(ctx)?;
        match f(n) {
            Ok(out) => Ok(Value::Number(out)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

fn lift2_number(
    ctx: &dyn FunctionContext,
    a: &CompiledExpr,
    b: &CompiledExpr,
    f: impl Fn(f64, f64) -> Result<f64, ExcelError>,
) -> Value {
    let a = array_lift::eval_arg(ctx, a);
    let b = array_lift::eval_arg(ctx, b);
    array_lift::lift2(a, b, |a, b| {
        let a = a.coerce_to_number_with_ctx(ctx)?;
        let b = b.coerce_to_number_with_ctx(ctx)?;
        match f(a, b) {
            Ok(out) => Ok(Value::Number(out)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

fn lift1_bool(
    ctx: &dyn FunctionContext,
    expr: &CompiledExpr,
    f: impl Fn(f64) -> Result<bool, ExcelError>,
) -> Value {
    let value = array_lift::eval_arg(ctx, expr);
    array_lift::lift1(value, |v| {
        let n = v.coerce_to_number_with_ctx(ctx)?;
        match f(n) {
            Ok(out) => Ok(Value::Bool(out)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

// ----------------------------------------------------------------------
// Degrees/radians + trig helpers
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "RADIANS",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: radians_fn,
    }
}

fn radians_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::radians)
}

inventory::submit! {
    FunctionSpec {
        name: "DEGREES",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: degrees_fn,
    }
}

fn degrees_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::degrees)
}

inventory::submit! {
    FunctionSpec {
        name: "COT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: cot_fn,
    }
}

fn cot_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::cot)
}

inventory::submit! {
    FunctionSpec {
        name: "CSC",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: csc_fn,
    }
}

fn csc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::csc)
}

inventory::submit! {
    FunctionSpec {
        name: "SEC",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: sec_fn,
    }
}

fn sec_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::sec)
}

inventory::submit! {
    FunctionSpec {
        name: "ACOT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: acot_fn,
    }
}

fn acot_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::acot)
}

// ----------------------------------------------------------------------
// Hyperbolic trig
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "SINH",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: sinh_fn,
    }
}

fn sinh_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::sinh)
}

inventory::submit! {
    FunctionSpec {
        name: "COSH",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: cosh_fn,
    }
}

fn cosh_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::cosh)
}

inventory::submit! {
    FunctionSpec {
        name: "TANH",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: tanh_fn,
    }
}

fn tanh_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::tanh)
}

inventory::submit! {
    FunctionSpec {
        name: "ASINH",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: asinh_fn,
    }
}

fn asinh_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::asinh)
}

inventory::submit! {
    FunctionSpec {
        name: "ACOSH",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: acosh_fn,
    }
}

fn acosh_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::acosh)
}

inventory::submit! {
    FunctionSpec {
        name: "ATANH",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: atanh_fn,
    }
}

fn atanh_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::atanh)
}

inventory::submit! {
    FunctionSpec {
        name: "COTH",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: coth_fn,
    }
}

fn coth_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::coth)
}

inventory::submit! {
    FunctionSpec {
        name: "CSCH",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: csch_fn,
    }
}

fn csch_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::csch)
}

inventory::submit! {
    FunctionSpec {
        name: "SECH",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: sech_fn,
    }
}

fn sech_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::sech)
}

inventory::submit! {
    FunctionSpec {
        name: "ACOTH",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: acoth_fn,
    }
}

fn acoth_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::acoth)
}

// ----------------------------------------------------------------------
// Combinatorics + integer helpers
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "FACT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: fact_fn,
    }
}

fn fact_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::fact)
}

inventory::submit! {
    FunctionSpec {
        name: "FACTDOUBLE",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: factdouble_fn,
    }
}

fn factdouble_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::factdouble)
}

inventory::submit! {
    FunctionSpec {
        name: "COMBIN",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: combin_fn,
    }
}

fn combin_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift2_number(ctx, &args[0], &args[1], crate::functions::math::combin)
}

inventory::submit! {
    FunctionSpec {
        name: "COMBINA",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: combina_fn,
    }
}

fn combina_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift2_number(ctx, &args[0], &args[1], crate::functions::math::combina)
}

inventory::submit! {
    FunctionSpec {
        name: "PERMUT",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: permut_fn,
    }
}

fn permut_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift2_number(ctx, &args[0], &args[1], crate::functions::math::permut)
}

inventory::submit! {
    FunctionSpec {
        name: "PERMUTATIONA",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: permutationa_fn,
    }
}

fn permutationa_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift2_number(
        ctx,
        &args[0],
        &args[1],
        crate::functions::math::permutationa,
    )
}

inventory::submit! {
    FunctionSpec {
        name: "MROUND",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: mround_fn,
    }
}

fn mround_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift2_number(ctx, &args[0], &args[1], crate::functions::math::mround)
}

inventory::submit! {
    FunctionSpec {
        name: "EVEN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: even_fn,
    }
}

fn even_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::even)
}

inventory::submit! {
    FunctionSpec {
        name: "ODD",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: odd_fn,
    }
}

fn odd_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::odd)
}

inventory::submit! {
    FunctionSpec {
        name: "ISEVEN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Number],
        implementation: iseven_fn,
    }
}

fn iseven_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_bool(ctx, &args[0], crate::functions::math::iseven)
}

inventory::submit! {
    FunctionSpec {
        name: "ISODD",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Number],
        implementation: isodd_fn,
    }
}

fn isodd_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_bool(ctx, &args[0], crate::functions::math::isodd)
}

inventory::submit! {
    FunctionSpec {
        name: "QUOTIENT",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: quotient_fn,
    }
}

fn quotient_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift2_number(ctx, &args[0], &args[1], crate::functions::math::quotient)
}

inventory::submit! {
    FunctionSpec {
        name: "SQRTPI",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: sqrtpi_fn,
    }
}

fn sqrtpi_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    lift1_number(ctx, &args[0], crate::functions::math::sqrtpi)
}

inventory::submit! {
    FunctionSpec {
        name: "DELTA",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: delta_fn,
    }
}

fn delta_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number1 = array_lift::eval_arg(ctx, &args[0]);
    let number2 = match args.get(1) {
        None | Some(CompiledExpr::Blank) => Value::Number(0.0),
        Some(expr) => array_lift::eval_arg(ctx, expr),
    };
    array_lift::lift2(number1, number2, |a, b| {
        let a = a.coerce_to_number_with_ctx(ctx)?;
        let b = b.coerce_to_number_with_ctx(ctx)?;
        match crate::functions::math::delta(a, b) {
            Ok(out) => Ok(Value::Number(out)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "GESTEP",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: gestep_fn,
    }
}

fn gestep_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    let step = match args.get(1) {
        None | Some(CompiledExpr::Blank) => Value::Number(0.0),
        Some(expr) => array_lift::eval_arg(ctx, expr),
    };
    array_lift::lift2(number, step, |a, b| {
        let a = a.coerce_to_number_with_ctx(ctx)?;
        let b = b.coerce_to_number_with_ctx(ctx)?;
        match crate::functions::math::gestep(a, b) {
            Ok(out) => Ok(Value::Number(out)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

// ----------------------------------------------------------------------
// Aggregates over argument lists
// ----------------------------------------------------------------------

fn push_numbers_from_scalar(
    ctx: &dyn FunctionContext,
    out: &mut Vec<f64>,
    value: Value,
) -> Result<(), ErrorKind> {
    match value {
        Value::Error(e) => Err(e),
        Value::Number(n) => {
            out.push(n);
            Ok(())
        }
        Value::Bool(b) => {
            out.push(if b { 1.0 } else { 0.0 });
            Ok(())
        }
        Value::Blank => Ok(()),
        Value::Text(s) => {
            out.push(Value::Text(s).coerce_to_number_with_ctx(ctx)?);
            Ok(())
        }
        Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
        Value::Array(arr) => {
            for v in arr.iter() {
                match v {
                    Value::Error(e) => return Err(*e),
                    Value::Number(n) => out.push(*n),
                    Value::Lambda(_) => return Err(ErrorKind::Value),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Blank
                    | Value::Array(_)
                    | Value::Spill { .. }
                    | Value::Reference(_)
                    | Value::ReferenceUnion(_) => {}
                }
            }
            Ok(())
        }
        Value::Reference(_) | Value::ReferenceUnion(_) | Value::Lambda(_) | Value::Spill { .. } => {
            Err(ErrorKind::Value)
        }
    }
}

fn push_numbers_from_reference(
    ctx: &dyn FunctionContext,
    out: &mut Vec<f64>,
    reference: crate::functions::Reference,
) -> Result<(), ErrorKind> {
    for addr in ctx.iter_reference_cells(&reference) {
        let v = ctx.get_cell_value(&reference.sheet_id, addr);
        match v {
            Value::Error(e) => return Err(e),
            Value::Number(n) => out.push(n),
            Value::Lambda(_) => return Err(ErrorKind::Value),
            Value::Bool(_)
            | Value::Text(_)
            | Value::Entity(_)
            | Value::Record(_)
            | Value::Blank
            | Value::Array(_)
            | Value::Spill { .. }
            | Value::Reference(_)
            | Value::ReferenceUnion(_) => {}
        }
    }
    Ok(())
}

fn push_numbers_from_reference_union(
    ctx: &dyn FunctionContext,
    out: &mut Vec<f64>,
    ranges: Vec<crate::functions::Reference>,
) -> Result<(), ErrorKind> {
    let mut seen = HashSet::new();
    for reference in ranges {
        for addr in ctx.iter_reference_cells(&reference) {
            if !seen.insert((reference.sheet_id.clone(), addr)) {
                continue;
            }
            let v = ctx.get_cell_value(&reference.sheet_id, addr);
            match v {
                Value::Error(e) => return Err(e),
                Value::Number(n) => out.push(n),
                Value::Lambda(_) => return Err(ErrorKind::Value),
                Value::Bool(_)
                | Value::Text(_)
                | Value::Entity(_)
                | Value::Record(_)
                | Value::Blank
                | Value::Array(_)
                | Value::Spill { .. }
                | Value::Reference(_)
                | Value::ReferenceUnion(_) => {}
            }
        }
    }
    Ok(())
}

fn push_numbers_from_arg(
    ctx: &dyn FunctionContext,
    out: &mut Vec<f64>,
    arg: ArgValue,
) -> Result<(), ErrorKind> {
    match arg {
        ArgValue::Scalar(v) => push_numbers_from_scalar(ctx, out, v),
        ArgValue::Reference(r) => push_numbers_from_reference(ctx, out, r),
        ArgValue::ReferenceUnion(ranges) => push_numbers_from_reference_union(ctx, out, ranges),
    }
}

fn collect_numbers(
    ctx: &dyn FunctionContext,
    args: &[CompiledExpr],
) -> Result<Vec<f64>, ErrorKind> {
    let mut out = Vec::new();
    for expr in args {
        push_numbers_from_arg(ctx, &mut out, ctx.eval_arg(expr))?;
    }
    Ok(out)
}

inventory::submit! {
    FunctionSpec {
        name: "GCD",
        min_args: 1,
        max_args: 255,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: gcd_fn,
    }
}

fn gcd_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::math::gcd(&values) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "LCM",
        min_args: 1,
        max_args: 255,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: lcm_fn,
    }
}

fn lcm_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::math::lcm(&values) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "MULTINOMIAL",
        min_args: 1,
        max_args: 255,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: multinomial_fn,
    }
}

fn multinomial_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::math::multinomial(&values) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

// ----------------------------------------------------------------------
// SERIESSUM + SUMX* helpers
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "SERIESSUM",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Any],
        implementation: seriessum_fn,
    }
}

fn collect_coefficients(ctx: &dyn FunctionContext, arg: ArgValue) -> Result<Vec<f64>, ErrorKind> {
    fn coerce_coeff(ctx: &dyn FunctionContext, value: Value) -> Result<f64, ErrorKind> {
        let n = value.coerce_to_number_with_ctx(ctx)?;
        Ok(n)
    }

    match arg {
        ArgValue::Scalar(v) => match v {
            Value::Array(arr) => {
                let mut out = Vec::with_capacity(arr.values.len());
                for v in arr.values {
                    out.push(coerce_coeff(ctx, v)?);
                }
                Ok(out)
            }
            other => Ok(vec![coerce_coeff(ctx, other)?]),
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            ctx.record_reference(&r);
            let mut out = Vec::with_capacity(r.size() as usize);
            for addr in r.iter_cells() {
                let v = ctx.get_cell_value(&r.sheet_id, addr);
                out.push(coerce_coeff(ctx, v)?);
            }
            Ok(out)
        }
        ArgValue::ReferenceUnion(_) => Err(ErrorKind::Value),
    }
}

fn seriessum_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let n = array_lift::eval_arg(ctx, &args[1]);
    let m = array_lift::eval_arg(ctx, &args[2]);
    let coeffs = match collect_coefficients(ctx, ctx.eval_arg(&args[3])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    array_lift::lift3(x, n, m, |x, n, m| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let n = n.coerce_to_number_with_ctx(ctx)?;
        let m = m.coerce_to_number_with_ctx(ctx)?;
        match crate::functions::math::seriessum(x, n, m, &coeffs) {
            Ok(out) => Ok(Value::Number(out)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

fn arg_to_numeric_sequence(
    ctx: &dyn FunctionContext,
    arg: ArgValue,
) -> Result<Vec<Option<f64>>, ErrorKind> {
    match arg {
        ArgValue::Scalar(v) => match v {
            Value::Error(e) => Err(e),
            Value::Number(n) => Ok(vec![Some(n)]),
            Value::Bool(b) => Ok(vec![Some(if b { 1.0 } else { 0.0 })]),
            Value::Blank => Ok(vec![None]),
            Value::Text(s) => Ok(vec![Some(Value::Text(s).coerce_to_number_with_ctx(ctx)?)]),
            Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
            Value::Array(arr) => {
                let mut out = Vec::with_capacity(arr.values.len());
                for v in arr.iter() {
                    match v {
                        Value::Error(e) => return Err(*e),
                        Value::Number(n) => out.push(Some(*n)),
                        Value::Lambda(_) => return Err(ErrorKind::Value),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Entity(_)
                        | Value::Record(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Spill { .. }
                        | Value::Reference(_)
                        | Value::ReferenceUnion(_) => out.push(None),
                    }
                }
                Ok(out)
            }
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            ctx.record_reference(&r);
            let rows = (r.end.row - r.start.row + 1) as usize;
            let cols = (r.end.col - r.start.col + 1) as usize;
            let mut out = Vec::with_capacity(rows.saturating_mul(cols));
            for addr in r.iter_cells() {
                let v = ctx.get_cell_value(&r.sheet_id, addr);
                match v {
                    Value::Error(e) => return Err(e),
                    Value::Number(n) => out.push(Some(n)),
                    Value::Lambda(_) => return Err(ErrorKind::Value),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Blank
                    | Value::Array(_)
                    | Value::Spill { .. }
                    | Value::Reference(_)
                    | Value::ReferenceUnion(_) => out.push(None),
                }
            }
            Ok(out)
        }
        ArgValue::ReferenceUnion(ranges) => {
            let mut seen = HashSet::new();
            let mut out = Vec::new();
            for r in ranges {
                let r = r.normalized();
                ctx.record_reference(&r);
                let rows = (r.end.row - r.start.row + 1) as usize;
                let cols = (r.end.col - r.start.col + 1) as usize;
                out.reserve(rows.saturating_mul(cols));
                for addr in r.iter_cells() {
                    if !seen.insert((r.sheet_id.clone(), addr)) {
                        continue;
                    }
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Err(e),
                        Value::Number(n) => out.push(Some(n)),
                        Value::Lambda(_) => return Err(ErrorKind::Value),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Entity(_)
                        | Value::Record(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Spill { .. }
                        | Value::Reference(_)
                        | Value::ReferenceUnion(_) => out.push(None),
                    }
                }
            }
            Ok(out)
        }
    }
}

fn collect_numeric_pairs(
    ctx: &dyn FunctionContext,
    left_expr: &CompiledExpr,
    right_expr: &CompiledExpr,
) -> Result<(Vec<f64>, Vec<f64>), ErrorKind> {
    let left = arg_to_numeric_sequence(ctx, ctx.eval_arg(left_expr))?;
    let right = arg_to_numeric_sequence(ctx, ctx.eval_arg(right_expr))?;
    if left.len() != right.len() {
        return Err(ErrorKind::NA);
    }

    let mut xs = Vec::new();
    let mut ys = Vec::new();
    for (lx, ry) in left.into_iter().zip(right.into_iter()) {
        let (Some(x), Some(y)) = (lx, ry) else {
            continue;
        };
        xs.push(x);
        ys.push(y);
    }
    Ok((xs, ys))
}

inventory::submit! {
    FunctionSpec {
        name: "SUMXMY2",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: sumxmy2_fn,
    }
}

fn sumxmy2_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let (xs, ys) = match collect_numeric_pairs(ctx, &args[0], &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::math::sumxmy2(&xs, &ys) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "SUMX2MY2",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: sumx2my2_fn,
    }
}

fn sumx2my2_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let (xs, ys) = match collect_numeric_pairs(ctx, &args[0], &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::math::sumx2my2(&xs, &ys) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "SUMX2PY2",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: sumx2py2_fn,
    }
}

fn sumx2py2_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let (xs, ys) = match collect_numeric_pairs(ctx, &args[0], &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::math::sumx2py2(&xs, &ys) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
