use std::collections::HashSet;

use crate::eval::CompiledExpr;
use crate::functions::{ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

const VAR_ARGS: usize = 255;

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
            let n = Value::Text(s).coerce_to_number_with_ctx(ctx)?;
            out.push(n);
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
        name: "KURT",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: kurt_fn,
    }
}

fn kurt_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    match crate::functions::statistical::kurt(&values) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "SKEW",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: skew_fn,
    }
}

fn skew_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    match crate::functions::statistical::skew(&values) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "SKEW.P",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: skew_p_fn,
    }
}

fn skew_p_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    match crate::functions::statistical::skew_p(&values) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
