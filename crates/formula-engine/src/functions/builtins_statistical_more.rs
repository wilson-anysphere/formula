use std::collections::HashSet;

use crate::eval::CompiledExpr;
use crate::functions::{ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{Array, ErrorKind, Value};

fn push_frequency_data_numbers(
    ctx: &dyn FunctionContext,
    out: &mut Vec<f64>,
    arg: ArgValue,
) -> Result<(), ErrorKind> {
    match arg {
        ArgValue::Scalar(v) => push_frequency_data_numbers_from_value(ctx, out, v),
        ArgValue::Reference(reference) => {
            for addr in ctx.iter_reference_cells(&reference) {
                let v = ctx.get_cell_value(&reference.sheet_id, addr);
                push_frequency_data_numbers_from_value(ctx, out, v)?;
            }
            Ok(())
        }
        ArgValue::ReferenceUnion(ranges) => {
            let mut seen = HashSet::new();
            for reference in ranges {
                for addr in ctx.iter_reference_cells(&reference) {
                    if !seen.insert((reference.sheet_id.clone(), addr)) {
                        continue;
                    }
                    let v = ctx.get_cell_value(&reference.sheet_id, addr);
                    push_frequency_data_numbers_from_value(ctx, out, v)?;
                }
            }
            Ok(())
        }
    }
}

fn push_frequency_data_numbers_from_value(
    _ctx: &dyn FunctionContext,
    out: &mut Vec<f64>,
    value: Value,
) -> Result<(), ErrorKind> {
    match value {
        Value::Error(e) => Err(e),
        Value::Number(n) => {
            if !n.is_finite() {
                return Err(ErrorKind::Num);
            }
            out.push(n);
            Ok(())
        }
        Value::Array(arr) => {
            for v in arr.iter() {
                match v {
                    Value::Error(e) => return Err(*e),
                    Value::Number(n) => {
                        if !n.is_finite() {
                            return Err(ErrorKind::Num);
                        }
                        out.push(*n);
                    }
                    Value::Lambda(_) => return Err(ErrorKind::Value),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Blank
                    | Value::Array(_)
                    | Value::Spill { .. }
                    | Value::Reference(_)
                    | Value::ReferenceUnion(_) => {
                        // Excel ignores non-numeric values in the data array.
                    }
                }
            }
            Ok(())
        }
        Value::Lambda(_) => Err(ErrorKind::Value),
        Value::Bool(_) | Value::Text(_) | Value::Entity(_) | Value::Record(_) | Value::Blank => {
            Ok(())
        }
        Value::Reference(_) | Value::ReferenceUnion(_) | Value::Spill { .. } => {
            Err(ErrorKind::Value)
        }
    }
}

fn push_frequency_bin_numbers(
    ctx: &dyn FunctionContext,
    out: &mut Vec<f64>,
    arg: ArgValue,
) -> Result<(), ErrorKind> {
    match arg {
        ArgValue::Scalar(v) => push_frequency_bin_numbers_from_value(ctx, out, v),
        ArgValue::Reference(reference) => {
            for addr in ctx.iter_reference_cells(&reference) {
                let v = ctx.get_cell_value(&reference.sheet_id, addr);
                push_frequency_bin_numbers_from_value(ctx, out, v)?;
            }
            Ok(())
        }
        ArgValue::ReferenceUnion(ranges) => {
            let mut seen = HashSet::new();
            for reference in ranges {
                for addr in ctx.iter_reference_cells(&reference) {
                    if !seen.insert((reference.sheet_id.clone(), addr)) {
                        continue;
                    }
                    let v = ctx.get_cell_value(&reference.sheet_id, addr);
                    push_frequency_bin_numbers_from_value(ctx, out, v)?;
                }
            }
            Ok(())
        }
    }
}

fn push_frequency_bin_numbers_from_value(
    ctx: &dyn FunctionContext,
    out: &mut Vec<f64>,
    value: Value,
) -> Result<(), ErrorKind> {
    match value {
        Value::Error(e) => Err(e),
        Value::Lambda(_) => Err(ErrorKind::Value),
        Value::Array(arr) => {
            for v in arr.iter() {
                let n = match v {
                    Value::Error(e) => return Err(*e),
                    Value::Lambda(_) => return Err(ErrorKind::Value),
                    other => other.coerce_to_number_with_ctx(ctx)?,
                };
                if !n.is_finite() {
                    return Err(ErrorKind::Num);
                }
                out.push(n);
            }
            Ok(())
        }
        other => {
            let n = match other {
                Value::Reference(_) | Value::ReferenceUnion(_) | Value::Spill { .. } => {
                    return Err(ErrorKind::Value);
                }
                other => other.coerce_to_number_with_ctx(ctx)?,
            };
            if !n.is_finite() {
                return Err(ErrorKind::Num);
            }
            out.push(n);
            Ok(())
        }
    }
}

inventory::submit! {
    FunctionSpec {
        name: "FREQUENCY",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: frequency_fn,
    }
}

fn frequency_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut data = Vec::new();
    if let Err(e) = push_frequency_data_numbers(ctx, &mut data, ctx.eval_arg(&args[0])) {
        return Value::Error(e);
    }

    let mut bins = Vec::new();
    if let Err(e) = push_frequency_bin_numbers(ctx, &mut bins, ctx.eval_arg(&args[1])) {
        return Value::Error(e);
    }

    let counts = match crate::functions::statistical::frequency(&data, &bins) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let mut values = Vec::new();
    if values.try_reserve_exact(counts.len()).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for count in counts {
        let n = count as f64;
        if !n.is_finite() {
            return Value::Error(ErrorKind::Num);
        }
        values.push(Value::Number(n));
    }

    Value::Array(Array::new(values.len(), 1, values))
}
