use std::collections::{HashMap, HashSet};

use crate::eval::CompiledExpr;
use crate::functions::array_lift;
use crate::functions::{ArgValue, ArraySupport, FunctionContext, FunctionSpec, Reference};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{Array, ErrorKind, Value};

const VAR_ARGS: usize = 255;

inventory::submit! {
    FunctionSpec {
        name: "CHOOSE",
        min_args: 2,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Number, ValueType::Any],
        implementation: choose_fn,
    }
}

fn choose_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let choices = &args[1..];
    let index_value = array_lift::eval_arg(ctx, &args[0]);

    match index_value {
        Value::Error(e) => Value::Error(e),
        Value::Array(arr) => choose_array(ctx, &arr, choices),
        other => choose_scalar(ctx, other, choices),
    }
}

fn choose_scalar(ctx: &dyn FunctionContext, index_value: Value, choices: &[CompiledExpr]) -> Value {
    if let Value::Error(e) = index_value {
        return Value::Error(e);
    }

    let idx = match index_value.coerce_to_i64() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    if idx < 1 {
        return Value::Error(ErrorKind::Value);
    }

    let choice_idx = match usize::try_from(idx - 1) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Value),
    };
    let Some(expr) = choices.get(choice_idx) else {
        return Value::Error(ErrorKind::Value);
    };

    arg_value_to_value(ctx.eval_arg(expr))
}

fn choose_array(ctx: &dyn FunctionContext, indices: &Array, choices: &[CompiledExpr]) -> Value {
    if choices.is_empty() {
        return Value::Error(ErrorKind::Value);
    }

    let max_index = i64::try_from(choices.len()).unwrap_or(i64::MAX);
    let mut normalized = Vec::with_capacity(indices.values.len());
    let mut used_indices: HashSet<usize> = HashSet::new();

    for v in &indices.values {
        let res = match v {
            Value::Error(e) => Err(*e),
            other => match other.coerce_to_i64() {
                Ok(idx) if idx >= 1 && idx <= max_index => {
                    let zero_based = (idx - 1) as usize;
                    used_indices.insert(zero_based);
                    Ok(zero_based)
                }
                Ok(_) => Err(ErrorKind::Value),
                Err(e) => Err(e),
            },
        };
        normalized.push(res);
    }

    let mut evaluated: HashMap<usize, ArgValue> = HashMap::with_capacity(used_indices.len());
    for idx in &used_indices {
        if let Some(expr) = choices.get(*idx) {
            evaluated.insert(*idx, ctx.eval_arg(expr));
        }
    }

    // When `CHOOSE` receives an array of indices and the selected values are multi-cell references,
    // Excel produces a multi-area reference union (commonly used with `SUM(CHOOSE({1,2},range1,range2))`).
    // Single-cell references behave like scalar values and should not force union output.
    let mut union_candidate = false;
    let mut all_refs = true;
    for arg in evaluated.values() {
        match arg {
            ArgValue::Reference(r) => {
                if !r.is_single_cell() {
                    union_candidate = true;
                }
            }
            ArgValue::ReferenceUnion(_) => union_candidate = true,
            _ => {
                all_refs = false;
                break;
            }
        }
    }
    union_candidate &= all_refs;

    if union_candidate {
        let mut ranges: Vec<Reference> = Vec::new();
        for idx in normalized {
            let idx = match idx {
                Ok(v) => v,
                Err(e) => return Value::Error(e),
            };
            let Some(arg) = evaluated.get(&idx) else {
                return Value::Error(ErrorKind::Value);
            };
            match arg {
                ArgValue::Reference(r) => ranges.push(*r),
                ArgValue::ReferenceUnion(rs) => ranges.extend(rs.iter().copied()),
                ArgValue::Scalar(Value::Error(e)) => return Value::Error(*e),
                _ => return Value::Error(ErrorKind::Value),
            }
        }

        return Value::ReferenceUnion(ranges);
    }

    let mut out_values: Vec<Value> = Vec::with_capacity(indices.values.len());
    for idx in normalized {
        match idx {
            Err(e) => out_values.push(Value::Error(e)),
            Ok(idx) => {
                let Some(arg) = evaluated.get(&idx) else {
                    out_values.push(Value::Error(ErrorKind::Value));
                    continue;
                };
                out_values.push(arg_value_to_scalar(ctx, arg));
            }
        }
    }

    Value::Array(Array::new(indices.rows, indices.cols, out_values))
}

fn arg_value_to_scalar(ctx: &dyn FunctionContext, arg: &ArgValue) -> Value {
    match arg.clone() {
        ArgValue::Scalar(v) => scalarize_value(v),
        ArgValue::Reference(r) => {
            if r.is_single_cell() {
                ctx.get_cell_value(r.sheet_id, r.start)
            } else {
                Value::Error(ErrorKind::Value)
            }
        }
        ArgValue::ReferenceUnion(_) => Value::Error(ErrorKind::Value),
    }
}

fn scalarize_value(value: Value) -> Value {
    match value {
        Value::Array(arr) => {
            if arr.rows == 1 && arr.cols == 1 {
                arr.top_left()
            } else {
                Value::Error(ErrorKind::Value)
            }
        }
        Value::Lambda(_) | Value::Spill { .. } | Value::Reference(_) | Value::ReferenceUnion(_) => {
            Value::Error(ErrorKind::Value)
        }
        other => other,
    }
}

fn arg_value_to_value(arg: ArgValue) -> Value {
    match arg {
        ArgValue::Scalar(v) => v,
        ArgValue::Reference(r) => Value::Reference(r),
        ArgValue::ReferenceUnion(ranges) => Value::ReferenceUnion(ranges),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IFS",
        min_args: 2,
        // Excel supports up to 127 condition/value pairs.
        max_args: 254,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: ifs_fn,
    }
}

fn ifs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.len() % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }
    ifs_pairs(ctx, args)
}

fn ifs_pairs(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.len() < 2 {
        return Value::Error(ErrorKind::NA);
    }

    let cond_expr = &args[0];
    let value_expr = &args[1];
    let remaining = &args[2..];

    let cond_val = array_lift::eval_arg(ctx, cond_expr);
    match cond_val {
        Value::Array(ref arr) => {
            let mut needs_true = false;
            let mut needs_false = false;
            for el in arr.iter() {
                match el.coerce_to_bool() {
                    Ok(true) => needs_true = true,
                    Ok(false) => needs_false = true,
                    Err(_) => {}
                }
                if needs_true && needs_false {
                    break;
                }
            }

            if needs_true && needs_false {
                let true_val = array_lift::eval_arg(ctx, value_expr);
                let false_val = ifs_pairs(ctx, remaining);
                return array_lift::lift3(cond_val, true_val, false_val, |cond, t, f| {
                    if cond.coerce_to_bool()? {
                        Ok(t.clone())
                    } else {
                        Ok(f.clone())
                    }
                });
            }

            if needs_true {
                let true_val = array_lift::eval_arg(ctx, value_expr);
                return array_lift::lift2(cond_val, true_val, |cond, t| {
                    if cond.coerce_to_bool()? {
                        Ok(t.clone())
                    } else {
                        Ok(Value::Error(ErrorKind::NA))
                    }
                });
            }

            if needs_false {
                let false_val = ifs_pairs(ctx, remaining);
                return array_lift::lift2(cond_val, false_val, |cond, f| {
                    if cond.coerce_to_bool()? {
                        Ok(Value::Error(ErrorKind::NA))
                    } else {
                        Ok(f.clone())
                    }
                });
            }

            // The condition array contains only errors/invalid values, so avoid forcing evaluation of
            // any branch expressions and simply map the coercion errors.
            array_lift::lift1(cond_val, |cond| {
                let _ = cond.coerce_to_bool()?;
                Ok(Value::Error(ErrorKind::NA))
            })
        }
        other => {
            if let Value::Error(e) = other {
                return Value::Error(e);
            }

            let cond = match other.coerce_to_bool() {
                Ok(b) => b,
                Err(e) => return Value::Error(e),
            };

            if cond {
                array_lift::eval_arg(ctx, value_expr)
            } else {
                ifs_pairs(ctx, remaining)
            }
        }
    }
}
