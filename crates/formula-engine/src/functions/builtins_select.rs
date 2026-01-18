use std::collections::HashMap;

use crate::eval::CompiledExpr;
use crate::functions::array_lift;
use crate::functions::{ArgValue, ArraySupport, FunctionContext, FunctionSpec, Reference};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{cmp_case_insensitive, try_vec_with_capacity, Array, ErrorKind, Value};
use std::cmp::Ordering;

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
        Value::Array(arr) => {
            if arr.rows == 1 && arr.cols == 1 {
                choose_scalar(ctx, arr.top_left(), choices)
            } else {
                choose_array(ctx, &arr, choices)
            }
        }
        other => choose_scalar(ctx, other, choices),
    }
}

fn choose_scalar(ctx: &dyn FunctionContext, index_value: Value, choices: &[CompiledExpr]) -> Value {
    if let Value::Error(e) = index_value {
        return Value::Error(e);
    }

    let idx = match index_value.coerce_to_i64_with_ctx(ctx) {
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
    let mut normalized: Vec<Result<usize, ErrorKind>> = match try_vec_with_capacity(indices.values.len())
    {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut used: Vec<bool> = match try_vec_with_capacity(choices.len()) {
        Ok(mut v) => {
            v.resize(choices.len(), false);
            v
        }
        Err(e) => return Value::Error(e),
    };

    for v in &indices.values {
        let res = match v {
            Value::Error(e) => Err(*e),
            other => match other.coerce_to_i64_with_ctx(ctx) {
                Ok(idx) if idx >= 1 && idx <= max_index => {
                    let zero_based = (idx - 1) as usize;
                    if let Some(slot) = used.get_mut(zero_based) {
                        *slot = true;
                    }
                    Ok(zero_based)
                }
                Ok(_) => Err(ErrorKind::Value),
                Err(e) => Err(e),
            },
        };
        normalized.push(res);
    }

    let used_len = used.iter().filter(|&&b| b).count();
    let mut evaluated: HashMap<usize, ArgValue> = HashMap::new();
    if evaluated.try_reserve(used_len).is_err() {
        debug_assert!(false, "CHOOSE allocation failed (evaluated args)");
        return Value::Error(ErrorKind::Num);
    }
    for (idx, is_used) in used.iter().copied().enumerate() {
        if !is_used {
            continue;
        }
        if let Some(expr) = choices.get(idx) {
            evaluated.insert(idx, ctx.eval_arg(expr));
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
        let mut cap = 0usize;
        for arg in evaluated.values() {
            cap = cap.saturating_add(match arg {
                ArgValue::Reference(_) => 1,
                ArgValue::ReferenceUnion(rs) => rs.len(),
                _ => 0,
            });
        }
        let mut ranges: Vec<Reference> = match try_vec_with_capacity(cap) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        for idx in normalized {
            let idx = match idx {
                Ok(v) => v,
                Err(e) => return Value::Error(e),
            };
            let Some(arg) = evaluated.get(&idx) else {
                return Value::Error(ErrorKind::Value);
            };
            match arg {
                ArgValue::Reference(r) => ranges.push(r.clone()),
                ArgValue::ReferenceUnion(rs) => ranges.extend(rs.iter().cloned()),
                ArgValue::Scalar(Value::Error(e)) => return Value::Error(*e),
                _ => return Value::Error(ErrorKind::Value),
            }
        }

        return Value::ReferenceUnion(ranges);
    }

    let shape = array_lift::Shape {
        rows: indices.rows,
        cols: indices.cols,
    };

    let mut evaluated_values: HashMap<usize, Value> = HashMap::new();
    if evaluated_values.try_reserve(evaluated.len()).is_err() {
        debug_assert!(false, "CHOOSE allocation failed (evaluated values)");
        return Value::Error(ErrorKind::Num);
    }
    for (idx, arg) in evaluated {
        let value = choose_value_from_arg(ctx, arg);
        evaluated_values.insert(idx, value);
    }

    let mut out_values: Vec<Value> = match try_vec_with_capacity(indices.values.len()) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for (pos, idx) in normalized.into_iter().enumerate() {
        match idx {
            Err(e) => out_values.push(Value::Error(e)),
            Ok(idx) => {
                let Some(value) = evaluated_values.get(&idx) else {
                    out_values.push(Value::Error(ErrorKind::Value));
                    continue;
                };
                out_values.push(match value {
                    Value::Array(arr) => {
                        if arr.rows == 1 && arr.cols == 1 {
                            arr.values[0].clone()
                        } else if arr.rows == shape.rows && arr.cols == shape.cols {
                            arr.values
                                .get(pos)
                                .cloned()
                                .unwrap_or(Value::Error(ErrorKind::Value))
                        } else {
                            Value::Error(ErrorKind::Value)
                        }
                    }
                    other => other.clone(),
                });
            }
        }
    }

    Value::Array(Array::new(shape.rows, shape.cols, out_values))
}

fn choose_value_from_arg(ctx: &dyn FunctionContext, arg: ArgValue) -> Value {
    match arg {
        ArgValue::Scalar(v) => match v {
            Value::Lambda(_)
            | Value::Spill { .. }
            | Value::Reference(_)
            | Value::ReferenceUnion(_) => Value::Error(ErrorKind::Value),
            other => other,
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            ctx.record_reference(&r);
            if r.is_single_cell() {
                ctx.get_cell_value(&r.sheet_id, r.start)
            } else {
                Value::Error(ErrorKind::Value)
            }
        }
        ArgValue::ReferenceUnion(_) => Value::Error(ErrorKind::Value),
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
                match el.coerce_to_bool_with_ctx(ctx) {
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
                    if cond.coerce_to_bool_with_ctx(ctx)? {
                        Ok(t.clone())
                    } else {
                        Ok(f.clone())
                    }
                });
            }

            if needs_true {
                let true_val = array_lift::eval_arg(ctx, value_expr);
                return array_lift::lift2(cond_val, true_val, |cond, t| {
                    if cond.coerce_to_bool_with_ctx(ctx)? {
                        Ok(t.clone())
                    } else {
                        Ok(Value::Error(ErrorKind::NA))
                    }
                });
            }

            if needs_false {
                let false_val = ifs_pairs(ctx, remaining);
                return array_lift::lift2(cond_val, false_val, |cond, f| {
                    if cond.coerce_to_bool_with_ctx(ctx)? {
                        Ok(Value::Error(ErrorKind::NA))
                    } else {
                        Ok(f.clone())
                    }
                });
            }

            // The condition array contains only errors/invalid values, so avoid forcing evaluation of
            // any branch expressions and simply map the coercion errors.
            array_lift::lift1(cond_val, |cond| {
                let _ = cond.coerce_to_bool_with_ctx(ctx)?;
                Ok(Value::Error(ErrorKind::NA))
            })
        }
        other => {
            if let Value::Error(e) = other {
                return Value::Error(e);
            }

            let cond = match other.coerce_to_bool_with_ctx(ctx) {
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

inventory::submit! {
    FunctionSpec {
        name: "SWITCH",
        min_args: 3,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: switch_fn,
    }
}

fn switch_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.len() < 3 {
        return Value::Error(ErrorKind::Value);
    }

    let expr_val = array_lift::eval_arg(ctx, &args[0]);
    if let Value::Error(e) = expr_val {
        // Match Excel/IF-like semantics: if the discriminant is an error, return it without
        // forcing evaluation of any case expressions or the default branch.
        return Value::Error(e);
    }
    let (pairs, default) = if (args.len() - 1) % 2 == 0 {
        (&args[1..], None)
    } else {
        (&args[1..args.len() - 1], Some(&args[args.len() - 1]))
    };

    if pairs.len() < 2 || pairs.len() % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }

    switch_pairs(ctx, expr_val, pairs, default)
}

fn switch_pairs(
    ctx: &dyn FunctionContext,
    expr_val: Value,
    pairs: &[CompiledExpr],
    default: Option<&CompiledExpr>,
) -> Value {
    if pairs.is_empty() {
        return default
            .map(|expr| array_lift::eval_arg(ctx, expr))
            .unwrap_or(Value::Error(ErrorKind::NA));
    }
    if pairs.len() < 2 {
        return Value::Error(ErrorKind::Value);
    }

    let case_expr = &pairs[0];
    let result_expr = &pairs[1];
    let remaining = &pairs[2..];

    let case_val = array_lift::eval_arg(ctx, case_expr);
    let cond_val = array_lift::lift2(expr_val.clone(), case_val, |expr, case| {
        Ok(Value::Bool(excel_eq(expr, case)?))
    });

    match cond_val {
        Value::Array(ref arr) => {
            let mut needs_true = false;
            let mut needs_false = false;
            for el in arr.iter() {
                match el.coerce_to_bool_with_ctx(ctx) {
                    Ok(true) => needs_true = true,
                    Ok(false) => needs_false = true,
                    Err(_) => {}
                }
                if needs_true && needs_false {
                    break;
                }
            }

            if needs_true && needs_false {
                let true_val = array_lift::eval_arg(ctx, result_expr);
                let false_val = switch_pairs(ctx, expr_val, remaining, default);
                return array_lift::lift3(cond_val, true_val, false_val, |cond, t, f| {
                    if cond.coerce_to_bool_with_ctx(ctx)? {
                        Ok(t.clone())
                    } else {
                        Ok(f.clone())
                    }
                });
            }

            if needs_true {
                let true_val = array_lift::eval_arg(ctx, result_expr);
                return array_lift::lift2(cond_val, true_val, |cond, t| {
                    if cond.coerce_to_bool_with_ctx(ctx)? {
                        Ok(t.clone())
                    } else {
                        Ok(Value::Error(ErrorKind::NA))
                    }
                });
            }

            if needs_false {
                let false_val = switch_pairs(ctx, expr_val, remaining, default);
                return array_lift::lift2(cond_val, false_val, |cond, f| {
                    if cond.coerce_to_bool_with_ctx(ctx)? {
                        Ok(Value::Error(ErrorKind::NA))
                    } else {
                        Ok(f.clone())
                    }
                });
            }

            // Only errors/invalid comparisons; map the coercion errors without forcing any branch evaluation.
            array_lift::lift1(cond_val, |cond| {
                let _ = cond.coerce_to_bool_with_ctx(ctx)?;
                Ok(Value::Error(ErrorKind::NA))
            })
        }
        other => {
            if let Value::Error(e) = other {
                return Value::Error(e);
            }
            let matched = match other.coerce_to_bool_with_ctx(ctx) {
                Ok(b) => b,
                Err(e) => return Value::Error(e),
            };

            if matched {
                array_lift::eval_arg(ctx, result_expr)
            } else {
                switch_pairs(ctx, expr_val, remaining, default)
            }
        }
    }
}

fn excel_eq(left: &Value, right: &Value) -> Result<bool, ErrorKind> {
    if let Value::Error(e) = left {
        return Err(*e);
    }
    if let Value::Error(e) = right {
        return Err(*e);
    }

    fn normalize_rich(value: Value) -> Result<Value, ErrorKind> {
        match value {
            Value::Entity(entity) => Ok(Value::Text(entity.display)),
            Value::Record(record) => {
                if let Some(display_field) = record.display_field.as_deref() {
                    if let Some(value) = record.get_field_case_insensitive(display_field) {
                        return Ok(Value::Text(value.coerce_to_string()?));
                    }
                }
                Ok(Value::Text(record.display))
            }
            other => Ok(other),
        }
    }

    let left = normalize_rich(left.clone())?;
    let right = normalize_rich(right.clone())?;
    if matches!(
        &left,
        Value::Array(_)
            | Value::Record(_)
            | Value::Lambda(_)
            | Value::Spill { .. }
            | Value::Reference(_)
            | Value::ReferenceUnion(_)
    ) || matches!(
        &right,
        Value::Array(_)
            | Value::Record(_)
            | Value::Lambda(_)
            | Value::Spill { .. }
            | Value::Reference(_)
            | Value::ReferenceUnion(_)
    ) {
        return Err(ErrorKind::Value);
    }

    // Blank coerces to the other type for comparisons.
    let (l, r) = match (left, right) {
        (Value::Blank, Value::Number(b)) => (Value::Number(0.0), Value::Number(b)),
        (Value::Number(a), Value::Blank) => (Value::Number(a), Value::Number(0.0)),
        (Value::Blank, Value::Bool(b)) => (Value::Bool(false), Value::Bool(b)),
        (Value::Bool(a), Value::Blank) => (Value::Bool(a), Value::Bool(false)),
        (Value::Blank, Value::Text(b)) => (Value::Text(String::new()), Value::Text(b)),
        (Value::Text(a), Value::Blank) => (Value::Text(a), Value::Text(String::new())),
        (l, r) => (l, r),
    };

    let ord = match (&l, &r) {
        (Value::Number(a), Value::Number(b)) => a.partial_cmp(b).unwrap_or(Ordering::Equal),
        (Value::Text(a), Value::Text(b)) => cmp_case_insensitive(a, b),
        (Value::Bool(a), Value::Bool(b)) => a.cmp(b),
        (Value::Entity(a), Value::Entity(b)) => cmp_case_insensitive(&a.display, &b.display),
        (Value::Record(a), Value::Record(b)) => cmp_case_insensitive(&a.display, &b.display),
        (Value::Entity(a), Value::Text(b)) => cmp_case_insensitive(&a.display, b),
        (Value::Text(a), Value::Entity(b)) => cmp_case_insensitive(a, &b.display),
        (Value::Record(a), Value::Text(b)) => cmp_case_insensitive(&a.display, b),
        (Value::Text(a), Value::Record(b)) => cmp_case_insensitive(a, &b.display),
        (Value::Entity(a), Value::Record(b)) => cmp_case_insensitive(&a.display, &b.display),
        (Value::Record(a), Value::Entity(b)) => cmp_case_insensitive(&a.display, &b.display),
        // Type precedence (approximate Excel): numbers < text < booleans.
        (
            Value::Number(_),
            Value::Text(_) | Value::Entity(_) | Value::Record(_) | Value::Bool(_),
        ) => Ordering::Less,
        (Value::Text(_) | Value::Entity(_) | Value::Record(_), Value::Bool(_)) => Ordering::Less,
        (Value::Text(_) | Value::Entity(_) | Value::Record(_), Value::Number(_)) => {
            Ordering::Greater
        }
        (
            Value::Bool(_),
            Value::Number(_) | Value::Text(_) | Value::Entity(_) | Value::Record(_),
        ) => Ordering::Greater,
        // Blank should have been coerced above.
        (Value::Blank, Value::Blank) => Ordering::Equal,
        (Value::Blank, _) => Ordering::Less,
        (_, Value::Blank) => Ordering::Greater,
        // Errors are handled above.
        (Value::Error(_), _) | (_, Value::Error(_)) => Ordering::Equal,
        (Value::Entity(_), _)
        | (_, Value::Entity(_))
        | (Value::Record(_), _)
        | (_, Value::Record(_)) => Ordering::Equal,
        // Arrays/spill markers/lambdas/references are rejected above.
        (Value::Array(_), _)
        | (_, Value::Array(_))
        | (Value::Lambda(_), _)
        | (_, Value::Lambda(_))
        | (Value::Spill { .. }, _)
        | (_, Value::Spill { .. })
        | (Value::Reference(_), _)
        | (_, Value::Reference(_))
        | (Value::ReferenceUnion(_), _)
        | (_, Value::ReferenceUnion(_)) => Ordering::Equal,
    };

    Ok(ord == Ordering::Equal)
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
