use crate::eval::CompiledExpr;
use crate::functions::array_lift;
use crate::functions::math::criteria::Criteria;
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

const VAR_ARGS: usize = 255;

inventory::submit! {
    FunctionSpec {
        name: "RAND",
        min_args: 0,
        max_args: 0,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[],
        implementation: rand_fn,
    }
}

fn rand_fn(ctx: &dyn FunctionContext, _args: &[CompiledExpr]) -> Value {
    Value::Number(ctx.volatile_rand())
}

inventory::submit! {
    FunctionSpec {
        name: "RANDBETWEEN",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: randbetween_fn,
    }
}

fn randbetween_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let bottom = match eval_scalar_arg(ctx, &args[0]).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let top = match eval_scalar_arg(ctx, &args[1]).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    if !bottom.is_finite() || !top.is_finite() {
        return Value::Error(ErrorKind::Num);
    }

    let low_f = bottom.ceil();
    let high_f = top.floor();
    if low_f < (i64::MIN as f64)
        || low_f > (i64::MAX as f64)
        || high_f < (i64::MIN as f64)
        || high_f > (i64::MAX as f64)
    {
        return Value::Error(ErrorKind::Num);
    }

    let low = low_f as i64;
    let high = high_f as i64;
    if low > high {
        return Value::Error(ErrorKind::Num);
    }

    let span = match high.checked_sub(low).and_then(|d| d.checked_add(1)) {
        Some(v) if v > 0 => v as u64,
        _ => return Value::Error(ErrorKind::Num),
    };

    let offset = (ctx.volatile_rand_u64() % span) as i64;
    Value::Number((low + offset) as f64)
}

inventory::submit! {
    FunctionSpec {
        name: "SUM",
        min_args: 0,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: sum,
    }
}

fn sum(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut acc = 0.0;

    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => match v {
                Value::Error(e) => return Value::Error(e),
                Value::Number(n) => acc += n,
                Value::Bool(b) => acc += if b { 1.0 } else { 0.0 },
                Value::Blank => {}
                Value::Text(s) => match Value::Text(s).coerce_to_number() {
                    Ok(n) => acc += n,
                    Err(e) => return Value::Error(e),
                },
                Value::Array(arr) => {
                    for v in arr.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => acc += n,
                             Value::Bool(_)
                             | Value::Text(_)
                             | Value::Blank
                             | Value::Array(_)
                             | Value::Lambda(_)
                             | Value::Spill { .. }
                             | Value::Reference(_)
                             | Value::ReferenceUnion(_) => {}
                         }
                     }
                 }
                 Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                Value::Spill { .. } => return Value::Error(ErrorKind::Value),
                Value::Reference(_) | Value::ReferenceUnion(_) => {
                    return Value::Error(ErrorKind::Value)
                }
            },
            ArgValue::Reference(r) => {
                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => acc += n,
                        // Excel quirk: logicals/text in references are ignored by SUM.
                         Value::Bool(_)
                         | Value::Text(_)
                         | Value::Blank
                         | Value::Array(_)
                         | Value::Lambda(_)
                         | Value::Spill { .. }
                         | Value::Reference(_)
                         | Value::ReferenceUnion(_) => {}
                     }
                 }
             }
             ArgValue::ReferenceUnion(ranges) => {
                let mut seen = std::collections::HashSet::new();
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => acc += n,
                            // Excel quirk: logicals/text in references are ignored by SUM.
                             Value::Bool(_)
                             | Value::Text(_)
                             | Value::Blank
                             | Value::Array(_)
                             | Value::Lambda(_)
                             | Value::Spill { .. }
                             | Value::Reference(_)
                             | Value::ReferenceUnion(_) => {}
                         }
                     }
                 }
             }
         }
    }

    Value::Number(acc)
}

inventory::submit! {
    FunctionSpec {
        name: "AVERAGE",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: average,
    }
}

fn average(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut acc = 0.0;
    let mut count = 0u64;

    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => match v {
                Value::Error(e) => return Value::Error(e),
                Value::Number(n) => {
                    acc += n;
                    count += 1;
                }
                Value::Bool(b) => {
                    acc += if b { 1.0 } else { 0.0 };
                    count += 1;
                }
                Value::Blank => {}
                Value::Text(s) => match Value::Text(s).coerce_to_number() {
                    Ok(n) => {
                        acc += n;
                        count += 1;
                    }
                    Err(e) => return Value::Error(e),
                },
                Value::Array(arr) => {
                    for v in arr.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => {
                                acc += n;
                                count += 1;
                            }
                             Value::Bool(_)
                             | Value::Text(_)
                             | Value::Blank
                             | Value::Array(_)
                             | Value::Lambda(_)
                             | Value::Spill { .. }
                             | Value::Reference(_)
                             | Value::ReferenceUnion(_) => {}
                         }
                     }
                 }
                 Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                Value::Spill { .. } => return Value::Error(ErrorKind::Value),
                Value::Reference(_) | Value::ReferenceUnion(_) => {
                    return Value::Error(ErrorKind::Value)
                }
            },
            ArgValue::Reference(r) => {
                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => {
                            acc += n;
                            count += 1;
                        }
                        // Ignore logical/text/blank in references.
                         Value::Bool(_)
                         | Value::Text(_)
                         | Value::Blank
                         | Value::Array(_)
                         | Value::Lambda(_)
                         | Value::Spill { .. }
                         | Value::Reference(_)
                         | Value::ReferenceUnion(_) => {}
                     }
                 }
             }
             ArgValue::ReferenceUnion(ranges) => {
                let mut seen = std::collections::HashSet::new();
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => {
                                acc += n;
                                count += 1;
                            }
                            // Ignore logical/text/blank in references.
                             Value::Bool(_)
                             | Value::Text(_)
                             | Value::Blank
                             | Value::Array(_)
                             | Value::Lambda(_)
                             | Value::Spill { .. }
                             | Value::Reference(_)
                             | Value::ReferenceUnion(_) => {}
                         }
                     }
                 }
             }
         }
    }

    if count == 0 {
        return Value::Error(ErrorKind::Div0);
    }
    Value::Number(acc / (count as f64))
}

inventory::submit! {
    FunctionSpec {
        name: "MIN",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: min_fn,
    }
}

fn min_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut best: Option<f64> = None;

    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => match v {
                Value::Array(arr) => {
                    for v in arr.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => best = Some(best.map(|b| b.min(*n)).unwrap_or(*n)),
                             Value::Bool(_)
                             | Value::Text(_)
                             | Value::Blank
                             | Value::Array(_)
                             | Value::Lambda(_)
                             | Value::Spill { .. }
                             | Value::Reference(_)
                             | Value::ReferenceUnion(_) => {}
                         }
                     }
                 }
                 other => {
                     let n = match other.coerce_to_number() {
                        Ok(n) => n,
                        Err(e) => return Value::Error(e),
                    };
                    best = Some(best.map(|b| b.min(n)).unwrap_or(n));
                }
            },
            ArgValue::Reference(r) => {
                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => best = Some(best.map(|b| b.min(n)).unwrap_or(n)),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Lambda(_)
                        | Value::Spill { .. }
                        | Value::Reference(_)
                        | Value::ReferenceUnion(_) => {}
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                let mut seen = std::collections::HashSet::new();
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => best = Some(best.map(|b| b.min(n)).unwrap_or(n)),
                             Value::Bool(_)
                             | Value::Text(_)
                             | Value::Blank
                             | Value::Array(_)
                             | Value::Lambda(_)
                             | Value::Spill { .. }
                             | Value::Reference(_)
                             | Value::ReferenceUnion(_) => {}
                         }
                     }
                 }
             }
         }
    }

    Value::Number(best.unwrap_or(0.0))
}

inventory::submit! {
    FunctionSpec {
        name: "MAX",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: max_fn,
    }
}

fn max_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut best: Option<f64> = None;

    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => match v {
                Value::Array(arr) => {
                    for v in arr.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => best = Some(best.map(|b| b.max(*n)).unwrap_or(*n)),
                             Value::Bool(_)
                             | Value::Text(_)
                             | Value::Blank
                             | Value::Array(_)
                             | Value::Lambda(_)
                             | Value::Spill { .. }
                             | Value::Reference(_)
                             | Value::ReferenceUnion(_) => {}
                         }
                     }
                 }
                 other => {
                     let n = match other.coerce_to_number() {
                        Ok(n) => n,
                        Err(e) => return Value::Error(e),
                    };
                    best = Some(best.map(|b| b.max(n)).unwrap_or(n));
                }
            },
            ArgValue::Reference(r) => {
                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => best = Some(best.map(|b| b.max(n)).unwrap_or(n)),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Lambda(_)
                        | Value::Spill { .. }
                        | Value::Reference(_)
                        | Value::ReferenceUnion(_) => {}
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                let mut seen = std::collections::HashSet::new();
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => best = Some(best.map(|b| b.max(n)).unwrap_or(n)),
                             Value::Bool(_)
                             | Value::Text(_)
                             | Value::Blank
                             | Value::Array(_)
                             | Value::Lambda(_)
                             | Value::Spill { .. }
                             | Value::Reference(_)
                             | Value::ReferenceUnion(_) => {}
                         }
                     }
                 }
             }
         }
    }

    Value::Number(best.unwrap_or(0.0))
}

inventory::submit! {
    FunctionSpec {
        name: "COUNT",
        min_args: 0,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: count_fn,
    }
}

fn count_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut total = 0u64;
    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => match v {
                Value::Array(arr) => {
                    for v in arr.iter() {
                        if matches!(v, Value::Number(_)) {
                            total += 1;
                        }
                    }
                }
                other => {
                    if matches!(other, Value::Number(_)) {
                        total += 1;
                    }
                }
            },
            ArgValue::Reference(r) => {
                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    if matches!(v, Value::Number(_)) {
                        total += 1;
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                let mut seen = std::collections::HashSet::new();
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        if matches!(v, Value::Number(_)) {
                            total += 1;
                        }
                    }
                }
            }
        }
    }
    Value::Number(total as f64)
}

inventory::submit! {
    FunctionSpec {
        name: "COUNTIF",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: countif_fn,
    }
}

fn countif_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let criteria_value = eval_scalar_arg(ctx, &args[1]);
    if let Value::Error(e) = criteria_value {
        return Value::Error(e);
    }
    let criteria = match Criteria::parse_with_date_system(&criteria_value, ctx.date_system()) {
        Ok(c) => c,
        Err(e) => return Value::Error(e),
    };

    fn reference_cell_count(reference: &crate::functions::Reference) -> u64 {
        let r = reference.normalized();
        let rows = (u64::from(r.end.row)).saturating_sub(u64::from(r.start.row)) + 1;
        let cols = (u64::from(r.end.col)).saturating_sub(u64::from(r.start.col)) + 1;
        rows.saturating_mul(cols)
    }

    // `iter_reference_cells` is sparse when the backend supports it. COUNTIF must still account
    // for implicit blanks, but only when the criteria can actually match blank cells.
    let blank_matches = criteria.matches(&Value::Blank);

    let mut count = 0u64;
    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(r) => {
            let mut seen = 0u64;
            for addr in ctx.iter_reference_cells(&r) {
                seen += 1;
                let v = ctx.get_cell_value(&r.sheet_id, addr);
                if criteria.matches(&v) {
                    count += 1;
                }
            }
            if blank_matches {
                count += reference_cell_count(&r).saturating_sub(seen);
            }
        }
        ArgValue::ReferenceUnion(ranges) => {
            let mut seen = std::collections::HashSet::new();
            for r in ranges {
                if blank_matches {
                    for addr in r.iter_cells() {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        if criteria.matches(&v) {
                            count += 1;
                        }
                    }
                } else {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        if criteria.matches(&v) {
                            count += 1;
                        }
                    }
                }
            }
        }
        ArgValue::Scalar(Value::Array(arr)) => {
            for v in arr.iter() {
                if criteria.matches(v) {
                    count += 1;
                }
            }
        }
        ArgValue::Scalar(Value::Error(e)) => return Value::Error(e),
        ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    }
    Value::Number(count as f64)
}

inventory::submit! {
    FunctionSpec {
        name: "SUMPRODUCT",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: sumproduct_fn,
    }
}

fn sumproduct_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    fn arg_to_values(ctx: &dyn FunctionContext, arg: ArgValue) -> Result<Vec<Value>, Value> {
        match arg {
            ArgValue::Reference(r) => {
                let r = r.normalized();
                let rows = r.end.row - r.start.row + 1;
                let cols = r.end.col - r.start.col + 1;
                let len = (rows as usize).saturating_mul(cols as usize);
                let mut values = Vec::with_capacity(len);
                for addr in r.iter_cells() {
                    values.push(ctx.get_cell_value(&r.sheet_id, addr));
                }
                Ok(values)
            }
            ArgValue::ReferenceUnion(_) => Err(Value::Error(ErrorKind::Value)),
            ArgValue::Scalar(Value::Array(arr)) => Ok(arr.values),
            ArgValue::Scalar(Value::Error(e)) => Err(Value::Error(e)),
            ArgValue::Scalar(v) => Ok(vec![v]),
        }
    }

    let va = match arg_to_values(ctx, ctx.eval_arg(&args[0])) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let vb = match arg_to_values(ctx, ctx.eval_arg(&args[1])) {
        Ok(v) => v,
        Err(e) => return e,
    };

    // SUMPRODUCT broadcasts 1x1 scalars to the other array length.
    let (va, vb) = match (va.len(), vb.len()) {
        (0, _) | (_, 0) => return Value::Error(ErrorKind::Value),
        (1, len) if len != 1 => (vec![va[0].clone(); len], vb),
        (len, 1) if len != 1 => (va, vec![vb[0].clone(); len]),
        _ => (va, vb),
    };

    match crate::functions::math::sumproduct(&[&va, &vb]) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "COUNTA",
        min_args: 0,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: counta_fn,
    }
}

fn counta_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut total = 0u64;
    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => match v {
                Value::Array(arr) => {
                    for v in arr.iter() {
                        if !matches!(v, Value::Blank) {
                            total += 1;
                        }
                    }
                }
                other => {
                    if !matches!(other, Value::Blank) {
                        total += 1;
                    }
                }
            },
            ArgValue::Reference(r) => {
                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    if !matches!(v, Value::Blank) {
                        total += 1;
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                let mut seen = std::collections::HashSet::new();
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        if !matches!(v, Value::Blank) {
                            total += 1;
                        }
                    }
                }
            }
        }
    }
    Value::Number(total as f64)
}

inventory::submit! {
    FunctionSpec {
        name: "COUNTBLANK",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: countblank_fn,
    }
}

fn countblank_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut total = 0u64;
    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => match v {
                Value::Array(arr) => {
                    for v in arr.iter() {
                        if matches!(v, Value::Blank)
                            || matches!(v, Value::Text(ref s) if s.is_empty())
                        {
                            total += 1;
                        }
                    }
                }
                other => {
                    if matches!(other, Value::Blank)
                        || matches!(other, Value::Text(ref s) if s.is_empty())
                    {
                        total += 1;
                    }
                }
            },
            ArgValue::Reference(r) => {
                let size = r.size();
                let mut non_blank = 0u64;
                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    if !matches!(v, Value::Blank)
                        && !matches!(v, Value::Text(ref s) if s.is_empty())
                    {
                        non_blank += 1;
                    }
                }
                total += size.saturating_sub(non_blank);
            }
            ArgValue::ReferenceUnion(ranges) => {
                let size = reference_union_size(&ranges);
                let mut seen = std::collections::HashSet::new();
                let mut non_blank = 0u64;
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        if !matches!(v, Value::Blank)
                            && !matches!(v, Value::Text(ref s) if s.is_empty())
                        {
                            non_blank += 1;
                        }
                    }
                }
                total += size.saturating_sub(non_blank);
            }
        }
    }
    Value::Number(total as f64)
}

fn reference_union_size(ranges: &[crate::functions::Reference]) -> u64 {
    let rects: Vec<crate::functions::Reference> = ranges.iter().map(|r| r.normalized()).collect();
    if rects.is_empty() {
        return 0;
    }

    // Convert to half-open row slabs: [start, end+1)
    let mut row_bounds: Vec<u32> = Vec::with_capacity(rects.len() * 2);
    for r in &rects {
        row_bounds.push(r.start.row);
        row_bounds.push(r.end.row.saturating_add(1));
    }
    row_bounds.sort_unstable();
    row_bounds.dedup();

    let mut total: u64 = 0;
    for rows in row_bounds.windows(2) {
        let y0 = rows[0];
        let y1 = rows[1];
        if y1 <= y0 {
            continue;
        }

        let mut intervals: Vec<(u32, u32)> = Vec::new();
        for r in &rects {
            let r_end = r.end.row.saturating_add(1);
            if r.start.row <= y0 && r_end >= y1 {
                intervals.push((r.start.col, r.end.col.saturating_add(1)));
            }
        }

        if intervals.is_empty() {
            continue;
        }

        intervals.sort_by_key(|(s, _e)| *s);

        let mut cur_s = intervals[0].0;
        let mut cur_e = intervals[0].1;
        let mut len: u64 = 0;
        for (s, e) in intervals.into_iter().skip(1) {
            if s > cur_e {
                len += (cur_e - cur_s) as u64;
                cur_s = s;
                cur_e = e;
            } else {
                cur_e = cur_e.max(e);
            }
        }
        len += (cur_e - cur_s) as u64;

        total += (y1 - y0) as u64 * len;
    }

    total
}

inventory::submit! {
    FunctionSpec {
        name: "ROUND",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: round_fn,
    }
}

fn round_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    round_impl(ctx, args, RoundMode::Nearest)
}

inventory::submit! {
    FunctionSpec {
        name: "ROUNDDOWN",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: rounddown_fn,
    }
}

fn rounddown_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    round_impl(ctx, args, RoundMode::Down)
}

inventory::submit! {
    FunctionSpec {
        name: "ROUNDUP",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: roundup_fn,
    }
}

fn roundup_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    round_impl(ctx, args, RoundMode::Up)
}

inventory::submit! {
    FunctionSpec {
        name: "TRUNC",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: trunc_fn,
    }
}

fn trunc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    let digits = if args.len() == 2 {
        array_lift::eval_arg(ctx, &args[1])
    } else {
        Value::Number(0.0)
    };
    array_lift::lift2(number, digits, |number, digits| {
        let number = number.coerce_to_number()?;
        let digits = digits.coerce_to_i64()?;
        Ok(Value::Number(round_with_mode(
            number,
            digits as i32,
            RoundMode::Down,
        )))
    })
}

fn round_impl(ctx: &dyn FunctionContext, args: &[CompiledExpr], mode: RoundMode) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    let digits = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(number, digits, |number, digits| {
        let number = number.coerce_to_number()?;
        let digits = digits.coerce_to_i64()?;
        Ok(Value::Number(round_with_mode(number, digits as i32, mode)))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "INT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: int_fn,
    }
}

fn int_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(value, |v| Ok(Value::Number(v.coerce_to_number()?.floor())))
}

inventory::submit! {
    FunctionSpec {
        name: "ABS",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: abs_fn,
    }
}

fn abs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(value, |v| Ok(Value::Number(v.coerce_to_number()?.abs())))
}

inventory::submit! {
    FunctionSpec {
        name: "MOD",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: mod_fn,
    }
}

fn mod_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let n = array_lift::eval_arg(ctx, &args[0]);
    let d = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(n, d, |n, d| {
        let n = n.coerce_to_number()?;
        let d = d.coerce_to_number()?;
        if d == 0.0 {
            return Err(ErrorKind::Div0);
        }
        Ok(Value::Number(n - d * (n / d).floor()))
    })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum RoundMode {
    Nearest,
    Down,
    Up,
}

fn round_with_mode(n: f64, digits: i32, mode: RoundMode) -> f64 {
    let factor = 10f64.powi(digits.saturating_abs());
    if !factor.is_finite() || factor == 0.0 {
        return n;
    }

    let scaled = if digits >= 0 { n * factor } else { n / factor };
    let rounded = match mode {
        RoundMode::Down => scaled.trunc(),
        RoundMode::Up => {
            if scaled.is_sign_negative() {
                scaled.trunc() - if scaled.fract() == 0.0 { 0.0 } else { 1.0 }
            } else {
                scaled.trunc() + if scaled.fract() == 0.0 { 0.0 } else { 1.0 }
            }
        }
        RoundMode::Nearest => {
            // Excel rounds halves away from zero.
            let frac = scaled.fract().abs();
            let base = scaled.trunc();
            if frac < 0.5 {
                base
            } else {
                base + scaled.signum()
            }
        }
    };

    if digits >= 0 {
        rounded / factor
    } else {
        rounded * factor
    }
}
inventory::submit! {
    FunctionSpec {
        name: "SIGN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: sign_fn,
    }
}

fn sign_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(number, |number| {
        let number = number.coerce_to_number()?;
        if !number.is_finite() {
            return Err(ErrorKind::Num);
        }
        if number > 0.0 {
            Ok(Value::Number(1.0))
        } else if number < 0.0 {
            Ok(Value::Number(-1.0))
        } else {
            Ok(Value::Number(0.0))
        }
    })
}
