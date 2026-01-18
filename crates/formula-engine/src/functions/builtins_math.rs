use crate::eval::{CellAddr, CompiledExpr};
use crate::functions::array_lift;
use crate::functions::math::criteria::Criteria;
use crate::functions::{
    eval_scalar_arg, volatile_rand_u64_below, ArgValue, ArraySupport, FunctionContext, FunctionSpec,
};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::simd;
use crate::value::{parse_number, Array, ErrorKind, Value};

const VAR_ARGS: usize = 255;
const SIMD_AGGREGATE_BLOCK: usize = 1024;

// Rich values (Entity/Record) are treated like text for aggregate functions:
// - In scalar positions (e.g. `=SUM(A1, Entity)`), they are non-numeric and return `#VALUE!`.
// - When iterating references/arrays (e.g. `=SUM(A1:A2)`), they are ignored like text/logicals.

// Criteria aggregates (SUMIF/AVERAGEIF) require extra per-element coercion work before the SIMD
// kernels can run. Avoid paying that overhead for tiny ranges.
const SIMD_CRITERIA_ARRAY_MIN_LEN: usize = 32;

#[inline]
fn count_value_to_f64(v: &Value) -> f64 {
    match v {
        // COUNT counts numeric cells and ignores everything else (including errors).
        Value::Number(_) => 0.0,
        _ => f64::NAN,
    }
}

#[inline]
fn coerce_countif_value_to_number(v: &Value, locale: crate::value::NumberLocale) -> Option<f64> {
    match v {
        Value::Number(n) => Some(*n),
        Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
        // COUNTIF numeric criteria treat blank as 0 for comparison.
        Value::Blank => Some(0.0),
        Value::Text(s) => parse_number(s, locale).ok(),
        Value::Entity(_) | Value::Record(_) => None,
        // Criteria matching uses implicit intersection for array candidates.
        Value::Array(arr) => coerce_countif_value_to_number(&arr.top_left(), locale),
        Value::Error(_)
        | Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => None,
    }
}

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
    let bottom = match eval_scalar_arg(ctx, &args[0]).coerce_to_number_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let top = match eval_scalar_arg(ctx, &args[1]).coerce_to_number_with_ctx(ctx) {
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

    let offset = volatile_rand_u64_below(ctx, span) as i64;
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
                Value::Text(s) => match Value::Text(s).coerce_to_number_with_ctx(ctx) {
                    Ok(n) => acc += n,
                    Err(e) => return Value::Error(e),
                },
                Value::Entity(_) | Value::Record(_) => return Value::Error(ErrorKind::Value),
                Value::Array(arr) => {
                    let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                    let mut len = 0usize;
                    let mut sum = 0.0;
                    let mut saw_nan = false;

                    for v in arr.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan = true;
                                } else if !saw_nan {
                                    buf[len] = *n;
                                    len += 1;
                                    if len == SIMD_AGGREGATE_BLOCK {
                                        sum += simd::sum_ignore_nan_f64(&buf);
                                        len = 0;
                                    }
                                }
                            }
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
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

                    if saw_nan {
                        acc += f64::NAN;
                    } else {
                        if len > 0 {
                            sum += simd::sum_ignore_nan_f64(&buf[..len]);
                        }
                        acc += sum;
                    }
                }
                Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                Value::Spill { .. } => return Value::Error(ErrorKind::Value),
                Value::Reference(_) | Value::ReferenceUnion(_) => {
                    return Value::Error(ErrorKind::Value)
                }
            },
            ArgValue::Reference(r) => {
                let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                let mut len = 0usize;
                let mut sum = 0.0;
                let mut saw_nan = false;

                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => {
                            if n.is_nan() {
                                saw_nan = true;
                            } else if !saw_nan {
                                buf[len] = n;
                                len += 1;
                                if len == SIMD_AGGREGATE_BLOCK {
                                    sum += simd::sum_ignore_nan_f64(&buf);
                                    len = 0;
                                }
                            }
                        }
                        Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                        // Excel quirk: logicals/text in references are ignored by SUM.
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

                if saw_nan {
                    acc += f64::NAN;
                } else {
                    if len > 0 {
                        sum += simd::sum_ignore_nan_f64(&buf[..len]);
                    }
                    acc += sum;
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                let mut seen = std::collections::HashSet::new();
                let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                let mut len = 0usize;
                let mut sum = 0.0;
                let mut saw_nan = false;
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan = true;
                                } else if !saw_nan {
                                    buf[len] = n;
                                    len += 1;
                                    if len == SIMD_AGGREGATE_BLOCK {
                                        sum += simd::sum_ignore_nan_f64(&buf);
                                        len = 0;
                                    }
                                }
                            }
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                            // Excel quirk: logicals/text in references are ignored by SUM.
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

                if saw_nan {
                    acc += f64::NAN;
                } else {
                    if len > 0 {
                        sum += simd::sum_ignore_nan_f64(&buf[..len]);
                    }
                    acc += sum;
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
    let mut saw_nan = false;

    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => match v {
                Value::Error(e) => return Value::Error(e),
                Value::Number(n) => {
                    if n.is_nan() {
                        saw_nan = true;
                        count += 1;
                    } else if !saw_nan {
                        acc += n;
                        count += 1;
                    }
                }
                Value::Bool(b) => {
                    // Scalar logicals participate in AVERAGE.
                    if !saw_nan {
                        acc += if b { 1.0 } else { 0.0 };
                        count += 1;
                    }
                }
                Value::Blank => {}
                Value::Text(s) => {
                    // Scalar text is coerced (and errors propagate), even if we've already seen a
                    // NaN elsewhere.
                    let n = match Value::Text(s).coerce_to_number_with_ctx(ctx) {
                        Ok(n) => n,
                        Err(e) => return Value::Error(e),
                    };
                    if n.is_nan() {
                        saw_nan = true;
                        count += 1;
                    } else if !saw_nan {
                        acc += n;
                        count += 1;
                    }
                }
                Value::Entity(_) | Value::Record(_) => return Value::Error(ErrorKind::Value),
                Value::Array(arr) => {
                    let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                    let mut len = 0usize;
                    let mut sum = 0.0;
                    for v in arr.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan = true;
                                    count += 1;
                                } else if !saw_nan {
                                    buf[len] = *n;
                                    len += 1;
                                    count += 1;
                                    if len == SIMD_AGGREGATE_BLOCK {
                                        sum += simd::sum_ignore_nan_f64(&buf);
                                        len = 0;
                                    }
                                }
                            }
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
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

                    if !saw_nan {
                        if len > 0 {
                            sum += simd::sum_ignore_nan_f64(&buf[..len]);
                        }
                        acc += sum;
                    }
                }
                Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                Value::Spill { .. } => return Value::Error(ErrorKind::Value),
                Value::Reference(_) | Value::ReferenceUnion(_) => {
                    return Value::Error(ErrorKind::Value)
                }
            },
            ArgValue::Reference(r) => {
                let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                let mut len = 0usize;
                let mut sum = 0.0;
                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => {
                            if n.is_nan() {
                                saw_nan = true;
                                count += 1;
                            } else if !saw_nan {
                                buf[len] = n;
                                len += 1;
                                count += 1;
                                if len == SIMD_AGGREGATE_BLOCK {
                                    sum += simd::sum_ignore_nan_f64(&buf);
                                    len = 0;
                                }
                            }
                        }
                        Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                        // Ignore logical/text/blank in references.
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

                if !saw_nan {
                    if len > 0 {
                        sum += simd::sum_ignore_nan_f64(&buf[..len]);
                    }
                    acc += sum;
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                let mut seen = std::collections::HashSet::new();
                let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                let mut len = 0usize;
                let mut sum = 0.0;
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan = true;
                                    count += 1;
                                } else if !saw_nan {
                                    buf[len] = n;
                                    len += 1;
                                    count += 1;
                                    if len == SIMD_AGGREGATE_BLOCK {
                                        sum += simd::sum_ignore_nan_f64(&buf);
                                        len = 0;
                                    }
                                }
                            }
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                            // Ignore logical/text/blank in references.
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

                if !saw_nan {
                    if len > 0 {
                        sum += simd::sum_ignore_nan_f64(&buf[..len]);
                    }
                    acc += sum;
                }
            }
        }
    }

    if count == 0 {
        return Value::Error(ErrorKind::Div0);
    }
    if saw_nan {
        Value::Number(f64::NAN)
    } else {
        Value::Number(acc / (count as f64))
    }
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
    let mut saw_nan_number = false;

    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => match v {
                Value::Array(arr) => {
                    let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                    let mut len = 0usize;
                    let mut local_best: Option<f64> = None;

                    for v in arr.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan_number = true;
                                    continue;
                                }

                                buf[len] = *n;
                                len += 1;
                                if len == SIMD_AGGREGATE_BLOCK {
                                    if let Some(m) = simd::min_ignore_nan_f64(&buf) {
                                        local_best = Some(local_best.map_or(m, |b| b.min(m)));
                                    }
                                    len = 0;
                                }
                            }
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
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

                    if len > 0 {
                        if let Some(m) = simd::min_ignore_nan_f64(&buf[..len]) {
                            local_best = Some(local_best.map_or(m, |b| b.min(m)));
                        }
                    }
                    if let Some(m) = local_best {
                        best = Some(best.map_or(m, |b| b.min(m)));
                    }
                }
                other => {
                    let n = match other.coerce_to_number_with_ctx(ctx) {
                        Ok(n) => n,
                        Err(e) => return Value::Error(e),
                    };
                    if n.is_nan() {
                        saw_nan_number = true;
                    } else {
                        best = Some(best.map_or(n, |b| b.min(n)));
                    }
                }
            },
            ArgValue::Reference(r) => {
                let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                let mut len = 0usize;
                let mut local_best: Option<f64> = None;

                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => {
                            if n.is_nan() {
                                saw_nan_number = true;
                                continue;
                            }

                            buf[len] = n;
                            len += 1;
                            if len == SIMD_AGGREGATE_BLOCK {
                                if let Some(m) = simd::min_ignore_nan_f64(&buf) {
                                    local_best = Some(local_best.map_or(m, |b| b.min(m)));
                                }
                                len = 0;
                            }
                        }
                        Value::Lambda(_) => return Value::Error(ErrorKind::Value),
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

                if len > 0 {
                    if let Some(m) = simd::min_ignore_nan_f64(&buf[..len]) {
                        local_best = Some(local_best.map_or(m, |b| b.min(m)));
                    }
                }
                if let Some(m) = local_best {
                    best = Some(best.map_or(m, |b| b.min(m)));
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                let mut seen = std::collections::HashSet::new();
                let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                let mut len = 0usize;
                let mut local_best: Option<f64> = None;
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan_number = true;
                                    continue;
                                }

                                buf[len] = n;
                                len += 1;
                                if len == SIMD_AGGREGATE_BLOCK {
                                    if let Some(m) = simd::min_ignore_nan_f64(&buf) {
                                        local_best = Some(local_best.map_or(m, |b| b.min(m)));
                                    }
                                    len = 0;
                                }
                            }
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
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

                if len > 0 {
                    if let Some(m) = simd::min_ignore_nan_f64(&buf[..len]) {
                        local_best = Some(local_best.map_or(m, |b| b.min(m)));
                    }
                }
                if let Some(m) = local_best {
                    best = Some(best.map_or(m, |b| b.min(m)));
                }
            }
        }
    }

    match best {
        Some(n) => Value::Number(n),
        None if saw_nan_number => Value::Number(f64::NAN),
        None => Value::Number(0.0),
    }
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
    let mut saw_nan_number = false;

    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => match v {
                Value::Array(arr) => {
                    let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                    let mut len = 0usize;
                    let mut local_best: Option<f64> = None;

                    for v in arr.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan_number = true;
                                    continue;
                                }

                                buf[len] = *n;
                                len += 1;
                                if len == SIMD_AGGREGATE_BLOCK {
                                    if let Some(m) = simd::max_ignore_nan_f64(&buf) {
                                        local_best = Some(local_best.map_or(m, |b| b.max(m)));
                                    }
                                    len = 0;
                                }
                            }
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
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

                    if len > 0 {
                        if let Some(m) = simd::max_ignore_nan_f64(&buf[..len]) {
                            local_best = Some(local_best.map_or(m, |b| b.max(m)));
                        }
                    }
                    if let Some(m) = local_best {
                        best = Some(best.map_or(m, |b| b.max(m)));
                    }
                }
                other => {
                    let n = match other.coerce_to_number_with_ctx(ctx) {
                        Ok(n) => n,
                        Err(e) => return Value::Error(e),
                    };
                    if n.is_nan() {
                        saw_nan_number = true;
                    } else {
                        best = Some(best.map_or(n, |b| b.max(n)));
                    }
                }
            },
            ArgValue::Reference(r) => {
                let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                let mut len = 0usize;
                let mut local_best: Option<f64> = None;

                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => {
                            if n.is_nan() {
                                saw_nan_number = true;
                                continue;
                            }

                            buf[len] = n;
                            len += 1;
                            if len == SIMD_AGGREGATE_BLOCK {
                                if let Some(m) = simd::max_ignore_nan_f64(&buf) {
                                    local_best = Some(local_best.map_or(m, |b| b.max(m)));
                                }
                                len = 0;
                            }
                        }
                        Value::Lambda(_) => return Value::Error(ErrorKind::Value),
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

                if len > 0 {
                    if let Some(m) = simd::max_ignore_nan_f64(&buf[..len]) {
                        local_best = Some(local_best.map_or(m, |b| b.max(m)));
                    }
                }
                if let Some(m) = local_best {
                    best = Some(best.map_or(m, |b| b.max(m)));
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                let mut seen = std::collections::HashSet::new();
                let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                let mut len = 0usize;
                let mut local_best: Option<f64> = None;
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => {
                                if n.is_nan() {
                                    saw_nan_number = true;
                                    continue;
                                }

                                buf[len] = n;
                                len += 1;
                                if len == SIMD_AGGREGATE_BLOCK {
                                    if let Some(m) = simd::max_ignore_nan_f64(&buf) {
                                        local_best = Some(local_best.map_or(m, |b| b.max(m)));
                                    }
                                    len = 0;
                                }
                            }
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
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

                if len > 0 {
                    if let Some(m) = simd::max_ignore_nan_f64(&buf[..len]) {
                        local_best = Some(local_best.map_or(m, |b| b.max(m)));
                    }
                }
                if let Some(m) = local_best {
                    best = Some(best.map_or(m, |b| b.max(m)));
                }
            }
        }
    }

    match best {
        Some(n) => Value::Number(n),
        None if saw_nan_number => Value::Number(f64::NAN),
        None => Value::Number(0.0),
    }
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
                    let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                    let mut len = 0usize;
                    let mut count = 0usize;

                    for v in arr.iter() {
                        buf[len] = count_value_to_f64(v);
                        len += 1;
                        if len == SIMD_AGGREGATE_BLOCK {
                            count += simd::count_ignore_nan_f64(&buf);
                            len = 0;
                        }
                    }

                    if len > 0 {
                        count += simd::count_ignore_nan_f64(&buf[..len]);
                    }

                    total += count as u64;
                }
                other => {
                    if matches!(other, Value::Number(_)) {
                        total += 1;
                    }
                }
            },
            ArgValue::Reference(r) => {
                let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                let mut len = 0usize;
                let mut count = 0usize;

                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    buf[len] = count_value_to_f64(&v);
                    len += 1;
                    if len == SIMD_AGGREGATE_BLOCK {
                        count += simd::count_ignore_nan_f64(&buf);
                        len = 0;
                    }
                }

                if len > 0 {
                    count += simd::count_ignore_nan_f64(&buf[..len]);
                }

                total += count as u64;
            }
            ArgValue::ReferenceUnion(ranges) => {
                let mut seen = std::collections::HashSet::new();
                let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                let mut len = 0usize;
                let mut count = 0usize;

                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        buf[len] = count_value_to_f64(&v);
                        len += 1;
                        if len == SIMD_AGGREGATE_BLOCK {
                            count += simd::count_ignore_nan_f64(&buf);
                            len = 0;
                        }
                    }
                }

                if len > 0 {
                    count += simd::count_ignore_nan_f64(&buf[..len]);
                }

                total += count as u64;
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
    fn coerce_candidate_to_number(
        value: &Value,
        locale: crate::value::NumberLocale,
    ) -> Option<f64> {
        match value {
            Value::Number(n) => Some(*n),
            Value::Bool(b) => Some(if *b { 1.0 } else { 0.0 }),
            Value::Blank => Some(0.0),
            Value::Text(s) => parse_number(s, locale).ok(),
            Value::Record(_) | Value::Entity(_) => None,
            Value::Array(arr) => coerce_candidate_to_number(&arr.top_left(), locale),
            Value::Error(_)
            | Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => None,
        }
    }

    let criteria_value = eval_scalar_arg(ctx, &args[1]);
    if let Value::Error(e) = criteria_value {
        return Value::Error(e);
    }
    let criteria = match Criteria::parse_with_date_system_and_locales(
        &criteria_value,
        ctx.date_system(),
        ctx.value_locale(),
        ctx.now_utc(),
        ctx.locale_config(),
    ) {
        Ok(c) => c,
        Err(e) => return Value::Error(e),
    };

    // `iter_reference_cells` is sparse when the backend supports it. COUNTIF must still account
    // for implicit blanks, but only when the criteria can actually match blank cells.
    let blank_matches = criteria.matches(&Value::Blank);
    let numeric = criteria.as_numeric_criteria();
    let number_locale = ctx.number_locale();

    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(r) => {
            if let Some(numeric) = numeric {
                let mut count: u64 = 0;
                let mut seen_count: u64 = 0;
                let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                let mut len = 0usize;

                for addr in ctx.iter_reference_cells(&r) {
                    seen_count += 1;
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    let Some(n) = coerce_candidate_to_number(&v, number_locale) else {
                        continue;
                    };
                    if n.is_nan() {
                        let tmp = Value::Number(n);
                        if criteria.matches(&tmp) {
                            count += 1;
                        }
                        continue;
                    }

                    buf[len] = n;
                    len += 1;
                    if len == SIMD_AGGREGATE_BLOCK {
                        count += simd::count_if_f64(&buf, numeric) as u64;
                        len = 0;
                    }
                }
                if len > 0 {
                    count += simd::count_if_f64(&buf[..len], numeric) as u64;
                }
                if blank_matches {
                    count += r.size().saturating_sub(seen_count);
                }
                return Value::Number(count as f64);
            }

            let mut count: u64 = 0;
            let mut seen_count: u64 = 0;
            for addr in ctx.iter_reference_cells(&r) {
                seen_count += 1;
                let v = ctx.get_cell_value(&r.sheet_id, addr);
                if criteria.matches(&v) {
                    count += 1;
                }
            }

            if blank_matches {
                count += r.size().saturating_sub(seen_count);
            }
            Value::Number(count as f64)
        }
        ArgValue::ReferenceUnion(ranges) => {
            if let Some(numeric) = numeric {
                let union_size = if blank_matches {
                    match reference_union_size(&ranges) {
                        Ok(v) => Some(v),
                        Err(e) => return Value::Error(e),
                    }
                } else {
                    None
                };
                let mut count: u64 = 0;
                let mut seen_count: u64 = 0;
                let mut seen = std::collections::HashSet::new();
                let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                let mut len = 0usize;

                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        seen_count += 1;
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        let Some(n) = coerce_candidate_to_number(&v, number_locale) else {
                            continue;
                        };
                        if n.is_nan() {
                            let tmp = Value::Number(n);
                            if criteria.matches(&tmp) {
                                count += 1;
                            }
                            continue;
                        }

                        buf[len] = n;
                        len += 1;
                        if len == SIMD_AGGREGATE_BLOCK {
                            count += simd::count_if_f64(&buf, numeric) as u64;
                            len = 0;
                        }
                    }
                }
                if len > 0 {
                    count += simd::count_if_f64(&buf[..len], numeric) as u64;
                }
                if let Some(union_size) = union_size {
                    count += union_size.saturating_sub(seen_count);
                }
                return Value::Number(count as f64);
            }

            let union_size = if blank_matches {
                match reference_union_size(&ranges) {
                    Ok(v) => Some(v),
                    Err(e) => return Value::Error(e),
                }
            } else {
                None
            };
            let mut count: u64 = 0;
            let mut seen_count: u64 = 0;
            let mut seen = std::collections::HashSet::new();
            for r in ranges {
                for addr in ctx.iter_reference_cells(&r) {
                    if !seen.insert((r.sheet_id.clone(), addr)) {
                        continue;
                    }
                    seen_count += 1;
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    if criteria.matches(&v) {
                        count += 1;
                    }
                }
            }
            if let Some(union_size) = union_size {
                count += union_size.saturating_sub(seen_count);
            }
            Value::Number(count as f64)
        }
        ArgValue::Scalar(Value::Array(arr)) => {
            if let Some(numeric) = numeric {
                let mut count: u64 = 0;
                let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
                let mut len = 0usize;

                for v in arr.iter() {
                    let Some(n) = coerce_candidate_to_number(v, number_locale) else {
                        continue;
                    };
                    if n.is_nan() {
                        let tmp = Value::Number(n);
                        if criteria.matches(&tmp) {
                            count += 1;
                        }
                        continue;
                    }

                    buf[len] = n;
                    len += 1;
                    if len == SIMD_AGGREGATE_BLOCK {
                        count += simd::count_if_f64(&buf, numeric) as u64;
                        len = 0;
                    }
                }

                if len > 0 {
                    count += simd::count_if_f64(&buf[..len], numeric) as u64;
                }

                return Value::Number(count as f64);
            }

            let mut count: u64 = 0;
            for v in arr.iter() {
                if criteria.matches(v) {
                    count += 1;
                }
            }
            Value::Number(count as f64)
        }
        ArgValue::Scalar(_) => Value::Error(ErrorKind::Value),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "COUNTIFS",
        min_args: 2,
        max_args: 254,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: countifs_fn,
    }
}

fn countifs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.len() < 2 || args.len() % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }

    #[derive(Debug, Clone)]
    enum CriteriaRange {
        Reference(crate::functions::Reference),
        Array(crate::value::Array),
    }

    impl CriteriaRange {
        fn shape(&self) -> (usize, usize) {
            match self {
                CriteriaRange::Reference(r) => {
                    let r = r.normalized();
                    (
                        (r.end.row - r.start.row + 1) as usize,
                        (r.end.col - r.start.col + 1) as usize,
                    )
                }
                CriteriaRange::Array(arr) => (arr.rows, arr.cols),
            }
        }

        fn record_reference(&self, ctx: &dyn FunctionContext) {
            if let CriteriaRange::Reference(r) = self {
                ctx.record_reference(r);
            }
        }

        fn value_at_offset(
            &self,
            ctx: &dyn FunctionContext,
            row_off: usize,
            col_off: usize,
        ) -> Value {
            match self {
                CriteriaRange::Reference(r) => {
                    let addr = CellAddr {
                        row: r.start.row + row_off as u32,
                        col: r.start.col + col_off as u32,
                    };
                    ctx.get_cell_value(&r.sheet_id, addr)
                }
                CriteriaRange::Array(arr) => {
                    arr.get(row_off, col_off).cloned().unwrap_or(Value::Blank)
                }
            }
        }

        fn value_at(&self, ctx: &dyn FunctionContext, idx: usize, cols: usize) -> Value {
            match self {
                CriteriaRange::Reference(r) => {
                    let row_off = idx / cols;
                    let col_off = idx % cols;
                    let addr = CellAddr {
                        row: r.start.row + row_off as u32,
                        col: r.start.col + col_off as u32,
                    };
                    ctx.get_cell_value(&r.sheet_id, addr)
                }
                CriteriaRange::Array(arr) => arr.values.get(idx).cloned().unwrap_or(Value::Blank),
            }
        }
    }

    let pair_count = args.len() / 2;
    let mut ranges: Vec<CriteriaRange> = Vec::new();
    let mut criteria: Vec<Criteria> = Vec::new();
    if ranges.try_reserve_exact(pair_count).is_err() || criteria.try_reserve_exact(pair_count).is_err()
    {
        debug_assert!(false, "allocation failed (countifs criteria buffers, pairs={pair_count})");
        return Value::Error(ErrorKind::Num);
    }
    let mut shape: Option<(usize, usize)> = None;
    let date_system = ctx.date_system();

    for pair in args.chunks_exact(2) {
        let criteria_range = match ctx.eval_arg(&pair[0]) {
            ArgValue::Reference(r) => CriteriaRange::Reference(r.normalized()),
            ArgValue::Scalar(Value::Array(arr)) => CriteriaRange::Array(arr),
            ArgValue::ReferenceUnion(_) | ArgValue::Scalar(_) => {
                return Value::Error(ErrorKind::Value)
            }
        };

        let (rows, cols) = criteria_range.shape();
        match shape {
            None => shape = Some((rows, cols)),
            Some(expected) if expected != (rows, cols) => return Value::Error(ErrorKind::Value),
            Some(_) => {}
        }
        criteria_range.record_reference(ctx);
        ranges.push(criteria_range);

        let criteria_value = eval_scalar_arg(ctx, &pair[1]);
        if let Value::Error(e) = criteria_value {
            return Value::Error(e);
        }
        let compiled = match Criteria::parse_with_date_system_and_locales(
            &criteria_value,
            date_system,
            ctx.value_locale(),
            ctx.now_utc(),
            ctx.locale_config(),
        ) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };
        criteria.push(compiled);
    }

    let (rows, cols) = shape.unwrap_or((0, 0));
    let len = rows.saturating_mul(cols);
    if len == 0 {
        return Value::Number(0.0);
    }

    let mut blank_matches: Vec<bool> = Vec::new();
    if blank_matches.try_reserve_exact(criteria.len()).is_err() {
        debug_assert!(
            false,
            "allocation failed (countifs blank_matches, len={})",
            criteria.len()
        );
        return Value::Error(ErrorKind::Num);
    }
    for crit in criteria.iter() {
        blank_matches.push(crit.matches(&Value::Blank));
    }

    // When all criteria match blank cells and all ranges are references, implicit blanks contribute
    // to the result. Avoid scanning the full shape by counting total cells minus the set of
    // explicitly-stored cells that fail any criteria.
    let all_blank_matches = blank_matches.iter().all(|matches_blank| *matches_blank);
    if all_blank_matches
        && ranges
            .iter()
            .all(|r| matches!(r, CriteriaRange::Reference(_)))
    {
        let total_cells = (rows as u64).saturating_mul(cols as u64);
        let mut mismatches: std::collections::HashSet<u64> = std::collections::HashSet::new();

        for (range, crit) in ranges.iter().zip(criteria.iter()) {
            let CriteriaRange::Reference(r) = range else {
                debug_assert!(false, "all ranges are references");
                return Value::Error(ErrorKind::Value);
            };
            for addr in ctx.iter_reference_cells(r) {
                let v = ctx.get_cell_value(&r.sheet_id, addr);
                if crit.matches(&v) {
                    continue;
                }

                let row_off = (addr.row - r.start.row) as u64;
                let col_off = (addr.col - r.start.col) as u64;
                let idx = row_off.saturating_mul(cols as u64).saturating_add(col_off);
                mismatches.insert(idx);
            }
        }

        return Value::Number(total_cells.saturating_sub(mismatches.len() as u64) as f64);
    }

    // If any criteria cannot match implicit blanks, we can iterate the matching cells of that
    // criteria range and avoid scanning the entire shape.
    let driver_idx = blank_matches
        .iter()
        .enumerate()
        .find_map(|(idx, matches_blank)| {
            if *matches_blank {
                return None;
            }
            matches!(&ranges[idx], CriteriaRange::Reference(_)).then_some(idx)
        });

    let mut count: u64 = 0;
    if let Some(driver_idx) = driver_idx {
        let CriteriaRange::Reference(driver_range) = &ranges[driver_idx] else {
            debug_assert!(false, "driver_idx always points at a reference range");
            return Value::Error(ErrorKind::Value);
        };
        let driver_crit = &criteria[driver_idx];

        'cell: for addr in ctx.iter_reference_cells(driver_range) {
            let driver_val = ctx.get_cell_value(&driver_range.sheet_id, addr);
            if !driver_crit.matches(&driver_val) {
                continue;
            }

            let row_off = (addr.row - driver_range.start.row) as usize;
            let col_off = (addr.col - driver_range.start.col) as usize;

            for (idx, (range, crit)) in ranges.iter().zip(criteria.iter()).enumerate() {
                if idx == driver_idx {
                    continue;
                }
                let v = range.value_at_offset(ctx, row_off, col_off);
                if !crit.matches(&v) {
                    continue 'cell;
                }
            }
            count += 1;
        }
    } else {
        'row: for idx in 0..len {
            for (range, crit) in ranges.iter().zip(criteria.iter()) {
                let v = range.value_at(ctx, idx, cols);
                if !crit.matches(&v) {
                    continue 'row;
                }
            }
            count += 1;
        }
    }

    Value::Number(count as f64)
}

inventory::submit! {
    FunctionSpec {
        name: "SUMIF",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: sumif_fn,
    }
}

fn sumif_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let criteria_range = match Range2D::try_from_arg(ctx.eval_arg(&args[0])) {
        Ok(r) => r,
        Err(e) => return Value::Error(e),
    };
    criteria_range.record_reference(ctx);

    let criteria_value = eval_scalar_arg(ctx, &args[1]);
    if let Value::Error(e) = criteria_value {
        return Value::Error(e);
    }
    let criteria = match Criteria::parse_with_date_system_and_locales(
        &criteria_value,
        ctx.date_system(),
        ctx.value_locale(),
        ctx.now_utc(),
        ctx.locale_config(),
    ) {
        Ok(c) => c,
        Err(e) => return Value::Error(e),
    };

    // Excel treats `SUMIF(range, criteria,)` the same as omitting the optional sum_range
    // argument entirely (i.e. `SUMIF(range, criteria)`).
    let sum_range = match args.get(2) {
        None | Some(CompiledExpr::Blank) => None,
        Some(expr) => Some(match Range2D::try_from_arg(ctx.eval_arg(expr)) {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        }),
    };
    if let Some(ref sum_range) = sum_range {
        sum_range.record_reference(ctx);
    }

    let (rows, cols) = criteria_range.shape();
    if let Some(ref sum_range) = sum_range {
        let (sum_rows, sum_cols) = sum_range.shape();
        if rows != sum_rows || cols != sum_cols {
            return Value::Error(ErrorKind::Value);
        }
    }

    // SIMD fast path: SUMIF over array literals / spilled arrays with numeric criteria.
    //
    // This intentionally only triggers when we can safely coerce every criteria-range entry to a
    // number (using COUNTIF-style coercion). That ensures we never accidentally treat a non-numeric
    // value as a blank/0 due to NaN normalization inside the SIMD kernel.
    if let (Some(numeric), Range2D::Array(criteria_arr)) =
        (criteria.as_numeric_criteria(), &criteria_range)
    {
        let len = criteria_arr.values.len();
        if len >= SIMD_CRITERIA_ARRAY_MIN_LEN {
            let locale = ctx.number_locale();
            match &sum_range {
                None => {
                    let mut sum = 0.0;
                    let mut can_simd = true;
                    let mut crit_buf: Vec<f64> = Vec::new();
                    let mut sum_buf: Vec<f64> = Vec::new();
                    if crit_buf.try_reserve_exact(SIMD_AGGREGATE_BLOCK).is_err()
                        || sum_buf.try_reserve_exact(SIMD_AGGREGATE_BLOCK).is_err()
                    {
                        debug_assert!(false, "allocation failed (sumif simd buffers)");
                        can_simd = false;
                    }

                    for v in criteria_arr.iter() {
                        let Some(n) = coerce_countif_value_to_number(v, locale) else {
                            can_simd = false;
                            break;
                        };
                        if n.is_nan() {
                            can_simd = false;
                            break;
                        }
                        crit_buf.push(n);

                        match v {
                            Value::Number(x) => {
                                if x.is_nan() {
                                    can_simd = false;
                                    break;
                                }
                                sum_buf.push(*x);
                            }
                            Value::Error(_) | Value::Lambda(_) => {
                                can_simd = false;
                                break;
                            }
                            _ => sum_buf.push(f64::NAN),
                        }

                        if crit_buf.len() == SIMD_AGGREGATE_BLOCK {
                            sum += simd::sum_if_f64(&sum_buf, &crit_buf, numeric);
                            crit_buf.clear();
                            sum_buf.clear();
                        }
                    }

                    if can_simd {
                        if !crit_buf.is_empty() {
                            sum += simd::sum_if_f64(&sum_buf, &crit_buf, numeric);
                        }
                        return Value::Number(sum);
                    }
                }
                Some(Range2D::Array(sum_arr)) => {
                    let mut sum = 0.0;
                    let mut can_simd = true;
                    let mut crit_buf: Vec<f64> = Vec::new();
                    let mut sum_buf: Vec<f64> = Vec::new();
                    if crit_buf.try_reserve_exact(SIMD_AGGREGATE_BLOCK).is_err()
                        || sum_buf.try_reserve_exact(SIMD_AGGREGATE_BLOCK).is_err()
                    {
                        debug_assert!(false, "allocation failed (sumif simd buffers)");
                        can_simd = false;
                    }

                    for (crit_v, sum_v) in criteria_arr.iter().zip(sum_arr.iter()) {
                        let Some(n) = coerce_countif_value_to_number(crit_v, locale) else {
                            can_simd = false;
                            break;
                        };
                        if n.is_nan() {
                            can_simd = false;
                            break;
                        }
                        crit_buf.push(n);

                        match sum_v {
                            Value::Number(x) => {
                                if x.is_nan() {
                                    can_simd = false;
                                    break;
                                }
                                sum_buf.push(*x);
                            }
                            Value::Error(_) | Value::Lambda(_) => {
                                // Errors in sum_range must be able to short-circuit when criteria
                                // matches, so fall back to scalar evaluation when present.
                                can_simd = false;
                                break;
                            }
                            _ => sum_buf.push(f64::NAN),
                        }

                        if crit_buf.len() == SIMD_AGGREGATE_BLOCK {
                            sum += simd::sum_if_f64(&sum_buf, &crit_buf, numeric);
                            crit_buf.clear();
                            sum_buf.clear();
                        }
                    }

                    if can_simd {
                        if !crit_buf.is_empty() {
                            sum += simd::sum_if_f64(&sum_buf, &crit_buf, numeric);
                        }
                        return Value::Number(sum);
                    }
                }
                Some(Range2D::Reference(_)) => {}
            }
        }
    }

    let mut sum = 0.0;
    match (&sum_range, &criteria_range) {
        (Some(Range2D::Reference(sum_ref)), _) => {
            for addr in ctx.iter_reference_cells(sum_ref) {
                let row = (addr.row - sum_ref.start.row) as usize;
                let col = (addr.col - sum_ref.start.col) as usize;
                let crit_val = criteria_range.get(ctx, row, col);
                if !criteria.matches(&crit_val) {
                    continue;
                }
                match ctx.get_cell_value(&sum_ref.sheet_id, addr) {
                    Value::Number(n) => sum += n,
                    Value::Error(e) => return Value::Error(e),
                    Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                    _ => {}
                }
            }
        }
        (None, Range2D::Reference(criteria_ref)) => {
            for addr in ctx.iter_reference_cells(criteria_ref) {
                let crit_val = ctx.get_cell_value(&criteria_ref.sheet_id, addr);
                if !criteria.matches(&crit_val) {
                    continue;
                }
                match crit_val {
                    Value::Number(n) => sum += n,
                    Value::Error(e) => return Value::Error(e),
                    Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                    _ => {}
                }
            }
        }
        _ => {
            for row in 0..rows {
                for col in 0..cols {
                    let crit_val = criteria_range.get(ctx, row, col);
                    if !criteria.matches(&crit_val) {
                        continue;
                    }

                    let sum_val = match &sum_range {
                        Some(r) => r.get(ctx, row, col),
                        None => crit_val,
                    };

                    match sum_val {
                        Value::Number(n) => sum += n,
                        Value::Error(e) => return Value::Error(e),
                        Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                        _ => {}
                    }
                }
            }
        }
    }

    Value::Number(sum)
}

inventory::submit! {
    FunctionSpec {
        name: "SUMIFS",
        min_args: 3,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: sumifs_fn,
    }
}

fn sumifs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.len() < 3 || (args.len() - 1) % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }

    let sum_range = match Range2D::try_from_arg(ctx.eval_arg(&args[0])) {
        Ok(r) => r,
        Err(e) => return Value::Error(e),
    };
    sum_range.record_reference(ctx);
    let (rows, cols) = sum_range.shape();

    let mut criteria_ranges = Vec::new();
    let mut criteria = Vec::new();

    for pair in args[1..].chunks_exact(2) {
        let range = match Range2D::try_from_arg(ctx.eval_arg(&pair[0])) {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        let (r_rows, r_cols) = range.shape();
        if r_rows != rows || r_cols != cols {
            return Value::Error(ErrorKind::Value);
        }

        let crit_value = eval_scalar_arg(ctx, &pair[1]);
        if let Value::Error(e) = crit_value {
            return Value::Error(e);
        }
        let crit = match Criteria::parse_with_date_system_and_locales(
            &crit_value,
            ctx.date_system(),
            ctx.value_locale(),
            ctx.now_utc(),
            ctx.locale_config(),
        ) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };

        criteria_ranges.push(range);
        criteria.push(crit);
    }

    for range in &criteria_ranges {
        range.record_reference(ctx);
    }

    let mut sum = 0.0;
    match &sum_range {
        Range2D::Reference(sum_ref) => {
            'cell: for addr in ctx.iter_reference_cells(sum_ref) {
                let row = (addr.row - sum_ref.start.row) as usize;
                let col = (addr.col - sum_ref.start.col) as usize;
                for (range, crit) in criteria_ranges.iter().zip(criteria.iter()) {
                    let v = range.get(ctx, row, col);
                    if !crit.matches(&v) {
                        continue 'cell;
                    }
                }

                match ctx.get_cell_value(&sum_ref.sheet_id, addr) {
                    Value::Number(n) => sum += n,
                    Value::Error(e) => return Value::Error(e),
                    Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                    _ => {}
                }
            }
        }
        _ => {
            for row in 0..rows {
                for col in 0..cols {
                    let mut matches = true;
                    for (range, crit) in criteria_ranges.iter().zip(criteria.iter()) {
                        let v = range.get(ctx, row, col);
                        if !crit.matches(&v) {
                            matches = false;
                            break;
                        }
                    }
                    if !matches {
                        continue;
                    }

                    match sum_range.get(ctx, row, col) {
                        Value::Number(n) => sum += n,
                        Value::Error(e) => return Value::Error(e),
                        Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                        _ => {}
                    }
                }
            }
        }
    }

    Value::Number(sum)
}

inventory::submit! {
    FunctionSpec {
        name: "AVERAGEIF",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: averageif_fn,
    }
}

fn averageif_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let criteria_range = match Range2D::try_from_arg(ctx.eval_arg(&args[0])) {
        Ok(r) => r,
        Err(e) => return Value::Error(e),
    };
    criteria_range.record_reference(ctx);

    let criteria_value = eval_scalar_arg(ctx, &args[1]);
    if let Value::Error(e) = criteria_value {
        return Value::Error(e);
    }
    let criteria = match Criteria::parse_with_date_system_and_locales(
        &criteria_value,
        ctx.date_system(),
        ctx.value_locale(),
        ctx.now_utc(),
        ctx.locale_config(),
    ) {
        Ok(c) => c,
        Err(e) => return Value::Error(e),
    };

    // Excel treats `AVERAGEIF(range, criteria,)` the same as omitting the optional average_range
    // argument entirely (i.e. `AVERAGEIF(range, criteria)`).
    let average_range = match args.get(2) {
        None | Some(CompiledExpr::Blank) => None,
        Some(expr) => Some(match Range2D::try_from_arg(ctx.eval_arg(expr)) {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        }),
    };
    if let Some(ref average_range) = average_range {
        average_range.record_reference(ctx);
    }

    let (rows, cols) = criteria_range.shape();
    if let Some(ref average_range) = average_range {
        let (avg_rows, avg_cols) = average_range.shape();
        if rows != avg_rows || cols != avg_cols {
            return Value::Error(ErrorKind::Value);
        }
    }

    // SIMD fast path: AVERAGEIF over array literals / spilled arrays with numeric criteria.
    //
    // See SUMIF for details on why we require every criteria-range entry to be numerically
    // coercible for this optimization.
    if let (Some(numeric), Range2D::Array(criteria_arr)) =
        (criteria.as_numeric_criteria(), &criteria_range)
    {
        let len = criteria_arr.values.len();
        if len >= SIMD_CRITERIA_ARRAY_MIN_LEN {
            let locale = ctx.number_locale();
            match &average_range {
                None => {
                    let mut sum = 0.0;
                    let mut count = 0u64;
                    let mut can_simd = true;
                    let mut crit_buf: Vec<f64> = Vec::new();
                    let mut avg_buf: Vec<f64> = Vec::new();
                    if crit_buf.try_reserve_exact(SIMD_AGGREGATE_BLOCK).is_err()
                        || avg_buf.try_reserve_exact(SIMD_AGGREGATE_BLOCK).is_err()
                    {
                        debug_assert!(false, "allocation failed (averageif simd buffers)");
                        can_simd = false;
                    }

                    for v in criteria_arr.iter() {
                        let Some(n) = coerce_countif_value_to_number(v, locale) else {
                            can_simd = false;
                            break;
                        };
                        if n.is_nan() {
                            can_simd = false;
                            break;
                        }
                        crit_buf.push(n);

                        match v {
                            Value::Number(x) => {
                                if x.is_nan() {
                                    can_simd = false;
                                    break;
                                }
                                avg_buf.push(*x);
                            }
                            Value::Error(_) | Value::Lambda(_) => {
                                can_simd = false;
                                break;
                            }
                            _ => avg_buf.push(f64::NAN),
                        }

                        if crit_buf.len() == SIMD_AGGREGATE_BLOCK {
                            let (s, c) = simd::sum_count_if_f64(&avg_buf, &crit_buf, numeric);
                            sum += s;
                            count += c as u64;
                            crit_buf.clear();
                            avg_buf.clear();
                        }
                    }

                    if can_simd {
                        if !crit_buf.is_empty() {
                            let (s, c) = simd::sum_count_if_f64(&avg_buf, &crit_buf, numeric);
                            sum += s;
                            count += c as u64;
                        }
                        if count == 0 {
                            return Value::Error(ErrorKind::Div0);
                        }
                        return Value::Number(sum / count as f64);
                    }
                }
                Some(Range2D::Array(avg_arr)) => {
                    let mut sum = 0.0;
                    let mut count = 0u64;
                    let mut can_simd = true;
                    let mut crit_buf: Vec<f64> = Vec::new();
                    let mut avg_buf: Vec<f64> = Vec::new();
                    if crit_buf.try_reserve_exact(SIMD_AGGREGATE_BLOCK).is_err()
                        || avg_buf.try_reserve_exact(SIMD_AGGREGATE_BLOCK).is_err()
                    {
                        debug_assert!(false, "allocation failed (averageif simd buffers)");
                        can_simd = false;
                    }

                    for (crit_v, avg_v) in criteria_arr.iter().zip(avg_arr.iter()) {
                        let Some(n) = coerce_countif_value_to_number(crit_v, locale) else {
                            can_simd = false;
                            break;
                        };
                        if n.is_nan() {
                            can_simd = false;
                            break;
                        }
                        crit_buf.push(n);

                        match avg_v {
                            Value::Number(x) => {
                                if x.is_nan() {
                                    can_simd = false;
                                    break;
                                }
                                avg_buf.push(*x);
                            }
                            Value::Error(_) | Value::Lambda(_) => {
                                can_simd = false;
                                break;
                            }
                            _ => avg_buf.push(f64::NAN),
                        }

                        if crit_buf.len() == SIMD_AGGREGATE_BLOCK {
                            let (s, c) = simd::sum_count_if_f64(&avg_buf, &crit_buf, numeric);
                            sum += s;
                            count += c as u64;
                            crit_buf.clear();
                            avg_buf.clear();
                        }
                    }

                    if can_simd {
                        if !crit_buf.is_empty() {
                            let (s, c) = simd::sum_count_if_f64(&avg_buf, &crit_buf, numeric);
                            sum += s;
                            count += c as u64;
                        }
                        if count == 0 {
                            return Value::Error(ErrorKind::Div0);
                        }
                        return Value::Number(sum / count as f64);
                    }
                }
                Some(Range2D::Reference(_)) => {}
            }
        }
    }

    let mut sum = 0.0;
    let mut count = 0u64;
    match (&average_range, &criteria_range) {
        (Some(Range2D::Reference(avg_ref)), _) => {
            for addr in ctx.iter_reference_cells(avg_ref) {
                let row = (addr.row - avg_ref.start.row) as usize;
                let col = (addr.col - avg_ref.start.col) as usize;
                let crit_val = criteria_range.get(ctx, row, col);
                if !criteria.matches(&crit_val) {
                    continue;
                }
                match ctx.get_cell_value(&avg_ref.sheet_id, addr) {
                    Value::Number(n) => {
                        sum += n;
                        count += 1;
                    }
                    Value::Error(e) => return Value::Error(e),
                    Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                    _ => {}
                }
            }
        }
        (None, Range2D::Reference(criteria_ref)) => {
            for addr in ctx.iter_reference_cells(criteria_ref) {
                let crit_val = ctx.get_cell_value(&criteria_ref.sheet_id, addr);
                if !criteria.matches(&crit_val) {
                    continue;
                }
                match crit_val {
                    Value::Number(n) => {
                        sum += n;
                        count += 1;
                    }
                    Value::Error(e) => return Value::Error(e),
                    Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                    _ => {}
                }
            }
        }
        _ => {
            for row in 0..rows {
                for col in 0..cols {
                    let crit_val = criteria_range.get(ctx, row, col);
                    if !criteria.matches(&crit_val) {
                        continue;
                    }

                    let avg_val = match &average_range {
                        Some(r) => r.get(ctx, row, col),
                        None => crit_val,
                    };

                    match avg_val {
                        Value::Number(n) => {
                            sum += n;
                            count += 1;
                        }
                        Value::Error(e) => return Value::Error(e),
                        Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                        _ => {}
                    }
                }
            }
        }
    }

    if count == 0 {
        return Value::Error(ErrorKind::Div0);
    }
    Value::Number(sum / count as f64)
}

inventory::submit! {
    FunctionSpec {
        name: "AVERAGEIFS",
        min_args: 3,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: averageifs_fn,
    }
}

fn averageifs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.len() < 3 || (args.len() - 1) % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }

    let average_range = match Range2D::try_from_arg(ctx.eval_arg(&args[0])) {
        Ok(r) => r,
        Err(e) => return Value::Error(e),
    };
    average_range.record_reference(ctx);
    let (rows, cols) = average_range.shape();

    let mut criteria_ranges = Vec::new();
    let mut criteria = Vec::new();

    for pair in args[1..].chunks_exact(2) {
        let range = match Range2D::try_from_arg(ctx.eval_arg(&pair[0])) {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        let (r_rows, r_cols) = range.shape();
        if r_rows != rows || r_cols != cols {
            return Value::Error(ErrorKind::Value);
        }

        let crit_value = eval_scalar_arg(ctx, &pair[1]);
        if let Value::Error(e) = crit_value {
            return Value::Error(e);
        }
        let crit = match Criteria::parse_with_date_system_and_locales(
            &crit_value,
            ctx.date_system(),
            ctx.value_locale(),
            ctx.now_utc(),
            ctx.locale_config(),
        ) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };

        criteria_ranges.push(range);
        criteria.push(crit);
    }

    for range in &criteria_ranges {
        range.record_reference(ctx);
    }

    let mut sum = 0.0;
    let mut count = 0u64;
    match &average_range {
        Range2D::Reference(avg_ref) => {
            'cell: for addr in ctx.iter_reference_cells(avg_ref) {
                let row = (addr.row - avg_ref.start.row) as usize;
                let col = (addr.col - avg_ref.start.col) as usize;
                for (range, crit) in criteria_ranges.iter().zip(criteria.iter()) {
                    let v = range.get(ctx, row, col);
                    if !crit.matches(&v) {
                        continue 'cell;
                    }
                }

                match ctx.get_cell_value(&avg_ref.sheet_id, addr) {
                    Value::Number(n) => {
                        sum += n;
                        count += 1;
                    }
                    Value::Error(e) => return Value::Error(e),
                    Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                    _ => {}
                }
            }
        }
        _ => {
            for row in 0..rows {
                for col in 0..cols {
                    let mut matches = true;
                    for (range, crit) in criteria_ranges.iter().zip(criteria.iter()) {
                        let v = range.get(ctx, row, col);
                        if !crit.matches(&v) {
                            matches = false;
                            break;
                        }
                    }
                    if !matches {
                        continue;
                    }

                    match average_range.get(ctx, row, col) {
                        Value::Number(n) => {
                            sum += n;
                            count += 1;
                        }
                        Value::Error(e) => return Value::Error(e),
                        Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                        _ => {}
                    }
                }
            }
        }
    }

    if count == 0 {
        return Value::Error(ErrorKind::Div0);
    }
    Value::Number(sum / count as f64)
}

inventory::submit! {
    FunctionSpec {
        name: "MAXIFS",
        min_args: 3,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: maxifs_fn,
    }
}

fn maxifs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.len() < 3 || (args.len() - 1) % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }

    let max_range = match Range2D::try_from_arg(ctx.eval_arg(&args[0])) {
        Ok(r) => r,
        Err(e) => return Value::Error(e),
    };
    max_range.record_reference(ctx);
    let (rows, cols) = max_range.shape();

    let mut criteria_ranges = Vec::new();
    let mut criteria = Vec::new();

    for pair in args[1..].chunks_exact(2) {
        let range = match Range2D::try_from_arg(ctx.eval_arg(&pair[0])) {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        let (r_rows, r_cols) = range.shape();
        if r_rows != rows || r_cols != cols {
            return Value::Error(ErrorKind::Value);
        }

        let crit_value = eval_scalar_arg(ctx, &pair[1]);
        if let Value::Error(e) = crit_value {
            return Value::Error(e);
        }
        let crit = match Criteria::parse_with_date_system_and_locales(
            &crit_value,
            ctx.date_system(),
            ctx.value_locale(),
            ctx.now_utc(),
            ctx.locale_config(),
        ) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };

        criteria_ranges.push(range);
        criteria.push(crit);
    }

    for range in &criteria_ranges {
        range.record_reference(ctx);
    }

    let mut best: Option<f64> = None;
    let mut earliest_error: Option<(usize, usize, ErrorKind)> = None;
    match &max_range {
        Range2D::Reference(max_ref) => {
            // `iter_reference_cells` is sparse when the backend supports it. It is not
            // guaranteed to yield cells in row-major order, so we track the earliest included
            // error explicitly (row-major within the range) to match Excel semantics without
            // allocating/sorting the entire range.
            'cell: for addr in ctx.iter_reference_cells(max_ref) {
                let value = ctx.get_cell_value(&max_ref.sheet_id, addr);
                let (row, col) = match value {
                    // Only numeric values contribute; errors propagate only when the row is included.
                    Value::Number(_) | Value::Error(_) | Value::Lambda(_) => (
                        (addr.row - max_ref.start.row) as usize,
                        (addr.col - max_ref.start.col) as usize,
                    ),
                    _ => continue,
                };

                for (range, crit) in criteria_ranges.iter().zip(criteria.iter()) {
                    let v = range.get(ctx, row, col);
                    if !crit.matches(&v) {
                        continue 'cell;
                    }
                }

                match value {
                    Value::Number(n) => best = Some(best.map_or(n, |b| b.max(n))),
                    Value::Error(e) => {
                        if row == 0 && col == 0 {
                            return Value::Error(e);
                        }
                        match earliest_error {
                            None => earliest_error = Some((row, col, e)),
                            Some((best_row, best_col, _)) => {
                                if (row, col) < (best_row, best_col) {
                                    earliest_error = Some((row, col, e));
                                }
                            }
                        }
                    }
                    Value::Lambda(_) => {
                        if row == 0 && col == 0 {
                            return Value::Error(ErrorKind::Value);
                        }
                        match earliest_error {
                            None => earliest_error = Some((row, col, ErrorKind::Value)),
                            Some((best_row, best_col, _)) => {
                                if (row, col) < (best_row, best_col) {
                                    earliest_error = Some((row, col, ErrorKind::Value));
                                }
                            }
                        }
                    }
                    _ => {
                        debug_assert!(false, "filtered above");
                        continue;
                    }
                }
            }
        }
        Range2D::Array(_) => {
            for row in 0..rows {
                'col: for col in 0..cols {
                    let value = max_range.get(ctx, row, col);
                    match value {
                        Value::Number(_) | Value::Error(_) | Value::Lambda(_) => {}
                        _ => continue,
                    }

                    for (range, crit) in criteria_ranges.iter().zip(criteria.iter()) {
                        let v = range.get(ctx, row, col);
                        if !crit.matches(&v) {
                            continue 'col;
                        }
                    }

                    match value {
                        Value::Number(n) => best = Some(best.map_or(n, |b| b.max(n))),
                        Value::Error(e) => {
                            if row == 0 && col == 0 {
                                return Value::Error(e);
                            }
                            match earliest_error {
                                None => earliest_error = Some((row, col, e)),
                                Some((best_row, best_col, _)) => {
                                    if (row, col) < (best_row, best_col) {
                                        earliest_error = Some((row, col, e));
                                    }
                                }
                            }
                        }
                        Value::Lambda(_) => {
                            if row == 0 && col == 0 {
                                return Value::Error(ErrorKind::Value);
                            }
                            match earliest_error {
                                None => earliest_error = Some((row, col, ErrorKind::Value)),
                                Some((best_row, best_col, _)) => {
                                    if (row, col) < (best_row, best_col) {
                                        earliest_error = Some((row, col, ErrorKind::Value));
                                    }
                                }
                            }
                        }
                        _ => {
                            debug_assert!(false, "filtered above");
                            continue;
                        }
                    }
                }
            }
        }
    }

    if let Some((_, _, err)) = earliest_error {
        return Value::Error(err);
    }
    Value::Number(best.unwrap_or(0.0))
}

inventory::submit! {
    FunctionSpec {
        name: "MINIFS",
        min_args: 3,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: minifs_fn,
    }
}

fn minifs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.len() < 3 || (args.len() - 1) % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }

    let min_range = match Range2D::try_from_arg(ctx.eval_arg(&args[0])) {
        Ok(r) => r,
        Err(e) => return Value::Error(e),
    };
    min_range.record_reference(ctx);
    let (rows, cols) = min_range.shape();

    let mut criteria_ranges = Vec::new();
    let mut criteria = Vec::new();

    for pair in args[1..].chunks_exact(2) {
        let range = match Range2D::try_from_arg(ctx.eval_arg(&pair[0])) {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        let (r_rows, r_cols) = range.shape();
        if r_rows != rows || r_cols != cols {
            return Value::Error(ErrorKind::Value);
        }

        let crit_value = eval_scalar_arg(ctx, &pair[1]);
        if let Value::Error(e) = crit_value {
            return Value::Error(e);
        }
        let crit = match Criteria::parse_with_date_system_and_locales(
            &crit_value,
            ctx.date_system(),
            ctx.value_locale(),
            ctx.now_utc(),
            ctx.locale_config(),
        ) {
            Ok(c) => c,
            Err(e) => return Value::Error(e),
        };

        criteria_ranges.push(range);
        criteria.push(crit);
    }

    for range in &criteria_ranges {
        range.record_reference(ctx);
    }

    let mut best: Option<f64> = None;
    let mut earliest_error: Option<(usize, usize, ErrorKind)> = None;
    match &min_range {
        Range2D::Reference(min_ref) => {
            // See MAXIFS for why we track errors explicitly (stable row-major propagation without
            // needing to allocate/sort dense address lists).
            'cell: for addr in ctx.iter_reference_cells(min_ref) {
                let value = ctx.get_cell_value(&min_ref.sheet_id, addr);
                let (row, col) = match value {
                    Value::Number(_) | Value::Error(_) | Value::Lambda(_) => (
                        (addr.row - min_ref.start.row) as usize,
                        (addr.col - min_ref.start.col) as usize,
                    ),
                    _ => continue,
                };

                for (range, crit) in criteria_ranges.iter().zip(criteria.iter()) {
                    let v = range.get(ctx, row, col);
                    if !crit.matches(&v) {
                        continue 'cell;
                    }
                }

                match value {
                    Value::Number(n) => best = Some(best.map_or(n, |b| b.min(n))),
                    Value::Error(e) => {
                        if row == 0 && col == 0 {
                            return Value::Error(e);
                        }
                        match earliest_error {
                            None => earliest_error = Some((row, col, e)),
                            Some((best_row, best_col, _)) => {
                                if (row, col) < (best_row, best_col) {
                                    earliest_error = Some((row, col, e));
                                }
                            }
                        }
                    }
                    Value::Lambda(_) => {
                        if row == 0 && col == 0 {
                            return Value::Error(ErrorKind::Value);
                        }
                        match earliest_error {
                            None => earliest_error = Some((row, col, ErrorKind::Value)),
                            Some((best_row, best_col, _)) => {
                                if (row, col) < (best_row, best_col) {
                                    earliest_error = Some((row, col, ErrorKind::Value));
                                }
                            }
                        }
                    }
                    _ => {
                        debug_assert!(false, "filtered above");
                        continue;
                    }
                }
            }
        }
        Range2D::Array(_) => {
            for row in 0..rows {
                'col: for col in 0..cols {
                    let value = min_range.get(ctx, row, col);
                    match value {
                        Value::Number(_) | Value::Error(_) | Value::Lambda(_) => {}
                        _ => continue,
                    }

                    for (range, crit) in criteria_ranges.iter().zip(criteria.iter()) {
                        let v = range.get(ctx, row, col);
                        if !crit.matches(&v) {
                            continue 'col;
                        }
                    }

                    match value {
                        Value::Number(n) => best = Some(best.map_or(n, |b| b.min(n))),
                        Value::Error(e) => {
                            if row == 0 && col == 0 {
                                return Value::Error(e);
                            }
                            match earliest_error {
                                None => earliest_error = Some((row, col, e)),
                                Some((best_row, best_col, _)) => {
                                    if (row, col) < (best_row, best_col) {
                                        earliest_error = Some((row, col, e));
                                    }
                                }
                            }
                        }
                        Value::Lambda(_) => {
                            if row == 0 && col == 0 {
                                return Value::Error(ErrorKind::Value);
                            }
                            match earliest_error {
                                None => earliest_error = Some((row, col, ErrorKind::Value)),
                                Some((best_row, best_col, _)) => {
                                    if (row, col) < (best_row, best_col) {
                                        earliest_error = Some((row, col, ErrorKind::Value));
                                    }
                                }
                            }
                        }
                        _ => {
                            debug_assert!(false, "filtered above");
                            continue;
                        }
                    }
                }
            }
        }
    }

    if let Some((_, _, err)) = earliest_error {
        return Value::Error(err);
    }
    Value::Number(best.unwrap_or(0.0))
}

#[derive(Clone)]
enum Range2D {
    Reference(crate::functions::Reference),
    Array(Array),
}

impl Range2D {
    fn try_from_arg(arg: ArgValue) -> Result<Self, ErrorKind> {
        match arg {
            ArgValue::Reference(r) => Ok(Self::Reference(r.normalized())),
            ArgValue::Scalar(Value::Array(arr)) => Ok(Self::Array(arr)),
            ArgValue::Scalar(Value::Reference(r)) => Ok(Self::Reference(r.normalized())),
            ArgValue::Scalar(Value::Error(e)) => Err(e),
            ArgValue::ReferenceUnion(_) | ArgValue::Scalar(_) => Err(ErrorKind::Value),
        }
    }

    fn record_reference(&self, ctx: &dyn FunctionContext) {
        if let Range2D::Reference(r) = self {
            ctx.record_reference(r);
        }
    }

    fn shape(&self) -> (usize, usize) {
        match self {
            Range2D::Reference(r) => {
                let rows = (r.end.row - r.start.row + 1) as usize;
                let cols = (r.end.col - r.start.col + 1) as usize;
                (rows, cols)
            }
            Range2D::Array(arr) => (arr.rows, arr.cols),
        }
    }

    fn get(&self, ctx: &dyn FunctionContext, row: usize, col: usize) -> Value {
        match self {
            Range2D::Reference(r) => {
                let addr = crate::eval::CellAddr {
                    row: r.start.row + row as u32,
                    col: r.start.col + col as u32,
                };
                ctx.get_cell_value(&r.sheet_id, addr)
            }
            Range2D::Array(arr) => arr.get(row, col).cloned().unwrap_or(Value::Blank),
        }
    }
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
    enum SumproductOperand {
        Scalar(Value),
        Array(Vec<Value>),
        Reference(crate::functions::Reference),
    }

    impl SumproductOperand {
        fn len(&self) -> usize {
            match self {
                Self::Scalar(_) => 1,
                Self::Array(values) => values.len(),
                Self::Reference(r) => {
                    let rows = r.end.row - r.start.row + 1;
                    let cols = r.end.col - r.start.col + 1;
                    (rows as usize).saturating_mul(cols as usize)
                }
            }
        }
    }

    fn arg_to_operand(
        ctx: &dyn FunctionContext,
        arg: ArgValue,
    ) -> Result<SumproductOperand, Value> {
        match arg {
            ArgValue::Reference(r) => {
                let r = r.normalized();
                ctx.record_reference(&r);
                Ok(SumproductOperand::Reference(r))
            }
            ArgValue::ReferenceUnion(_) => Err(Value::Error(ErrorKind::Value)),
            ArgValue::Scalar(Value::Array(arr)) => Ok(SumproductOperand::Array(arr.values)),
            ArgValue::Scalar(Value::Error(e)) => Err(Value::Error(e)),
            ArgValue::Scalar(v) => Ok(SumproductOperand::Scalar(v)),
        }
    }

    // Match Excel-style argument evaluation: process arguments in order so we still record
    // precedents for earlier reference arguments even if later args are errors.
    let a = match arg_to_operand(ctx, ctx.eval_arg(&args[0])) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let b = match arg_to_operand(ctx, ctx.eval_arg(&args[1])) {
        Ok(v) => v,
        Err(e) => return e,
    };

    let len_a = a.len();
    let len_b = b.len();
    if len_a == 0 || len_b == 0 {
        return Value::Error(ErrorKind::Value);
    }
    let len = len_a.max(len_b);
    if (len_a != len && len_a != 1) || (len_b != len && len_b != 1) {
        return Value::Error(ErrorKind::Value);
    }

    let locale = ctx.number_locale();

    // Non-reference inputs can use the shared math helper (already SIMD optimized) without extra
    // allocations.
    if !matches!(a, SumproductOperand::Reference(_))
        && !matches!(b, SumproductOperand::Reference(_))
    {
        let slice_a: &[Value] = match &a {
            SumproductOperand::Scalar(v) => std::slice::from_ref(v),
            SumproductOperand::Array(values) => values,
            SumproductOperand::Reference(_) => {
                debug_assert!(false, "non-reference branch received reference operand");
                return Value::Error(ErrorKind::Value);
            }
        };
        let slice_b: &[Value] = match &b {
            SumproductOperand::Scalar(v) => std::slice::from_ref(v),
            SumproductOperand::Array(values) => values,
            SumproductOperand::Reference(_) => {
                debug_assert!(false, "non-reference branch received reference operand");
                return Value::Error(ErrorKind::Value);
            }
        };

        let arrays: [&[Value]; 2] = [slice_a, slice_b];
        return match crate::functions::math::sumproduct(&arrays, locale) {
            Ok(v) => Value::Number(v),
            Err(e) => Value::Error(e),
        };
    }

    // Reference inputs can be extremely large (e.g. `A:A`), so avoid materializing `Vec<Value>`
    // and stream-coerce values into SIMD buffers.
    const BLOCK: usize = SIMD_AGGREGATE_BLOCK;

    fn flush(sum: &mut f64, buf_a: &[f64], buf_b: &[f64]) {
        *sum += simd::sumproduct_ignore_nan_f64(buf_a, buf_b);
    }

    let result = (|| -> Result<f64, ErrorKind> {
        let mut buf_a = [0.0_f64; BLOCK];
        let mut buf_b = [0.0_f64; BLOCK];
        let mut buf_len = 0usize;

        let mut sum = 0.0;
        let mut saw_nan = false;

        match (a, b) {
            (SumproductOperand::Reference(ra), SumproductOperand::Reference(rb)) => {
                if len_a == len && len_b == len {
                    for (addr_a, addr_b) in ra.iter_cells().zip(rb.iter_cells()) {
                        let va = ctx.get_cell_value(&ra.sheet_id, addr_a);
                        let xa = crate::functions::math::coerce_sumproduct_number(&va, locale)?;
                        let vb = ctx.get_cell_value(&rb.sheet_id, addr_b);
                        let xb = crate::functions::math::coerce_sumproduct_number(&vb, locale)?;

                        if xa.is_nan() || xb.is_nan() {
                            saw_nan = true;
                        }

                        if !saw_nan {
                            buf_a[buf_len] = xa;
                            buf_b[buf_len] = xb;
                            buf_len += 1;
                            if buf_len == BLOCK {
                                flush(&mut sum, &buf_a, &buf_b);
                                buf_len = 0;
                            }
                        }
                    }
                } else if len_a == 1 && len_b == len {
                    let Some(addr_a0) = ra.iter_cells().next() else {
                        debug_assert!(false, "len_a validated > 0 but iter_cells was empty");
                        return Err(ErrorKind::Value);
                    };
                    let va0 = ctx.get_cell_value(&ra.sheet_id, addr_a0);
                    let xa = crate::functions::math::coerce_sumproduct_number(&va0, locale)?;
                    saw_nan = xa.is_nan();

                    for addr_b in rb.iter_cells() {
                        let vb = ctx.get_cell_value(&rb.sheet_id, addr_b);
                        let xb = crate::functions::math::coerce_sumproduct_number(&vb, locale)?;

                        if saw_nan {
                            continue;
                        }
                        if xb.is_nan() {
                            saw_nan = true;
                            continue;
                        }

                        buf_a[buf_len] = xa;
                        buf_b[buf_len] = xb;
                        buf_len += 1;
                        if buf_len == BLOCK {
                            flush(&mut sum, &buf_a, &buf_b);
                            buf_len = 0;
                        }
                    }
                } else if len_b == 1 && len_a == len {
                    let mut iter_a = ra.iter_cells();
                    let Some(addr_a0) = iter_a.next() else {
                        debug_assert!(false, "len_a validated > 0 but iter_cells was empty");
                        return Err(ErrorKind::Value);
                    };
                    let Some(addr_b0) = rb.iter_cells().next() else {
                        debug_assert!(false, "len_b validated > 0 but iter_cells was empty");
                        return Err(ErrorKind::Value);
                    };

                    let va0 = ctx.get_cell_value(&ra.sheet_id, addr_a0);
                    let xa0 = crate::functions::math::coerce_sumproduct_number(&va0, locale)?;
                    let vb0 = ctx.get_cell_value(&rb.sheet_id, addr_b0);
                    let xb = crate::functions::math::coerce_sumproduct_number(&vb0, locale)?;

                    saw_nan = xa0.is_nan() || xb.is_nan();
                    if !saw_nan {
                        buf_a[buf_len] = xa0;
                        buf_b[buf_len] = xb;
                        buf_len += 1;
                    }

                    for addr_a in iter_a {
                        let va = ctx.get_cell_value(&ra.sheet_id, addr_a);
                        let xa = crate::functions::math::coerce_sumproduct_number(&va, locale)?;

                        if saw_nan {
                            continue;
                        }
                        if xa.is_nan() {
                            saw_nan = true;
                            continue;
                        }

                        buf_a[buf_len] = xa;
                        buf_b[buf_len] = xb;
                        buf_len += 1;
                        if buf_len == BLOCK {
                            flush(&mut sum, &buf_a, &buf_b);
                            buf_len = 0;
                        }
                    }
                } else {
                    debug_assert!(
                        false,
                        "broadcast validation should have handled all length combinations"
                    );
                    return Err(ErrorKind::Value);
                }
            }
            (SumproductOperand::Reference(ra), SumproductOperand::Array(vb)) => {
                if len_a == len && len_b == len {
                    for (idx, addr_a) in ra.iter_cells().enumerate() {
                        let va = ctx.get_cell_value(&ra.sheet_id, addr_a);
                        let xa = crate::functions::math::coerce_sumproduct_number(&va, locale)?;
                        let xb =
                            crate::functions::math::coerce_sumproduct_number(&vb[idx], locale)?;

                        if xa.is_nan() || xb.is_nan() {
                            saw_nan = true;
                        }
                        if !saw_nan {
                            buf_a[buf_len] = xa;
                            buf_b[buf_len] = xb;
                            buf_len += 1;
                            if buf_len == BLOCK {
                                flush(&mut sum, &buf_a, &buf_b);
                                buf_len = 0;
                            }
                        }
                    }
                } else if len_a == 1 && len_b == len {
                    let Some(addr_a0) = ra.iter_cells().next() else {
                        debug_assert!(false, "len_a validated > 0 but iter_cells was empty");
                        return Err(ErrorKind::Value);
                    };
                    let va0 = ctx.get_cell_value(&ra.sheet_id, addr_a0);
                    let xa = crate::functions::math::coerce_sumproduct_number(&va0, locale)?;
                    saw_nan = xa.is_nan();

                    for vb in &vb {
                        let xb = crate::functions::math::coerce_sumproduct_number(vb, locale)?;
                        if saw_nan {
                            continue;
                        }
                        if xb.is_nan() {
                            saw_nan = true;
                            continue;
                        }

                        buf_a[buf_len] = xa;
                        buf_b[buf_len] = xb;
                        buf_len += 1;
                        if buf_len == BLOCK {
                            flush(&mut sum, &buf_a, &buf_b);
                            buf_len = 0;
                        }
                    }
                } else if len_b == 1 && len_a == len {
                    let mut iter_a = ra.iter_cells();
                    let Some(addr_a0) = iter_a.next() else {
                        debug_assert!(false, "len_a validated > 0 but iter_cells was empty");
                        return Err(ErrorKind::Value);
                    };
                    let va0 = ctx.get_cell_value(&ra.sheet_id, addr_a0);
                    let xa0 = crate::functions::math::coerce_sumproduct_number(&va0, locale)?;
                    let xb = crate::functions::math::coerce_sumproduct_number(&vb[0], locale)?;

                    saw_nan = xa0.is_nan() || xb.is_nan();
                    if !saw_nan {
                        buf_a[buf_len] = xa0;
                        buf_b[buf_len] = xb;
                        buf_len += 1;
                    }

                    for addr_a in iter_a {
                        let va = ctx.get_cell_value(&ra.sheet_id, addr_a);
                        let xa = crate::functions::math::coerce_sumproduct_number(&va, locale)?;

                        if saw_nan {
                            continue;
                        }
                        if xa.is_nan() {
                            saw_nan = true;
                            continue;
                        }

                        buf_a[buf_len] = xa;
                        buf_b[buf_len] = xb;
                        buf_len += 1;
                        if buf_len == BLOCK {
                            flush(&mut sum, &buf_a, &buf_b);
                            buf_len = 0;
                        }
                    }
                } else {
                    debug_assert!(
                        false,
                        "broadcast validation should have handled all length combinations"
                    );
                    return Err(ErrorKind::Value);
                }
            }
            (SumproductOperand::Reference(ra), SumproductOperand::Scalar(vb)) => {
                let mut iter_a = ra.iter_cells();
                let Some(addr_a0) = iter_a.next() else {
                    debug_assert!(false, "len_a validated > 0 but iter_cells was empty");
                    return Err(ErrorKind::Value);
                };
                let va0 = ctx.get_cell_value(&ra.sheet_id, addr_a0);
                let xa0 = crate::functions::math::coerce_sumproduct_number(&va0, locale)?;
                let xb = crate::functions::math::coerce_sumproduct_number(&vb, locale)?;

                saw_nan = xa0.is_nan() || xb.is_nan();
                if !saw_nan {
                    buf_a[buf_len] = xa0;
                    buf_b[buf_len] = xb;
                    buf_len += 1;
                }

                for addr_a in iter_a {
                    let va = ctx.get_cell_value(&ra.sheet_id, addr_a);
                    let xa = crate::functions::math::coerce_sumproduct_number(&va, locale)?;

                    if saw_nan {
                        continue;
                    }
                    if xa.is_nan() {
                        saw_nan = true;
                        continue;
                    }

                    buf_a[buf_len] = xa;
                    buf_b[buf_len] = xb;
                    buf_len += 1;
                    if buf_len == BLOCK {
                        flush(&mut sum, &buf_a, &buf_b);
                        buf_len = 0;
                    }
                }
            }
            (SumproductOperand::Array(va), SumproductOperand::Reference(rb)) => {
                if len_a == len && len_b == len {
                    for (idx, addr_b) in rb.iter_cells().enumerate() {
                        let xa =
                            crate::functions::math::coerce_sumproduct_number(&va[idx], locale)?;
                        let vb = ctx.get_cell_value(&rb.sheet_id, addr_b);
                        let xb = crate::functions::math::coerce_sumproduct_number(&vb, locale)?;

                        if xa.is_nan() || xb.is_nan() {
                            saw_nan = true;
                        }
                        if !saw_nan {
                            buf_a[buf_len] = xa;
                            buf_b[buf_len] = xb;
                            buf_len += 1;
                            if buf_len == BLOCK {
                                flush(&mut sum, &buf_a, &buf_b);
                                buf_len = 0;
                            }
                        }
                    }
                } else if len_a == 1 && len_b == len {
                    let xa = crate::functions::math::coerce_sumproduct_number(&va[0], locale)?;
                    saw_nan = xa.is_nan();

                    for addr_b in rb.iter_cells() {
                        let vb = ctx.get_cell_value(&rb.sheet_id, addr_b);
                        let xb = crate::functions::math::coerce_sumproduct_number(&vb, locale)?;

                        if saw_nan {
                            continue;
                        }
                        if xb.is_nan() {
                            saw_nan = true;
                            continue;
                        }

                        buf_a[buf_len] = xa;
                        buf_b[buf_len] = xb;
                        buf_len += 1;
                        if buf_len == BLOCK {
                            flush(&mut sum, &buf_a, &buf_b);
                            buf_len = 0;
                        }
                    }
                } else if len_b == 1 && len_a == len {
                    // Preserve error precedence: for idx=0 we must coerce `va[0]` before the scalar.
                    let mut iter_b = rb.iter_cells();
                    let Some(addr_b0) = iter_b.next() else {
                        debug_assert!(false, "len_b validated > 0 but iter_cells was empty");
                        return Err(ErrorKind::Value);
                    };
                    let xa0 = crate::functions::math::coerce_sumproduct_number(&va[0], locale)?;
                    let vb0 = ctx.get_cell_value(&rb.sheet_id, addr_b0);
                    let xb = crate::functions::math::coerce_sumproduct_number(&vb0, locale)?;

                    saw_nan = xa0.is_nan() || xb.is_nan();
                    if !saw_nan {
                        buf_a[buf_len] = xa0;
                        buf_b[buf_len] = xb;
                        buf_len += 1;
                    }

                    for va in va.into_iter().skip(1) {
                        let xa = crate::functions::math::coerce_sumproduct_number(&va, locale)?;

                        if saw_nan {
                            continue;
                        }
                        if xa.is_nan() {
                            saw_nan = true;
                            continue;
                        }

                        buf_a[buf_len] = xa;
                        buf_b[buf_len] = xb;
                        buf_len += 1;
                        if buf_len == BLOCK {
                            flush(&mut sum, &buf_a, &buf_b);
                            buf_len = 0;
                        }
                    }
                } else {
                    debug_assert!(
                        false,
                        "broadcast validation should have handled all length combinations"
                    );
                    return Err(ErrorKind::Value);
                }
            }
            (SumproductOperand::Scalar(va), SumproductOperand::Reference(rb)) => {
                let xa = crate::functions::math::coerce_sumproduct_number(&va, locale)?;
                saw_nan = xa.is_nan();

                for addr_b in rb.iter_cells() {
                    let vb = ctx.get_cell_value(&rb.sheet_id, addr_b);
                    let xb = crate::functions::math::coerce_sumproduct_number(&vb, locale)?;

                    if saw_nan {
                        continue;
                    }
                    if xb.is_nan() {
                        saw_nan = true;
                        continue;
                    }

                    buf_a[buf_len] = xa;
                    buf_b[buf_len] = xb;
                    buf_len += 1;
                    if buf_len == BLOCK {
                        flush(&mut sum, &buf_a, &buf_b);
                        buf_len = 0;
                    }
                }
            }
            (SumproductOperand::Array(_), SumproductOperand::Array(_))
            | (SumproductOperand::Scalar(_), SumproductOperand::Scalar(_))
            | (SumproductOperand::Scalar(_), SumproductOperand::Array(_))
            | (SumproductOperand::Array(_), SumproductOperand::Scalar(_)) => {
                debug_assert!(
                    false,
                    "non-reference cases should have been handled by shared math::sumproduct path"
                );
                return Err(ErrorKind::Value);
            }
        }

        if saw_nan {
            return Ok(f64::NAN);
        }
        if buf_len > 0 {
            flush(&mut sum, &buf_a[..buf_len], &buf_b[..buf_len]);
        }
        Ok(sum)
    })();

    match result {
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
                let size = match reference_union_size(&ranges) {
                    Ok(v) => v,
                    Err(e) => return Value::Error(e),
                };
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

fn reference_union_size(ranges: &[crate::functions::Reference]) -> Result<u64, ErrorKind> {
    fn size_for_rects(rects: &[crate::functions::Reference]) -> Result<u64, ErrorKind> {
        if rects.is_empty() {
            return Ok(0);
        }

        // Convert to half-open row slabs: [start, end+1)
        let row_bound_len = rects.len().saturating_mul(2);
        let mut row_bounds: Vec<u32> = Vec::new();
        if row_bounds.try_reserve_exact(row_bound_len).is_err() {
            debug_assert!(
                false,
                "allocation failed (reference_union_size row_bounds, len={row_bound_len})"
            );
            return Err(ErrorKind::Num);
        }
        for r in rects {
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
            if intervals.try_reserve(rects.len()).is_err() {
                debug_assert!(false, "allocation failed (reference_union_size intervals)");
                return Err(ErrorKind::Num);
            }
            for r in rects {
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

        Ok(total)
    }

    if ranges.is_empty() {
        return Ok(0);
    }

    let mut total: u64 = 0;
    let mut by_sheet: std::collections::HashMap<
        crate::functions::SheetId,
        Vec<crate::functions::Reference>,
    > = std::collections::HashMap::new();
    for r in ranges {
        by_sheet
            .entry(r.sheet_id.clone())
            .or_default()
            .push(r.normalized());
    }

    for rects in by_sheet.into_values() {
        total = total.saturating_add(size_for_rects(&rects)?);
    }

    Ok(total)
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
        let number = number.coerce_to_number_with_ctx(ctx)?;
        let digits = digits.coerce_to_i64_with_ctx(ctx)?;
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
        let number = number.coerce_to_number_with_ctx(ctx)?;
        let digits = digits.coerce_to_i64_with_ctx(ctx)?;
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
    array_lift::lift1(value, |v| {
        Ok(Value::Number(v.coerce_to_number_with_ctx(ctx)?.floor()))
    })
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
    array_lift::lift1(value, |v| {
        Ok(Value::Number(v.coerce_to_number_with_ctx(ctx)?.abs()))
    })
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
        let n = n.coerce_to_number_with_ctx(ctx)?;
        let d = d.coerce_to_number_with_ctx(ctx)?;
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
        let number = number.coerce_to_number_with_ctx(ctx)?;
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

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
