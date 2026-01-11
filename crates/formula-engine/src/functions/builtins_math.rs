use crate::eval::CompiledExpr;
use crate::error::ExcelError;
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::simd::{CmpOp, NumericCriteria};
use crate::value::{ErrorKind, Value};

const VAR_ARGS: usize = 255;

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
                            Value::Bool(_) | Value::Text(_) | Value::Blank | Value::Array(_) | Value::Spill { .. } => {}
                        }
                    }
                }
                Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            },
            ArgValue::Reference(r) => {
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => acc += n,
                        // Excel quirk: logicals/text in references are ignored by SUM.
                        Value::Bool(_) | Value::Text(_) | Value::Blank | Value::Array(_) | Value::Spill { .. } => {}
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                for r in ranges {
                    for addr in r.iter_cells() {
                        let v = ctx.get_cell_value(r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => acc += n,
                            // Excel quirk: logicals/text in references are ignored by SUM.
                            Value::Bool(_) | Value::Text(_) | Value::Blank | Value::Array(_) | Value::Spill { .. } => {}
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
                            Value::Bool(_) | Value::Text(_) | Value::Blank | Value::Array(_) | Value::Spill { .. } => {}
                        }
                    }
                }
                Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            },
            ArgValue::Reference(r) => {
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => {
                            acc += n;
                            count += 1;
                        }
                        // Ignore logical/text/blank in references.
                        Value::Bool(_) | Value::Text(_) | Value::Blank | Value::Array(_) | Value::Spill { .. } => {}
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                for r in ranges {
                    for addr in r.iter_cells() {
                        let v = ctx.get_cell_value(r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => {
                                acc += n;
                                count += 1;
                            }
                            // Ignore logical/text/blank in references.
                            Value::Bool(_) | Value::Text(_) | Value::Blank | Value::Array(_) | Value::Spill { .. } => {}
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
            ArgValue::Scalar(v) => {
                let n = match v.coerce_to_number() {
                    Ok(n) => n,
                    Err(e) => return Value::Error(e),
                };
                best = Some(best.map(|b| b.min(n)).unwrap_or(n));
            }
            ArgValue::Reference(r) => {
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => best = Some(best.map(|b| b.min(n)).unwrap_or(n)),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Spill { .. } => {}
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                for r in ranges {
                    for addr in r.iter_cells() {
                        let v = ctx.get_cell_value(r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => best = Some(best.map(|b| b.min(n)).unwrap_or(n)),
                            Value::Bool(_) | Value::Text(_) | Value::Blank | Value::Array(_) | Value::Spill { .. } => {}
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
            ArgValue::Scalar(v) => {
                let n = match v.coerce_to_number() {
                    Ok(n) => n,
                    Err(e) => return Value::Error(e),
                };
                best = Some(best.map(|b| b.max(n)).unwrap_or(n));
            }
            ArgValue::Reference(r) => {
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => best = Some(best.map(|b| b.max(n)).unwrap_or(n)),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Spill { .. } => {}
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                for r in ranges {
                    for addr in r.iter_cells() {
                        let v = ctx.get_cell_value(r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => best = Some(best.map(|b| b.max(n)).unwrap_or(n)),
                            Value::Bool(_) | Value::Text(_) | Value::Blank | Value::Array(_) | Value::Spill { .. } => {}
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
            ArgValue::Scalar(v) => {
                if matches!(v, Value::Number(_)) {
                    total += 1;
                }
            }
            ArgValue::Reference(r) => {
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(r.sheet_id, addr);
                    if matches!(v, Value::Number(_)) {
                        total += 1;
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                for r in ranges {
                    for addr in r.iter_cells() {
                        let v = ctx.get_cell_value(r.sheet_id, addr);
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
        arg_types: &[ValueType::Any],
        implementation: countif_fn,
    }
}

fn countif_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let ranges = match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(r) => vec![r],
        ArgValue::ReferenceUnion(ranges) => ranges,
        ArgValue::Scalar(Value::Error(e)) => return Value::Error(e),
        ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };

    let criteria_value = eval_scalar_arg(ctx, &args[1]);
    if let Value::Error(e) = criteria_value {
        return Value::Error(e);
    }
    let Some(criteria) = parse_numeric_criteria(&criteria_value) else {
        return Value::Error(ErrorKind::Value);
    };

    let mut count = 0u64;
    for range in ranges {
        for addr in range.iter_cells() {
            let v = ctx.get_cell_value(range.sheet_id, addr);
            if let Value::Number(n) = v {
                if matches_numeric_criteria(n, criteria) {
                    count += 1;
                }
            }
        }
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
        arg_types: &[ValueType::Any],
        implementation: sumproduct_fn,
    }
}

fn sumproduct_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let a = match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(r) => r.normalized(),
        ArgValue::ReferenceUnion(_) => return Value::Error(ErrorKind::Value),
        ArgValue::Scalar(Value::Error(e)) => return Value::Error(e),
        ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };
    let b = match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(r) => r.normalized(),
        ArgValue::ReferenceUnion(_) => return Value::Error(ErrorKind::Value),
        ArgValue::Scalar(Value::Error(e)) => return Value::Error(e),
        ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };

    let rows_a = a.end.row - a.start.row + 1;
    let cols_a = a.end.col - a.start.col + 1;
    let rows_b = b.end.row - b.start.row + 1;
    let cols_b = b.end.col - b.start.col + 1;
    if rows_a != rows_b || cols_a != cols_b {
        return Value::Error(ErrorKind::Value);
    }

    let len = (rows_a as usize).saturating_mul(cols_a as usize);
    let mut va = Vec::with_capacity(len);
    let mut vb = Vec::with_capacity(len);
    for addr in a.iter_cells() {
        va.push(ctx.get_cell_value(a.sheet_id, addr));
    }
    for addr in b.iter_cells() {
        vb.push(ctx.get_cell_value(b.sheet_id, addr));
    }

    match crate::functions::math::sumproduct(&[&va, &vb]) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

fn parse_numeric_criteria(value: &Value) -> Option<NumericCriteria> {
    match value {
        Value::Number(n) => Some(NumericCriteria::new(CmpOp::Eq, *n)),
        Value::Bool(b) => Some(NumericCriteria::new(CmpOp::Eq, if *b { 1.0 } else { 0.0 })),
        Value::Text(s) => parse_numeric_criteria_str(s),
        _ => None,
    }
}

fn parse_numeric_criteria_str(raw: &str) -> Option<NumericCriteria> {
    let s = raw.trim();
    let (op, rest) = if let Some(r) = s.strip_prefix(">=") {
        (CmpOp::Ge, r)
    } else if let Some(r) = s.strip_prefix("<=") {
        (CmpOp::Le, r)
    } else if let Some(r) = s.strip_prefix("<>") {
        (CmpOp::Ne, r)
    } else if let Some(r) = s.strip_prefix('>') {
        (CmpOp::Gt, r)
    } else if let Some(r) = s.strip_prefix('<') {
        (CmpOp::Lt, r)
    } else if let Some(r) = s.strip_prefix('=') {
        (CmpOp::Eq, r)
    } else {
        (CmpOp::Eq, s)
    };

    let rhs: f64 = rest.trim().parse().ok()?;
    Some(NumericCriteria::new(op, rhs))
}

fn matches_numeric_criteria(value: f64, criteria: NumericCriteria) -> bool {
    match criteria.op {
        CmpOp::Eq => value == criteria.rhs,
        CmpOp::Ne => value != criteria.rhs,
        CmpOp::Lt => value < criteria.rhs,
        CmpOp::Le => value <= criteria.rhs,
        CmpOp::Gt => value > criteria.rhs,
        CmpOp::Ge => value >= criteria.rhs,
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
            ArgValue::Scalar(v) => {
                if !matches!(v, Value::Blank) {
                    total += 1;
                }
            }
            ArgValue::Reference(r) => {
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(r.sheet_id, addr);
                    if !matches!(v, Value::Blank) {
                        total += 1;
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                for r in ranges {
                    for addr in r.iter_cells() {
                        let v = ctx.get_cell_value(r.sheet_id, addr);
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
            ArgValue::Scalar(v) => {
                if matches!(v, Value::Blank) || matches!(v, Value::Text(ref s) if s.is_empty()) {
                    total += 1;
                }
            }
            ArgValue::Reference(r) => {
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(r.sheet_id, addr);
                    if matches!(v, Value::Blank) || matches!(v, Value::Text(ref s) if s.is_empty())
                    {
                        total += 1;
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                for r in ranges {
                    for addr in r.iter_cells() {
                        let v = ctx.get_cell_value(r.sheet_id, addr);
                        if matches!(v, Value::Blank)
                            || matches!(v, Value::Text(ref s) if s.is_empty())
                        {
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
        name: "ROUND",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
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
        array_support: ArraySupport::ScalarOnly,
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
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: roundup_fn,
    }
}

fn roundup_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    round_impl(ctx, args, RoundMode::Up)
}

fn round_impl(ctx: &dyn FunctionContext, args: &[CompiledExpr], mode: RoundMode) -> Value {
    let number = match eval_scalar_arg(ctx, &args[0]).coerce_to_number() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let digits = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    Value::Number(round_with_mode(number, digits as i32, mode))
}

inventory::submit! {
    FunctionSpec {
        name: "INT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: int_fn,
    }
}

fn int_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let n = match eval_scalar_arg(ctx, &args[0]).coerce_to_number() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    Value::Number(n.floor())
}

inventory::submit! {
    FunctionSpec {
        name: "ABS",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: abs_fn,
    }
}

fn abs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let n = match eval_scalar_arg(ctx, &args[0]).coerce_to_number() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    Value::Number(n.abs())
}

inventory::submit! {
    FunctionSpec {
        name: "MOD",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: mod_fn,
    }
}

fn mod_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let n = match eval_scalar_arg(ctx, &args[0]).coerce_to_number() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let d = match eval_scalar_arg(ctx, &args[1]).coerce_to_number() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    if d == 0.0 {
        return Value::Error(ErrorKind::Div0);
    }
    Value::Number(n - d * (n / d).floor())
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
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: sign_fn,
    }
}

fn sign_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = match eval_scalar_arg(ctx, &args[0]).coerce_to_number() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    if !number.is_finite() {
        return Value::Error(ErrorKind::Num);
    }
    if number > 0.0 {
        Value::Number(1.0)
    } else if number < 0.0 {
        Value::Number(-1.0)
    } else {
        Value::Number(0.0)
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

fn rand_fn(_ctx: &dyn FunctionContext, _args: &[CompiledExpr]) -> Value {
    Value::Number(crate::functions::math::rand())
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
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let top = match eval_scalar_arg(ctx, &args[1]).coerce_to_number() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    match crate::functions::math::randbetween(bottom, top) {
        Ok(n) => Value::Number(n as f64),
        Err(e) => Value::Error(match e {
            ExcelError::Div0 => ErrorKind::Div0,
            ExcelError::Value => ErrorKind::Value,
            ExcelError::Num => ErrorKind::Num,
        }),
    }
}
