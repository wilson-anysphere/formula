use std::collections::HashSet;

use crate::eval::CompiledExpr;
use crate::functions::statistical::{RankMethod, RankOrder};
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{Array, ErrorKind, Value};

const VAR_ARGS: usize = 255;

fn push_numbers_from_scalar(out: &mut Vec<f64>, value: Value) -> Result<(), ErrorKind> {
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
            let n = Value::Text(s).coerce_to_number()?;
            out.push(n);
            Ok(())
        }
        Value::Array(arr) => {
            for v in arr.iter() {
                match v {
                    Value::Error(e) => return Err(*e),
                    Value::Number(n) => out.push(*n),
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
    let sheet_id = reference.sheet_id;
    for addr in ctx.iter_reference_cells(reference) {
        let v = ctx.get_cell_value(sheet_id, addr);
        match v {
            Value::Error(e) => return Err(e),
            Value::Number(n) => out.push(n),
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
    Ok(())
}

fn push_numbers_from_reference_union(
    ctx: &dyn FunctionContext,
    out: &mut Vec<f64>,
    ranges: Vec<crate::functions::Reference>,
) -> Result<(), ErrorKind> {
    let mut seen = HashSet::new();
    for reference in ranges {
        let sheet_id = reference.sheet_id;
        for addr in ctx.iter_reference_cells(reference) {
            if !seen.insert((sheet_id, addr)) {
                continue;
            }
            let v = ctx.get_cell_value(sheet_id, addr);
            match v {
                Value::Error(e) => return Err(e),
                Value::Number(n) => out.push(n),
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
    Ok(())
}

fn push_numbers_from_arg(ctx: &dyn FunctionContext, out: &mut Vec<f64>, arg: ArgValue) -> Result<(), ErrorKind> {
    match arg {
        ArgValue::Scalar(v) => push_numbers_from_scalar(out, v),
        ArgValue::Reference(r) => push_numbers_from_reference(ctx, out, r),
        ArgValue::ReferenceUnion(ranges) => push_numbers_from_reference_union(ctx, out, ranges),
    }
}

fn collect_numbers(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Result<Vec<f64>, ErrorKind> {
    let mut out = Vec::new();
    for expr in args {
        push_numbers_from_arg(ctx, &mut out, ctx.eval_arg(expr))?;
    }
    Ok(out)
}

fn arg_to_numeric_sequence(ctx: &dyn FunctionContext, arg: ArgValue) -> Result<Vec<Option<f64>>, ErrorKind> {
    match arg {
        ArgValue::Scalar(v) => match v {
            Value::Error(e) => Err(e),
            Value::Number(n) => Ok(vec![Some(n)]),
            Value::Bool(b) => Ok(vec![Some(if b { 1.0 } else { 0.0 })]),
            Value::Blank => Ok(vec![None]),
            Value::Text(s) => Ok(vec![Some(Value::Text(s).coerce_to_number()?)]),
            Value::Array(arr) => {
                let mut out = Vec::with_capacity(arr.values.len());
                for v in arr.iter() {
                    match v {
                        Value::Error(e) => return Err(*e),
                        Value::Number(n) => out.push(Some(*n)),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Lambda(_)
                        | Value::Spill { .. }
                        | Value::Reference(_)
                        | Value::ReferenceUnion(_) => out.push(None),
                    }
                }
                Ok(out)
            }
            Value::Reference(_) | Value::ReferenceUnion(_) | Value::Lambda(_) | Value::Spill { .. } => {
                Err(ErrorKind::Value)
            }
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            let rows = (r.end.row - r.start.row + 1) as usize;
            let cols = (r.end.col - r.start.col + 1) as usize;
            let mut out = Vec::with_capacity(rows.saturating_mul(cols));
            for addr in r.iter_cells() {
                let v = ctx.get_cell_value(r.sheet_id, addr);
                match v {
                    Value::Error(e) => return Err(e),
                    Value::Number(n) => out.push(Some(n)),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Blank
                    | Value::Array(_)
                    | Value::Lambda(_)
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
                let rows = (r.end.row - r.start.row + 1) as usize;
                let cols = (r.end.col - r.start.col + 1) as usize;
                out.reserve(rows.saturating_mul(cols));
                for addr in r.iter_cells() {
                    if !seen.insert((r.sheet_id, addr)) {
                        continue;
                    }
                    let v = ctx.get_cell_value(r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Err(e),
                        Value::Number(n) => out.push(Some(n)),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Lambda(_)
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
        name: "STDEV.S",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: stdev_s_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "STDEV",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: stdev_s_fn,
    }
}

fn stdev_s_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::stdev_s(&values) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "STDEV.P",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: stdev_p_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "STDEVP",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: stdev_p_fn,
    }
}

fn stdev_p_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::stdev_p(&values) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "VAR.S",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: var_s_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "VAR",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: var_s_fn,
    }
}

fn var_s_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::var_s(&values) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "VAR.P",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: var_p_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "VARP",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: var_p_fn,
    }
}

fn var_p_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::var_p(&values) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "MEDIAN",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: median_fn,
    }
}

fn median_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::median(&values) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "MODE.SNGL",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: mode_sngl_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "MODE",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: mode_sngl_fn,
    }
}

fn mode_sngl_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::mode_sngl(&values) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "MODE.MULT",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: mode_mult_fn,
    }
}

fn mode_mult_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let modes = match crate::functions::statistical::mode_mult(&values) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let array_values: Vec<Value> = modes.into_iter().map(Value::Number).collect();
    Value::Array(Array::new(array_values.len(), 1, array_values))
}

inventory::submit! {
    FunctionSpec {
        name: "LARGE",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number],
        implementation: large_fn,
    }
}

fn large_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut values = Vec::new();
    if let Err(e) = push_numbers_from_arg(ctx, &mut values, ctx.eval_arg(&args[0])) {
        return Value::Error(e);
    }

    let k = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let Ok(k_usize) = usize::try_from(k) else {
        return Value::Error(ErrorKind::Num);
    };

    match crate::functions::statistical::large(&values, k_usize) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "SMALL",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number],
        implementation: small_fn,
    }
}

fn small_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut values = Vec::new();
    if let Err(e) = push_numbers_from_arg(ctx, &mut values, ctx.eval_arg(&args[0])) {
        return Value::Error(e);
    }

    let k = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let Ok(k_usize) = usize::try_from(k) else {
        return Value::Error(ErrorKind::Num);
    };

    match crate::functions::statistical::small(&values, k_usize) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "RANK.EQ",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Any, ValueType::Number],
        implementation: rank_eq_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "RANK",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Any, ValueType::Number],
        implementation: rank_eq_fn,
    }
}

fn rank_eq_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    rank_impl(ctx, args, RankMethod::Eq)
}

inventory::submit! {
    FunctionSpec {
        name: "RANK.AVG",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Any, ValueType::Number],
        implementation: rank_avg_fn,
    }
}

fn rank_avg_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    rank_impl(ctx, args, RankMethod::Avg)
}

fn rank_impl(ctx: &dyn FunctionContext, args: &[CompiledExpr], method: RankMethod) -> Value {
    let number = match eval_scalar_arg(ctx, &args[0]).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let mut values = Vec::new();
    if let Err(e) = push_numbers_from_arg(ctx, &mut values, ctx.eval_arg(&args[1])) {
        return Value::Error(e);
    }
    if values.is_empty() {
        return Value::Error(ErrorKind::NA);
    }

    let order = if args.len() == 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_number() {
            Ok(v) if v != 0.0 => RankOrder::Ascending,
            Ok(_) => RankOrder::Descending,
            Err(e) => return Value::Error(e),
        }
    } else {
        RankOrder::Descending
    };

    match crate::functions::statistical::rank(number, &values, order, method) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "PERCENTILE.INC",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number],
        implementation: percentile_inc_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "PERCENTILE",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number],
        implementation: percentile_inc_fn,
    }
}

fn percentile_inc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut values = Vec::new();
    if let Err(e) = push_numbers_from_arg(ctx, &mut values, ctx.eval_arg(&args[0])) {
        return Value::Error(e);
    }

    let k = match eval_scalar_arg(ctx, &args[1]).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    match crate::functions::statistical::percentile_inc(&values, k) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "PERCENTILE.EXC",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number],
        implementation: percentile_exc_fn,
    }
}

fn percentile_exc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut values = Vec::new();
    if let Err(e) = push_numbers_from_arg(ctx, &mut values, ctx.eval_arg(&args[0])) {
        return Value::Error(e);
    }

    let k = match eval_scalar_arg(ctx, &args[1]).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    match crate::functions::statistical::percentile_exc(&values, k) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "QUARTILE.INC",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number],
        implementation: quartile_inc_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "QUARTILE",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number],
        implementation: quartile_inc_fn,
    }
}

fn quartile_inc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut values = Vec::new();
    if let Err(e) = push_numbers_from_arg(ctx, &mut values, ctx.eval_arg(&args[0])) {
        return Value::Error(e);
    }

    let quart = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    match crate::functions::statistical::quartile_inc(&values, quart) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "QUARTILE.EXC",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number],
        implementation: quartile_exc_fn,
    }
}

fn quartile_exc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut values = Vec::new();
    if let Err(e) = push_numbers_from_arg(ctx, &mut values, ctx.eval_arg(&args[0])) {
        return Value::Error(e);
    }

    let quart = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    match crate::functions::statistical::quartile_exc(&values, quart) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "CORREL",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: correl_fn,
    }
}

fn correl_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let (xs, ys) = match collect_numeric_pairs(ctx, &args[0], &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::correl(&xs, &ys) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "COVARIANCE.S",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: covariance_s_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "COVAR",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: covariance_s_fn,
    }
}

fn covariance_s_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let (xs, ys) = match collect_numeric_pairs(ctx, &args[0], &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::covariance_s(&xs, &ys) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "COVARIANCE.P",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: covariance_p_fn,
    }
}

fn covariance_p_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let (xs, ys) = match collect_numeric_pairs(ctx, &args[0], &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::covariance_p(&xs, &ys) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}
