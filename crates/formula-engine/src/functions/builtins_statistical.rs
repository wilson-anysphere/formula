use std::collections::HashSet;

use crate::eval::{CompiledExpr, MAX_MATERIALIZED_ARRAY_CELLS};
use crate::functions::array_lift;
use crate::functions::statistical::{RankMethod, RankOrder};
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::simd;
use crate::value::{Array, ErrorKind, Value};

const VAR_ARGS: usize = 255;
const SIMD_AGGREGATE_BLOCK: usize = 1024;

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
    let reference = reference.normalized();
    ctx.record_reference(&reference);
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
        let reference = reference.normalized();
        ctx.record_reference(&reference);
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

#[derive(Debug, Default)]
struct NumbersA {
    /// Stored values that are not exactly `0`.
    values: Vec<f64>,
    /// Count of values that are exactly `0` (including implicit blanks).
    zeros: u64,
}

impl NumbersA {
    fn push(&mut self, value: f64) {
        if value == 0.0 {
            self.zeros = self.zeros.saturating_add(1);
        } else {
            self.values.push(value);
        }
    }

    fn push_zeros(&mut self, count: u64) {
        self.zeros = self.zeros.saturating_add(count);
    }

    fn count(&self) -> u64 {
        (self.values.len() as u64).saturating_add(self.zeros)
    }

    fn sum(&self) -> f64 {
        self.values.iter().copied().sum()
    }
}

fn push_numbers_a_from_scalar(out: &mut NumbersA, value: Value) -> Result<(), ErrorKind> {
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
        Value::Blank | Value::Text(_) | Value::Entity(_) | Value::Record(_) => {
            out.push(0.0);
            Ok(())
        }
        Value::Array(arr) => {
            for v in arr.iter() {
                match v {
                    Value::Error(e) => return Err(*e),
                    Value::Number(n) => out.push(*n),
                    Value::Bool(b) => out.push(if *b { 1.0 } else { 0.0 }),
                    Value::Blank | Value::Text(_) | Value::Entity(_) | Value::Record(_) => {
                        out.push(0.0)
                    }
                    Value::Lambda(_) => return Err(ErrorKind::Value),
                    Value::Array(_)
                    | Value::Spill { .. }
                    | Value::Reference(_)
                    | Value::ReferenceUnion(_) => out.push(0.0),
                }
            }
            Ok(())
        }
        Value::Reference(_) | Value::ReferenceUnion(_) | Value::Lambda(_) | Value::Spill { .. } => {
            Err(ErrorKind::Value)
        }
    }
}

fn push_numbers_a_from_reference(
    ctx: &dyn FunctionContext,
    out: &mut NumbersA,
    reference: crate::functions::Reference,
) -> Result<(), ErrorKind> {
    let size = reference.size();
    let mut seen = 0u64;
    for addr in ctx.iter_reference_cells(&reference) {
        seen = seen.saturating_add(1);
        let v = ctx.get_cell_value(&reference.sheet_id, addr);
        match v {
            Value::Error(e) => return Err(e),
            Value::Number(n) => out.push(n),
            Value::Bool(b) => out.push(if b { 1.0 } else { 0.0 }),
            Value::Blank | Value::Text(_) | Value::Entity(_) | Value::Record(_) => out.push(0.0),
            Value::Lambda(_) => return Err(ErrorKind::Value),
            Value::Array(_)
            | Value::Spill { .. }
            | Value::Reference(_)
            | Value::ReferenceUnion(_) => out.push(0.0),
        }
    }
    // Preserve Excel semantics: references include implicit blanks as zeros.
    out.push_zeros(size.saturating_sub(seen));
    Ok(())
}

fn push_numbers_a_from_reference_union(
    ctx: &dyn FunctionContext,
    out: &mut NumbersA,
    ranges: Vec<crate::functions::Reference>,
) -> Result<(), ErrorKind> {
    let size = reference_union_size(&ranges);
    let mut seen = HashSet::new();
    let mut seen_count: u64 = 0;
    for reference in ranges {
        for addr in ctx.iter_reference_cells(&reference) {
            if !seen.insert((reference.sheet_id.clone(), addr)) {
                continue;
            }
            seen_count = seen_count.saturating_add(1);
            let v = ctx.get_cell_value(&reference.sheet_id, addr);
            match v {
                Value::Error(e) => return Err(e),
                Value::Number(n) => out.push(n),
                Value::Bool(b) => out.push(if b { 1.0 } else { 0.0 }),
                Value::Blank | Value::Text(_) | Value::Entity(_) | Value::Record(_) => {
                    out.push(0.0)
                }
                Value::Lambda(_) => return Err(ErrorKind::Value),
                Value::Array(_)
                | Value::Spill { .. }
                | Value::Reference(_)
                | Value::ReferenceUnion(_) => out.push(0.0),
            }
        }
    }
    out.push_zeros(size.saturating_sub(seen_count));
    Ok(())
}

fn push_numbers_a_from_arg(
    ctx: &dyn FunctionContext,
    out: &mut NumbersA,
    arg: ArgValue,
) -> Result<(), ErrorKind> {
    match arg {
        ArgValue::Scalar(v) => push_numbers_a_from_scalar(out, v),
        ArgValue::Reference(r) => push_numbers_a_from_reference(ctx, out, r),
        ArgValue::ReferenceUnion(ranges) => push_numbers_a_from_reference_union(ctx, out, ranges),
    }
}

fn collect_numbers_a(
    ctx: &dyn FunctionContext,
    args: &[CompiledExpr],
) -> Result<NumbersA, ErrorKind> {
    let mut out = NumbersA::default();
    for expr in args {
        push_numbers_a_from_arg(ctx, &mut out, ctx.eval_arg(expr))?;
    }
    Ok(out)
}

inventory::submit! {
    FunctionSpec {
        name: "AVERAGEA",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: averagea_fn,
    }
}

fn averagea_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers_a(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let count = values.count();
    if count == 0 {
        return Value::Error(ErrorKind::Div0);
    }
    Value::Number(values.sum() / (count as f64))
}

inventory::submit! {
    FunctionSpec {
        name: "MAXA",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: maxa_fn,
    }
}

fn maxa_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers_a(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut best: Option<f64> = None;
    for &n in &values.values {
        best = Some(best.map_or(n, |b| b.max(n)));
    }
    if values.zeros > 0 {
        best = Some(best.map_or(0.0, |b| b.max(0.0)));
    }
    Value::Number(best.unwrap_or(0.0))
}

inventory::submit! {
    FunctionSpec {
        name: "MINA",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: mina_fn,
    }
}

fn mina_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers_a(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut best: Option<f64> = None;
    for &n in &values.values {
        best = Some(best.map_or(n, |b| b.min(n)));
    }
    if values.zeros > 0 {
        best = Some(best.map_or(0.0, |b| b.min(0.0)));
    }
    Value::Number(best.unwrap_or(0.0))
}

fn reference_union_size(ranges: &[crate::functions::Reference]) -> u64 {
    fn size_for_rects(rects: &[crate::functions::Reference]) -> u64 {
        if rects.is_empty() {
            return 0;
        }

        // Convert to half-open row slabs: [start, end+1)
        let mut row_bounds: Vec<u32> = Vec::new();
        if row_bounds
            .try_reserve_exact(rects.len().saturating_mul(2))
            .is_err()
        {
            debug_assert!(false, "allocation failed (reference_union_size row_bounds)");
            return u64::MAX;
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
                return u64::MAX;
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

        total
    }

    if ranges.is_empty() {
        return 0;
    }

    let mut total: u64 = 0;
    let mut by_sheet: std::collections::HashMap<
        crate::functions::SheetId,
        Vec<crate::functions::Reference>,
    > = std::collections::HashMap::new();
    if by_sheet.try_reserve(ranges.len()).is_err() {
        debug_assert!(false, "allocation failed (reference_union_size by_sheet)");
        return u64::MAX;
    }
    for r in ranges {
        by_sheet
            .entry(r.sheet_id.clone())
            .or_default()
            .push(r.normalized());
    }

    for rects in by_sheet.into_values() {
        total += size_for_rects(&rects);
    }

    total
}

inventory::submit! {
    FunctionSpec {
        name: "SUMSQ",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: sumsq_fn,
    }
}

fn sumsq_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    fn flush(buf: &[f64], sum: &mut f64, c: &mut f64) -> bool {
        if buf.is_empty() {
            return false;
        }
        let block_sum = simd::sumproduct_ignore_nan_f64(buf, buf);
        if !block_sum.is_finite() {
            return true;
        }

        // Kahan-compensated summation for better numeric stability while still allowing SIMD
        // acceleration for the hot (squaring + summing) inner loop.
        let y = block_sum - *c;
        let t = *sum + y;
        *c = (t - *sum) - y;
        *sum = t;
        !sum.is_finite()
    }

    let mut sum = 0.0;
    let mut c = 0.0;
    let mut saw_nonfinite = false;
    let mut buf = [0.0_f64; SIMD_AGGREGATE_BLOCK];
    let mut len = 0usize;

    let push_number = |n: f64,
                       sum: &mut f64,
                       c: &mut f64,
                       saw_nonfinite: &mut bool,
                       buf: &mut [f64; SIMD_AGGREGATE_BLOCK],
                       len: &mut usize| {
        if *saw_nonfinite {
            return;
        }
        if !n.is_finite() {
            *saw_nonfinite = true;
            *len = 0;
            return;
        }

        buf[*len] = n;
        *len += 1;
        if *len == SIMD_AGGREGATE_BLOCK {
            if flush(&buf[..], sum, c) {
                *saw_nonfinite = true;
                *len = 0;
            } else {
                *len = 0;
            }
        }
    };

    for expr in args {
        match ctx.eval_arg(expr) {
            ArgValue::Scalar(v) => match v {
                Value::Error(e) => return Value::Error(e),
                Value::Number(n) => {
                    push_number(n, &mut sum, &mut c, &mut saw_nonfinite, &mut buf, &mut len);
                }
                Value::Bool(b) => {
                    push_number(
                        if b { 1.0 } else { 0.0 },
                        &mut sum,
                        &mut c,
                        &mut saw_nonfinite,
                        &mut buf,
                        &mut len,
                    );
                }
                Value::Blank => {}
                Value::Text(s) => {
                    let n = match Value::Text(s).coerce_to_number_with_ctx(ctx) {
                        Ok(n) => n,
                        Err(e) => return Value::Error(e),
                    };
                    push_number(n, &mut sum, &mut c, &mut saw_nonfinite, &mut buf, &mut len);
                }
                Value::Entity(_) | Value::Record(_) => return Value::Error(ErrorKind::Value),
                Value::Array(arr) => {
                    for v in arr.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => {
                                push_number(
                                    *n,
                                    &mut sum,
                                    &mut c,
                                    &mut saw_nonfinite,
                                    &mut buf,
                                    &mut len,
                                );
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
                Value::Reference(_)
                | Value::ReferenceUnion(_)
                | Value::Lambda(_)
                | Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            },
            ArgValue::Reference(r) => {
                for addr in ctx.iter_reference_cells(&r) {
                    match ctx.get_cell_value(&r.sheet_id, addr) {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => {
                            push_number(
                                n,
                                &mut sum,
                                &mut c,
                                &mut saw_nonfinite,
                                &mut buf,
                                &mut len,
                            );
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
            ArgValue::ReferenceUnion(ranges) => {
                let mut seen = HashSet::new();
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        match ctx.get_cell_value(&r.sheet_id, addr) {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => {
                                push_number(
                                    n,
                                    &mut sum,
                                    &mut c,
                                    &mut saw_nonfinite,
                                    &mut buf,
                                    &mut len,
                                );
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
            }
        }
    }

    if !saw_nonfinite && len > 0 && flush(&buf[..len], &mut sum, &mut c) {
        saw_nonfinite = true;
    }

    if saw_nonfinite || !sum.is_finite() {
        Value::Error(ErrorKind::Num)
    } else {
        Value::Number(sum)
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DEVSQ",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: devsq_fn,
    }
}

fn devsq_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::devsq(&values) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "AVEDEV",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: avedev_fn,
    }
}

fn avedev_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::avedev(&values) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "GEOMEAN",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: geomean_fn,
    }
}

fn geomean_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::geomean(&values) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "HARMEAN",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: harmean_fn,
    }
}

fn harmean_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::harmean(&values) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "TRIMMEAN",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number],
        implementation: trimmean_fn,
    }
}

fn trimmean_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut values = Vec::new();
    if let Err(e) = push_numbers_from_arg(ctx, &mut values, ctx.eval_arg(&args[0])) {
        return Value::Error(e);
    }

    let percent = match eval_scalar_arg(ctx, &args[1]).coerce_to_number_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    match crate::functions::statistical::trimmean(&values, percent) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "STANDARDIZE",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: standardize_fn,
    }
}

fn standardize_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let mean = array_lift::eval_arg(ctx, &args[1]);
    let std_dev = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(x, mean, std_dev, |x, mean, std_dev| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let mean = mean.coerce_to_number_with_ctx(ctx)?;
        let std_dev = std_dev.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(crate::functions::statistical::standardize(
            x, mean, std_dev,
        )?))
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
                let total = arr.values.len();
                if total > MAX_MATERIALIZED_ARRAY_CELLS {
                    debug_assert!(
                        false,
                        "numeric sequence exceeds materialization limit (cells={total})"
                    );
                    return Err(ErrorKind::Spill);
                }
                let mut out: Vec<Option<f64>> = Vec::new();
                if out.try_reserve_exact(total).is_err() {
                    debug_assert!(false, "numeric sequence allocation failed (cells={total})");
                    return Err(ErrorKind::Num);
                }
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
            let total = match rows.checked_mul(cols) {
                Some(v) => v,
                None => return Err(ErrorKind::Spill),
            };
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Err(ErrorKind::Spill);
            }
            let mut out: Vec<Option<f64>> = Vec::new();
            if out.try_reserve_exact(total).is_err() {
                debug_assert!(false, "numeric sequence allocation failed (cells={total})");
                return Err(ErrorKind::Num);
            }
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
                let reserve = match rows.checked_mul(cols) {
                    Some(v) => v,
                    None => return Err(ErrorKind::Spill),
                };
                if out.len().saturating_add(reserve) > MAX_MATERIALIZED_ARRAY_CELLS {
                    return Err(ErrorKind::Spill);
                }
                if out.try_reserve(reserve).is_err() {
                    debug_assert!(false, "numeric sequence allocation failed (reserve={reserve})");
                    return Err(ErrorKind::Num);
                }
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::date::ExcelDateSystem;
    use crate::functions::SheetId;
    use crate::eval::CellAddr;
    use crate::value::Lambda;
    use chrono::{LocalResult, TimeZone};
    use std::collections::HashMap;

    struct PanicGetCellCtx;

    impl FunctionContext for PanicGetCellCtx {
        fn eval_arg(&self, _expr: &CompiledExpr) -> ArgValue {
            ArgValue::Scalar(Value::Blank)
        }

        fn eval_scalar(&self, _expr: &CompiledExpr) -> Value {
            Value::Blank
        }

        fn eval_formula(&self, _expr: &CompiledExpr) -> Value {
            Value::Blank
        }

        fn eval_formula_with_bindings(
            &self,
            _expr: &CompiledExpr,
            _bindings: &HashMap<String, Value>,
        ) -> Value {
            Value::Blank
        }

        fn capture_lexical_env(&self) -> HashMap<String, Value> {
            HashMap::new()
        }

        fn apply_implicit_intersection(&self, _reference: &crate::functions::Reference) -> Value {
            Value::Blank
        }

        fn get_cell_value(&self, _sheet_id: &SheetId, _addr: CellAddr) -> Value {
            panic!("unexpected get_cell_value call (materialization should have been guarded)");
        }

        fn iter_reference_cells<'a>(
            &'a self,
            _reference: &'a crate::functions::Reference,
        ) -> Box<dyn Iterator<Item = CellAddr> + 'a> {
            Box::new(std::iter::empty())
        }

        fn now_utc(&self) -> chrono::DateTime<chrono::Utc> {
            match chrono::Utc.timestamp_opt(0, 0) {
                LocalResult::Single(dt) => dt,
                other => {
                    debug_assert!(false, "expected epoch timestamp to be valid, got {other:?}");
                    chrono::DateTime::<chrono::Utc>::MIN_UTC
                }
            }
        }

        fn date_system(&self) -> ExcelDateSystem {
            ExcelDateSystem::EXCEL_1900
        }

        fn current_sheet_id(&self) -> usize {
            0
        }

        fn current_cell_addr(&self) -> CellAddr {
            CellAddr { row: 0, col: 0 }
        }

        fn push_local_scope(&self) {}

        fn pop_local_scope(&self) {}

        fn set_local(&self, _name: &str, _value: ArgValue) {}

        fn make_lambda(&self, _params: Vec<String>, _body: CompiledExpr) -> Value {
            Value::Error(ErrorKind::Value)
        }

        fn eval_lambda(&self, _lambda: &Lambda, _args: Vec<ArgValue>) -> Value {
            Value::Error(ErrorKind::Value)
        }

        fn volatile_rand_u64(&self) -> u64 {
            0
        }
    }

    #[test]
    fn arg_to_numeric_sequence_bails_out_before_materializing_oversize_reference() {
        let r = crate::functions::Reference {
            sheet_id: SheetId::Local(0),
            start: CellAddr { row: 0, col: 0 },
            end: CellAddr {
                row: MAX_MATERIALIZED_ARRAY_CELLS as u32,
                col: 0,
            },
        };
        let ctx = PanicGetCellCtx;
        let err = arg_to_numeric_sequence(&ctx, ArgValue::Reference(r)).unwrap_err();
        assert_eq!(err, ErrorKind::Spill);
    }
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

inventory::submit! {
    FunctionSpec {
        name: "STDEVA",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: stdeva_fn,
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

fn stdeva_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers_a(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::stdev_s_with_zeros(&values.values, values.zeros) {
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

inventory::submit! {
    FunctionSpec {
        name: "STDEVPA",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: stdevpa_fn,
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

fn stdevpa_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers_a(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::stdev_p_with_zeros(&values.values, values.zeros) {
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

inventory::submit! {
    FunctionSpec {
        name: "VARA",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: vara_fn,
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

fn vara_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers_a(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::var_s_with_zeros(&values.values, values.zeros) {
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

inventory::submit! {
    FunctionSpec {
        name: "VARPA",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: varpa_fn,
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

fn varpa_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers_a(ctx, args) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::var_p_with_zeros(&values.values, values.zeros) {
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

    let mut array_values: Vec<Value> = Vec::new();
    if array_values.try_reserve_exact(modes.len()).is_err() {
        debug_assert!(
            false,
            "allocation failed (MODE.MULT output, len={})",
            modes.len()
        );
        return Value::Error(ErrorKind::Num);
    }
    for value in modes {
        array_values.push(Value::Number(value));
    }
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

    let k = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64_with_ctx(ctx) {
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

    let k = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64_with_ctx(ctx) {
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
    let number = match eval_scalar_arg(ctx, &args[0]).coerce_to_number_with_ctx(ctx) {
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
        match eval_scalar_arg(ctx, &args[2]).coerce_to_number_with_ctx(ctx) {
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

    let k = match eval_scalar_arg(ctx, &args[1]).coerce_to_number_with_ctx(ctx) {
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

    let k = match eval_scalar_arg(ctx, &args[1]).coerce_to_number_with_ctx(ctx) {
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

    let quart = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64_with_ctx(ctx) {
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

    let quart = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64_with_ctx(ctx) {
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
        name: "PERCENTRANK.INC",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: percentrank_inc_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "PERCENTRANK",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: percentrank_inc_fn,
    }
}

fn percentrank_inc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    percentrank_impl(ctx, args, PercentrankVariant::Inclusive)
}

inventory::submit! {
    FunctionSpec {
        name: "PERCENTRANK.EXC",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: percentrank_exc_fn,
    }
}

fn percentrank_exc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    percentrank_impl(ctx, args, PercentrankVariant::Exclusive)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PercentrankVariant {
    Inclusive,
    Exclusive,
}

fn percentrank_impl(
    ctx: &dyn FunctionContext,
    args: &[CompiledExpr],
    variant: PercentrankVariant,
) -> Value {
    let mut values = Vec::new();
    if let Err(e) = push_numbers_from_arg(ctx, &mut values, ctx.eval_arg(&args[0])) {
        return Value::Error(e);
    }

    let x = match eval_scalar_arg(ctx, &args[1]).coerce_to_number_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let significance = if args.len() == 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_i64_with_ctx(ctx) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        3
    };
    if !(1..=15).contains(&significance) {
        return Value::Error(ErrorKind::Num);
    }

    let raw = match variant {
        PercentrankVariant::Inclusive => crate::functions::statistical::percentrank_inc(&values, x),
        PercentrankVariant::Exclusive => crate::functions::statistical::percentrank_exc(&values, x),
    };
    let raw = match raw {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let digits = significance as i32;
    let factor = 10f64.powi(digits);
    if !factor.is_finite() || factor == 0.0 || !raw.is_finite() {
        return Value::Error(ErrorKind::Num);
    }
    let out = (raw * factor).round() / factor;
    if out.is_finite() {
        Value::Number(out)
    } else {
        Value::Error(ErrorKind::Num)
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

inventory::submit! {
    FunctionSpec {
        name: "PEARSON",
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

inventory::submit! {
    FunctionSpec {
        name: "RSQ",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: rsq_fn,
    }
}

fn rsq_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // Excel signature: RSQ(known_y's, known_x's)
    let (xs, ys) = match collect_numeric_pairs(ctx, &args[1], &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::rsq(&xs, &ys) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "SLOPE",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: slope_fn,
    }
}

fn slope_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // Excel signature: SLOPE(known_y's, known_x's)
    let (xs, ys) = match collect_numeric_pairs(ctx, &args[1], &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::slope(&xs, &ys) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "INTERCEPT",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: intercept_fn,
    }
}

fn intercept_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // Excel signature: INTERCEPT(known_y's, known_x's)
    let (xs, ys) = match collect_numeric_pairs(ctx, &args[1], &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::intercept(&xs, &ys) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "STEYX",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: steyx_fn,
    }
}

fn steyx_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // Excel signature: STEYX(known_y's, known_x's)
    let (xs, ys) = match collect_numeric_pairs(ctx, &args[1], &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::statistical::steyx(&xs, &ys) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "FORECAST",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Any, ValueType::Any],
        implementation: forecast_linear_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "FORECAST.LINEAR",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Any, ValueType::Any],
        implementation: forecast_linear_fn,
    }
}

fn forecast_linear_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // Excel signature: FORECAST(x, known_y's, known_x's)
    let x = match eval_scalar_arg(ctx, &args[0]).coerce_to_number_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    // `collect_numeric_pairs` returns (xs, ys), so swap known_x / known_y.
    let (xs, ys) = match collect_numeric_pairs(ctx, &args[2], &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let slope = match crate::functions::statistical::slope(&xs, &ys) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let n = xs.len() as f64;
    if n == 0.0 {
        return Value::Error(ErrorKind::Div0);
    }
    let mean_x = xs.iter().sum::<f64>() / n;
    let mean_y = ys.iter().sum::<f64>() / n;
    if !mean_x.is_finite() || !mean_y.is_finite() {
        return Value::Error(ErrorKind::Num);
    }
    let intercept = mean_y - slope * mean_x;
    let out = intercept + slope * x;
    if out.is_finite() {
        Value::Number(out)
    } else {
        Value::Error(ErrorKind::Num)
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

// ------------------------------------------------------------------
// Discrete distributions + hypothesis tests (Excel compatibility)
// ------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "BINOM.DIST",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: binom_dist_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "BINOMDIST",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: binom_dist_fn,
    }
}

fn binom_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number_s = array_lift::eval_arg(ctx, &args[0]);
    let trials = array_lift::eval_arg(ctx, &args[1]);
    let probability_s = array_lift::eval_arg(ctx, &args[2]);
    let cumulative = array_lift::eval_arg(ctx, &args[3]);
    array_lift::lift4(
        number_s,
        trials,
        probability_s,
        cumulative,
        |number_s, trials, p, cumulative| {
            let number_s = number_s.coerce_to_number_with_ctx(ctx)?;
            let trials = trials.coerce_to_number_with_ctx(ctx)?;
            let p = p.coerce_to_number_with_ctx(ctx)?;
            let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
            Ok(Value::Number(crate::functions::statistical::binom_dist(
                number_s, trials, p, cumulative,
            )?))
        },
    )
}

inventory::submit! {
    FunctionSpec {
        name: "BINOM.DIST.RANGE",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: binom_dist_range_fn,
    }
}

fn binom_dist_range_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let trials = array_lift::eval_arg(ctx, &args[0]);
    let probability_s = array_lift::eval_arg(ctx, &args[1]);
    let number_s = array_lift::eval_arg(ctx, &args[2]);

    if args.len() == 3 {
        array_lift::lift3(trials, probability_s, number_s, |trials, p, number_s| {
            let trials = trials.coerce_to_number_with_ctx(ctx)?;
            let p = p.coerce_to_number_with_ctx(ctx)?;
            let number_s = number_s.coerce_to_number_with_ctx(ctx)?;
            Ok(Value::Number(
                crate::functions::statistical::binom_dist_range(trials, p, number_s, None)?,
            ))
        })
    } else {
        let number_s2 = array_lift::eval_arg(ctx, &args[3]);
        array_lift::lift4(
            trials,
            probability_s,
            number_s,
            number_s2,
            |trials, p, number_s, number_s2| {
                let trials = trials.coerce_to_number_with_ctx(ctx)?;
                let p = p.coerce_to_number_with_ctx(ctx)?;
                let number_s = number_s.coerce_to_number_with_ctx(ctx)?;
                let number_s2 = number_s2.coerce_to_number_with_ctx(ctx)?;
                Ok(Value::Number(
                    crate::functions::statistical::binom_dist_range(
                        trials,
                        p,
                        number_s,
                        Some(number_s2),
                    )?,
                ))
            },
        )
    }
}

inventory::submit! {
    FunctionSpec {
        name: "BINOM.INV",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: binom_inv_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "CRITBINOM",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: binom_inv_fn,
    }
}

fn binom_inv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let trials = array_lift::eval_arg(ctx, &args[0]);
    let probability_s = array_lift::eval_arg(ctx, &args[1]);
    let alpha = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(trials, probability_s, alpha, |trials, p, alpha| {
        let trials = trials.coerce_to_number_with_ctx(ctx)?;
        let p = p.coerce_to_number_with_ctx(ctx)?;
        let alpha = alpha.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(crate::functions::statistical::binom_inv(
            trials, p, alpha,
        )?))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "POISSON.DIST",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: poisson_dist_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "POISSON",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: poisson_dist_fn,
    }
}

fn poisson_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let mean = array_lift::eval_arg(ctx, &args[1]);
    let cumulative = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(x, mean, cumulative, |x, mean, cumulative| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let mean = mean.coerce_to_number_with_ctx(ctx)?;
        let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
        Ok(Value::Number(crate::functions::statistical::poisson_dist(
            x, mean, cumulative,
        )?))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "NEGBINOM.DIST",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: negbinom_dist_fn,
    }
}

fn negbinom_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number_f = array_lift::eval_arg(ctx, &args[0]);
    let number_s = array_lift::eval_arg(ctx, &args[1]);
    let probability_s = array_lift::eval_arg(ctx, &args[2]);
    let cumulative = array_lift::eval_arg(ctx, &args[3]);
    array_lift::lift4(
        number_f,
        number_s,
        probability_s,
        cumulative,
        |number_f, number_s, p, cumulative| {
            let number_f = number_f.coerce_to_number_with_ctx(ctx)?;
            let number_s = number_s.coerce_to_number_with_ctx(ctx)?;
            let p = p.coerce_to_number_with_ctx(ctx)?;
            let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
            Ok(Value::Number(crate::functions::statistical::negbinom_dist(
                number_f, number_s, p, cumulative,
            )?))
        },
    )
}

inventory::submit! {
    FunctionSpec {
        name: "NEGBINOMDIST",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: negbinomdist_fn,
    }
}

fn negbinomdist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number_f = array_lift::eval_arg(ctx, &args[0]);
    let number_s = array_lift::eval_arg(ctx, &args[1]);
    let probability_s = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(
        number_f,
        number_s,
        probability_s,
        |number_f, number_s, p| {
            let number_f = number_f.coerce_to_number_with_ctx(ctx)?;
            let number_s = number_s.coerce_to_number_with_ctx(ctx)?;
            let p = p.coerce_to_number_with_ctx(ctx)?;
            Ok(Value::Number(crate::functions::statistical::negbinom_dist(
                number_f, number_s, p, false,
            )?))
        },
    )
}

inventory::submit! {
    FunctionSpec {
        name: "HYPGEOM.DIST",
        min_args: 5,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Bool,
        ],
        implementation: hypgeom_dist_fn,
    }
}

fn hypgeom_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let sample_s = array_lift::eval_arg(ctx, &args[0]);
    let number_sample = array_lift::eval_arg(ctx, &args[1]);
    let population_s = array_lift::eval_arg(ctx, &args[2]);
    let number_pop = array_lift::eval_arg(ctx, &args[3]);
    let cumulative = array_lift::eval_arg(ctx, &args[4]);
    array_lift::lift5(
        sample_s,
        number_sample,
        population_s,
        number_pop,
        cumulative,
        |sample_s, number_sample, population_s, number_pop, cumulative| {
            let sample_s = sample_s.coerce_to_number_with_ctx(ctx)?;
            let number_sample = number_sample.coerce_to_number_with_ctx(ctx)?;
            let population_s = population_s.coerce_to_number_with_ctx(ctx)?;
            let number_pop = number_pop.coerce_to_number_with_ctx(ctx)?;
            let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
            Ok(Value::Number(crate::functions::statistical::hypgeom_dist(
                sample_s,
                number_sample,
                population_s,
                number_pop,
                cumulative,
            )?))
        },
    )
}

inventory::submit! {
    FunctionSpec {
        name: "HYPGEOMDIST",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: hypgeomdist_fn,
    }
}

fn hypgeomdist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let sample_s = array_lift::eval_arg(ctx, &args[0]);
    let number_sample = array_lift::eval_arg(ctx, &args[1]);
    let population_s = array_lift::eval_arg(ctx, &args[2]);
    let number_pop = array_lift::eval_arg(ctx, &args[3]);
    array_lift::lift4(
        sample_s,
        number_sample,
        population_s,
        number_pop,
        |sample_s, number_sample, population_s, number_pop| {
            let sample_s = sample_s.coerce_to_number_with_ctx(ctx)?;
            let number_sample = number_sample.coerce_to_number_with_ctx(ctx)?;
            let population_s = population_s.coerce_to_number_with_ctx(ctx)?;
            let number_pop = number_pop.coerce_to_number_with_ctx(ctx)?;
            Ok(Value::Number(crate::functions::statistical::hypgeom_dist(
                sample_s,
                number_sample,
                population_s,
                number_pop,
                false,
            )?))
        },
    )
}

#[derive(Debug, Clone)]
enum Range2D {
    Scalar(Value),
    Reference(crate::functions::Reference),
    Array(Array),
}

impl Range2D {
    fn try_from_arg(arg: ArgValue) -> Result<Self, ErrorKind> {
        match arg {
            ArgValue::Scalar(Value::Error(e)) => Err(e),
            ArgValue::Scalar(Value::Reference(r)) => Ok(Self::Reference(r.normalized())),
            ArgValue::Scalar(Value::Array(arr)) => Ok(Self::Array(arr)),
            ArgValue::Scalar(v) => Ok(Self::Scalar(v)),
            ArgValue::Reference(r) => Ok(Self::Reference(r.normalized())),
            ArgValue::ReferenceUnion(_) => Err(ErrorKind::Value),
        }
    }

    fn record_reference(&self, ctx: &dyn FunctionContext) {
        if let Range2D::Reference(r) = self {
            ctx.record_reference(r);
        }
    }

    fn shape(&self) -> (usize, usize) {
        match self {
            Range2D::Scalar(_) => (1, 1),
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
            Range2D::Scalar(v) => {
                if row == 0 && col == 0 {
                    v.clone()
                } else {
                    Value::Blank
                }
            }
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
        name: "PROB",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: prob_fn,
    }
}

fn prob_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x_range = match Range2D::try_from_arg(ctx.eval_arg(&args[0])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let prob_range = match Range2D::try_from_arg(ctx.eval_arg(&args[1])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    x_range.record_reference(ctx);
    prob_range.record_reference(ctx);

    let (rows_x, cols_x) = x_range.shape();
    let (rows_p, cols_p) = prob_range.shape();
    if (rows_x, cols_x) != (rows_p, cols_p) {
        return Value::Error(ErrorKind::NA);
    }
    let len = rows_x.saturating_mul(cols_x);
    if len > MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }

    let mut xs: Vec<f64> = Vec::new();
    if xs.try_reserve_exact(len).is_err() {
        debug_assert!(false, "allocation failed (PROB xs={len})");
        return Value::Error(ErrorKind::Num);
    }
    let mut ps: Vec<f64> = Vec::new();
    if ps.try_reserve_exact(len).is_err() {
        debug_assert!(false, "allocation failed (PROB ps={len})");
        return Value::Error(ErrorKind::Num);
    }
    for r in 0..rows_x {
        for c in 0..cols_x {
            let x = x_range.get(ctx, r, c).coerce_to_number_with_ctx(ctx);
            let p = prob_range.get(ctx, r, c).coerce_to_number_with_ctx(ctx);
            let x = match x {
                Ok(v) => v,
                Err(e) => return Value::Error(e),
            };
            let p = match p {
                Ok(v) => v,
                Err(e) => return Value::Error(e),
            };
            xs.push(x);
            ps.push(p);
        }
    }

    let lower = match eval_scalar_arg(ctx, &args[2]).coerce_to_number_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let upper = if args.len() == 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_number_with_ctx(ctx) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    match crate::functions::statistical::prob(&xs, &ps, lower, upper) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "Z.TEST",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: z_test_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "ZTEST",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: z_test_fn,
    }
}

fn z_test_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut values = Vec::new();
    if let Err(e) = push_numbers_from_arg(ctx, &mut values, ctx.eval_arg(&args[0])) {
        return Value::Error(e);
    }

    let x = match eval_scalar_arg(ctx, &args[1]).coerce_to_number_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let sigma = if args.len() == 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_number_with_ctx(ctx) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    match crate::functions::statistical::z_test(&values, x, sigma) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "T.TEST",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: t_test_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "TTEST",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: t_test_fn,
    }
}

fn t_test_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let tails = match eval_scalar_arg(ctx, &args[2]).coerce_to_i64_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let test_type = match eval_scalar_arg(ctx, &args[3]).coerce_to_i64_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let (xs, ys) = if test_type == 1 {
        match collect_numeric_pairs(ctx, &args[0], &args[1]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        let mut xs = Vec::new();
        if let Err(e) = push_numbers_from_arg(ctx, &mut xs, ctx.eval_arg(&args[0])) {
            return Value::Error(e);
        }
        let mut ys = Vec::new();
        if let Err(e) = push_numbers_from_arg(ctx, &mut ys, ctx.eval_arg(&args[1])) {
            return Value::Error(e);
        }
        (xs, ys)
    };

    match crate::functions::statistical::t_test(&xs, &ys, tails, test_type) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "F.TEST",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: f_test_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "FTEST",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: f_test_fn,
    }
}

fn f_test_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut xs = Vec::new();
    if let Err(e) = push_numbers_from_arg(ctx, &mut xs, ctx.eval_arg(&args[0])) {
        return Value::Error(e);
    }
    let mut ys = Vec::new();
    if let Err(e) = push_numbers_from_arg(ctx, &mut ys, ctx.eval_arg(&args[1])) {
        return Value::Error(e);
    }

    match crate::functions::statistical::f_test(&xs, &ys) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "CHISQ.TEST",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: chisq_test_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "CHITEST",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: chisq_test_fn,
    }
}

fn chisq_test_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let actual = match Range2D::try_from_arg(ctx.eval_arg(&args[0])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let expected = match Range2D::try_from_arg(ctx.eval_arg(&args[1])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    actual.record_reference(ctx);
    expected.record_reference(ctx);

    let (rows_a, cols_a) = actual.shape();
    let (rows_e, cols_e) = expected.shape();
    if (rows_a, cols_a) != (rows_e, cols_e) {
        return Value::Error(ErrorKind::NA);
    }
    let len = rows_a.saturating_mul(cols_a);
    if len > MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }

    let mut actual_vals: Vec<f64> = Vec::new();
    if actual_vals.try_reserve_exact(len).is_err() {
        debug_assert!(false, "allocation failed (CHISQ.TEST actual={len})");
        return Value::Error(ErrorKind::Num);
    }
    let mut expected_vals: Vec<f64> = Vec::new();
    if expected_vals.try_reserve_exact(len).is_err() {
        debug_assert!(false, "allocation failed (CHISQ.TEST expected={len})");
        return Value::Error(ErrorKind::Num);
    }
    for r in 0..rows_a {
        for c in 0..cols_a {
            let a = match actual.get(ctx, r, c).coerce_to_number_with_ctx(ctx) {
                Ok(v) => v,
                Err(e) => return Value::Error(e),
            };
            let e = match expected.get(ctx, r, c).coerce_to_number_with_ctx(ctx) {
                Ok(v) => v,
                Err(e) => return Value::Error(e),
            };
            actual_vals.push(a);
            expected_vals.push(e);
        }
    }

    match crate::functions::statistical::chisq_test(&actual_vals, &expected_vals, rows_a, cols_a) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
