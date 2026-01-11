use std::collections::HashSet;

use crate::error::ExcelError;
use crate::eval::{CellAddr, CompiledExpr};
use crate::functions::{ArgValue, ArraySupport, FunctionContext, FunctionSpec, Reference};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{Array, ErrorKind, Value};

const VAR_ARGS: usize = 255;

fn excel_error_kind(err: ExcelError) -> ErrorKind {
    match err {
        ExcelError::Div0 => ErrorKind::Div0,
        ExcelError::Value => ErrorKind::Value,
        ExcelError::Num => ErrorKind::Num,
    }
}

fn reference_to_array(ctx: &dyn FunctionContext, reference: Reference) -> Array {
    let reference = reference.normalized();
    let rows = (reference.end.row - reference.start.row + 1) as usize;
    let cols = (reference.end.col - reference.start.col + 1) as usize;
    let mut values = Vec::with_capacity(rows.saturating_mul(cols));
    for row in reference.start.row..=reference.end.row {
        for col in reference.start.col..=reference.end.col {
            values.push(ctx.get_cell_value(&reference.sheet_id, CellAddr { row, col }));
        }
    }
    Array::new(rows, cols, values)
}

fn implicit_intersection_union(ctx: &dyn FunctionContext, ranges: &[Reference]) -> Value {
    // Excel's implicit intersection on a multi-area reference is ambiguous; we approximate by
    // succeeding only when exactly one area intersects.
    let mut hits = Vec::new();
    for r in ranges {
        let v = ctx.apply_implicit_intersection(r);
        if !matches!(v, Value::Error(ErrorKind::Value)) {
            hits.push(v);
        }
    }
    match hits.as_slice() {
        [only] => only.clone(),
        _ => Value::Error(ErrorKind::Value),
    }
}

fn scalar_from_arg(ctx: &dyn FunctionContext, arg: ArgValue) -> Value {
    match arg {
        ArgValue::Scalar(v) => v,
        ArgValue::Reference(r) => ctx.apply_implicit_intersection(&r),
        ArgValue::ReferenceUnion(ranges) => implicit_intersection_union(ctx, &ranges),
    }
}

fn dynamic_value_from_arg(ctx: &dyn FunctionContext, arg: ArgValue) -> Value {
    match arg {
        ArgValue::Scalar(v) => v,
        ArgValue::Reference(r) => {
            if r.is_single_cell() {
                ctx.get_cell_value(&r.sheet_id, r.start)
            } else {
                Value::Array(reference_to_array(ctx, r))
            }
        }
        ArgValue::ReferenceUnion(_) => Value::Error(ErrorKind::Value),
    }
}

fn elementwise_unary(value: &Value, f: impl Fn(&Value) -> Value) -> Value {
    match value {
        Value::Array(arr) => Value::Array(Array::new(arr.rows, arr.cols, arr.iter().map(f).collect())),
        other => f(other),
    }
}

fn elementwise_binary(left: &Value, right: &Value, f: impl Fn(&Value, &Value) -> Value) -> Value {
    match (left, right) {
        (Value::Array(left_arr), Value::Array(right_arr)) => {
            if left_arr.rows == right_arr.rows && left_arr.cols == right_arr.cols {
                return Value::Array(Array::new(
                    left_arr.rows,
                    left_arr.cols,
                    left_arr
                        .values
                        .iter()
                        .zip(right_arr.values.iter())
                        .map(|(a, b)| f(a, b))
                        .collect(),
                ));
            }

            if left_arr.rows == 1 && left_arr.cols == 1 {
                let scalar = left_arr.values.get(0).unwrap_or(&Value::Blank);
                return Value::Array(Array::new(
                    right_arr.rows,
                    right_arr.cols,
                    right_arr.values.iter().map(|b| f(scalar, b)).collect(),
                ));
            }

            if right_arr.rows == 1 && right_arr.cols == 1 {
                let scalar = right_arr.values.get(0).unwrap_or(&Value::Blank);
                return Value::Array(Array::new(
                    left_arr.rows,
                    left_arr.cols,
                    left_arr.values.iter().map(|a| f(a, scalar)).collect(),
                ));
            }

            Value::Error(ErrorKind::Value)
        }
        (Value::Array(left_arr), right_scalar) => Value::Array(Array::new(
            left_arr.rows,
            left_arr.cols,
            left_arr.values.iter().map(|a| f(a, right_scalar)).collect(),
        )),
        (left_scalar, Value::Array(right_arr)) => Value::Array(Array::new(
            right_arr.rows,
            right_arr.cols,
            right_arr.values.iter().map(|b| f(left_scalar, b)).collect(),
        )),
        (left_scalar, right_scalar) => f(left_scalar, right_scalar),
    }
}

fn values_from_range_arg(ctx: &dyn FunctionContext, arg: ArgValue) -> Result<Vec<Value>, ErrorKind> {
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
        ArgValue::ReferenceUnion(ranges) => {
            let mut seen = HashSet::new();
            let mut values = Vec::new();
            for r in ranges {
                let r = r.normalized();
                let rows = r.end.row - r.start.row + 1;
                let cols = r.end.col - r.start.col + 1;
                values.reserve((rows as usize).saturating_mul(cols as usize));
                for addr in r.iter_cells() {
                    if !seen.insert((r.sheet_id.clone(), addr)) {
                        continue;
                    }
                    values.push(ctx.get_cell_value(&r.sheet_id, addr));
                }
            }
            Ok(values)
        }
        ArgValue::Scalar(Value::Array(arr)) => Ok(arr.values),
        ArgValue::Scalar(Value::Error(e)) => Err(e),
        ArgValue::Scalar(v) => Ok(vec![v]),
    }
}

fn append_values_for_aggregate(ctx: &dyn FunctionContext, arg: ArgValue, out: &mut Vec<Value>) {
    match arg {
        ArgValue::Scalar(Value::Array(arr)) => out.extend(arr.values),
        ArgValue::Scalar(v) => out.push(v),
        ArgValue::Reference(r) => {
            for addr in ctx.iter_reference_cells(&r) {
                out.push(ctx.get_cell_value(&r.sheet_id, addr));
            }
        }
        ArgValue::ReferenceUnion(ranges) => {
            let mut seen = HashSet::new();
            for r in ranges {
                for addr in ctx.iter_reference_cells(&r) {
                    if !seen.insert((r.sheet_id.clone(), addr)) {
                        continue;
                    }
                    out.push(ctx.get_cell_value(&r.sheet_id, addr));
                }
            }
        }
    }
}

inventory::submit! {
    FunctionSpec {
        name: "PI",
        min_args: 0,
        max_args: 0,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[],
        implementation: pi_fn,
    }
}

fn pi_fn(_ctx: &dyn FunctionContext, _args: &[CompiledExpr]) -> Value {
    Value::Number(crate::functions::math::pi())
}

inventory::submit! {
    FunctionSpec {
        name: "SIN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: sin_fn,
    }
}

fn sin_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[0]));
    elementwise_unary(&value, |elem| match elem {
        Value::Error(e) => Value::Error(*e),
        other => {
            let n = match other.coerce_to_number() {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            match crate::functions::math::sin(n) {
                Ok(out) => Value::Number(out),
                Err(e) => Value::Error(excel_error_kind(e)),
            }
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "COS",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: cos_fn,
    }
}

fn cos_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[0]));
    elementwise_unary(&value, |elem| match elem {
        Value::Error(e) => Value::Error(*e),
        other => {
            let n = match other.coerce_to_number() {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            match crate::functions::math::cos(n) {
                Ok(out) => Value::Number(out),
                Err(e) => Value::Error(excel_error_kind(e)),
            }
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "TAN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: tan_fn,
    }
}

fn tan_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[0]));
    elementwise_unary(&value, |elem| match elem {
        Value::Error(e) => Value::Error(*e),
        other => {
            let n = match other.coerce_to_number() {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            match crate::functions::math::tan(n) {
                Ok(out) => Value::Number(out),
                Err(e) => Value::Error(excel_error_kind(e)),
            }
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "ASIN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: asin_fn,
    }
}

fn asin_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[0]));
    elementwise_unary(&value, |elem| match elem {
        Value::Error(e) => Value::Error(*e),
        other => {
            let n = match other.coerce_to_number() {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            match crate::functions::math::asin(n) {
                Ok(out) => Value::Number(out),
                Err(e) => Value::Error(excel_error_kind(e)),
            }
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "ACOS",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: acos_fn,
    }
}

fn acos_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[0]));
    elementwise_unary(&value, |elem| match elem {
        Value::Error(e) => Value::Error(*e),
        other => {
            let n = match other.coerce_to_number() {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            match crate::functions::math::acos(n) {
                Ok(out) => Value::Number(out),
                Err(e) => Value::Error(excel_error_kind(e)),
            }
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "ATAN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: atan_fn,
    }
}

fn atan_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[0]));
    elementwise_unary(&value, |elem| match elem {
        Value::Error(e) => Value::Error(*e),
        other => {
            let n = match other.coerce_to_number() {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            match crate::functions::math::atan(n) {
                Ok(out) => Value::Number(out),
                Err(e) => Value::Error(excel_error_kind(e)),
            }
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "ATAN2",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: atan2_fn,
    }
}

fn atan2_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x_num = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[0]));
    let y_num = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[1]));
    elementwise_binary(&x_num, &y_num, |x_num, y_num| {
        if let Value::Error(e) = x_num {
            return Value::Error(*e);
        }
        if let Value::Error(e) = y_num {
            return Value::Error(*e);
        }
        let x = match x_num.coerce_to_number() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let y = match y_num.coerce_to_number() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        match crate::functions::math::atan2(x, y) {
            Ok(out) => Value::Number(out),
            Err(e) => Value::Error(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "EXP",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: exp_fn,
    }
}

fn exp_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[0]));
    elementwise_unary(&value, |elem| match elem {
        Value::Error(e) => Value::Error(*e),
        other => {
            let n = match other.coerce_to_number() {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            match crate::functions::math::exp(n) {
                Ok(out) => Value::Number(out),
                Err(e) => Value::Error(excel_error_kind(e)),
            }
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "LN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: ln_fn,
    }
}

fn ln_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[0]));
    elementwise_unary(&value, |elem| match elem {
        Value::Error(e) => Value::Error(*e),
        other => {
            let n = match other.coerce_to_number() {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            match crate::functions::math::ln(n) {
                Ok(out) => Value::Number(out),
                Err(e) => Value::Error(excel_error_kind(e)),
            }
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "LOG",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: log_fn,
    }
}

fn log_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[0]));
    if args.len() == 1 {
        return elementwise_unary(&number, |elem| match elem {
            Value::Error(e) => Value::Error(*e),
            other => {
                let n = match other.coerce_to_number() {
                    Ok(n) => n,
                    Err(e) => return Value::Error(e),
                };
                match crate::functions::math::log(n, None) {
                    Ok(out) => Value::Number(out),
                    Err(e) => Value::Error(excel_error_kind(e)),
                }
            }
        });
    }

    let base = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[1]));
    elementwise_binary(&number, &base, |num, base| {
        if let Value::Error(e) = num {
            return Value::Error(*e);
        }
        if let Value::Error(e) = base {
            return Value::Error(*e);
        }
        let n = match num.coerce_to_number() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let b = match base.coerce_to_number() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        match crate::functions::math::log(n, Some(b)) {
            Ok(out) => Value::Number(out),
            Err(e) => Value::Error(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "LOG10",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: log10_fn,
    }
}

fn log10_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[0]));
    elementwise_unary(&value, |elem| match elem {
        Value::Error(e) => Value::Error(*e),
        other => {
            let n = match other.coerce_to_number() {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            match crate::functions::math::log10(n) {
                Ok(out) => Value::Number(out),
                Err(e) => Value::Error(excel_error_kind(e)),
            }
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "SQRT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: sqrt_fn,
    }
}

fn sqrt_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[0]));
    elementwise_unary(&value, |elem| match elem {
        Value::Error(e) => Value::Error(*e),
        other => {
            let n = match other.coerce_to_number() {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            match crate::functions::math::sqrt(n) {
                Ok(out) => Value::Number(out),
                Err(e) => Value::Error(excel_error_kind(e)),
            }
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "POWER",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: power_fn,
    }
}

fn power_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[0]));
    let power = dynamic_value_from_arg(ctx, ctx.eval_arg(&args[1]));
    elementwise_binary(&number, &power, |num, power| {
        if let Value::Error(e) = num {
            return Value::Error(*e);
        }
        if let Value::Error(e) = power {
            return Value::Error(*e);
        }
        let n = match num.coerce_to_number() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let p = match power.coerce_to_number() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        match crate::functions::math::power(n, p) {
            Ok(out) => Value::Number(out),
            Err(e) => Value::Error(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "PRODUCT",
        min_args: 0,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: product_fn,
    }
}

fn product_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut values: Vec<f64> = Vec::new();

    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => match v {
                Value::Error(e) => return Value::Error(e),
                Value::Number(n) => values.push(n),
                Value::Bool(b) => values.push(if b { 1.0 } else { 0.0 }),
                Value::Blank => {}
                Value::Text(s) => match Value::Text(s).coerce_to_number() {
                    Ok(n) => values.push(n),
                    Err(e) => return Value::Error(e),
                },
                Value::Reference(_) | Value::ReferenceUnion(_) => {
                    return Value::Error(ErrorKind::Value)
                }
                Value::Array(arr) => {
                    for v in arr.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => values.push(*n),
                            Value::Bool(_)
                            | Value::Text(_)
                            | Value::Blank
                            | Value::Reference(_)
                            | Value::ReferenceUnion(_)
                            | Value::Array(_)
                            | Value::Lambda(_)
                            | Value::Spill { .. } => {}
                        }
                    }
                }
                Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            },
            ArgValue::Reference(r) => {
                for addr in ctx.iter_reference_cells(&r) {
                    match ctx.get_cell_value(&r.sheet_id, addr) {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => values.push(n),
                        // Excel quirk: logicals/text in references are ignored.
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Blank
                        | Value::Reference(_)
                        | Value::ReferenceUnion(_)
                        | Value::Array(_)
                        | Value::Lambda(_)
                        | Value::Spill { .. } => {}
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
                            Value::Number(n) => values.push(n),
                            Value::Bool(_)
                            | Value::Text(_)
                            | Value::Blank
                            | Value::Reference(_)
                            | Value::ReferenceUnion(_)
                            | Value::Array(_)
                            | Value::Lambda(_)
                            | Value::Spill { .. } => {}
                        }
                    }
                }
            }
        }
    }

    match crate::functions::math::product(&values) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "CEILING",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: ceiling_fn,
    }
}

fn ceiling_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = match scalar_from_arg(ctx, ctx.eval_arg(&args[0])).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let significance = match scalar_from_arg(ctx, ctx.eval_arg(&args[1])).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::math::ceiling(number, significance) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "FLOOR",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: floor_fn,
    }
}

fn floor_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = match scalar_from_arg(ctx, ctx.eval_arg(&args[0])).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let significance = match scalar_from_arg(ctx, ctx.eval_arg(&args[1])).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match crate::functions::math::floor(number, significance) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "CEILING.MATH",
        min_args: 1,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: ceiling_math_fn,
    }
}

fn ceiling_math_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = match scalar_from_arg(ctx, ctx.eval_arg(&args[0])).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let significance = if args.len() >= 2 {
        match scalar_from_arg(ctx, ctx.eval_arg(&args[1])).coerce_to_number() {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };
    let mode = if args.len() >= 3 {
        match scalar_from_arg(ctx, ctx.eval_arg(&args[2])).coerce_to_number() {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };
    match crate::functions::math::ceiling_math(number, significance, mode) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "FLOOR.MATH",
        min_args: 1,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: floor_math_fn,
    }
}

fn floor_math_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = match scalar_from_arg(ctx, ctx.eval_arg(&args[0])).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let significance = if args.len() >= 2 {
        match scalar_from_arg(ctx, ctx.eval_arg(&args[1])).coerce_to_number() {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };
    let mode = if args.len() >= 3 {
        match scalar_from_arg(ctx, ctx.eval_arg(&args[2])).coerce_to_number() {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };
    match crate::functions::math::floor_math(number, significance, mode) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "CEILING.PRECISE",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: ceiling_precise_fn,
    }
}

fn ceiling_precise_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = match scalar_from_arg(ctx, ctx.eval_arg(&args[0])).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let significance = if args.len() == 2 {
        match scalar_from_arg(ctx, ctx.eval_arg(&args[1])).coerce_to_number() {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };
    match crate::functions::math::ceiling_precise(number, significance) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "FLOOR.PRECISE",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: floor_precise_fn,
    }
}

fn floor_precise_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = match scalar_from_arg(ctx, ctx.eval_arg(&args[0])).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let significance = if args.len() == 2 {
        match scalar_from_arg(ctx, ctx.eval_arg(&args[1])).coerce_to_number() {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };
    match crate::functions::math::floor_precise(number, significance) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "ISO.CEILING",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: iso_ceiling_fn,
    }
}

fn iso_ceiling_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = match scalar_from_arg(ctx, ctx.eval_arg(&args[0])).coerce_to_number() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let significance = if args.len() == 2 {
        match scalar_from_arg(ctx, ctx.eval_arg(&args[1])).coerce_to_number() {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };
    match crate::functions::math::iso_ceiling(number, significance) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
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
    let criteria_range = match values_from_range_arg(ctx, ctx.eval_arg(&args[0])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let criteria = scalar_from_arg(ctx, ctx.eval_arg(&args[1]));
    let sum_range = if args.len() == 3 {
        match values_from_range_arg(ctx, ctx.eval_arg(&args[2])) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    match crate::functions::math::sumif(
        &criteria_range,
        &criteria,
        sum_range.as_deref(),
        ctx.date_system(),
    ) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(e),
    }
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

    let sum_range = match values_from_range_arg(ctx, ctx.eval_arg(&args[0])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let mut ranges: Vec<Vec<Value>> = Vec::new();
    let mut criteria_values: Vec<Value> = Vec::new();
    for pair in args[1..].chunks(2) {
        let range_values = match values_from_range_arg(ctx, ctx.eval_arg(&pair[0])) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        ranges.push(range_values);
        criteria_values.push(scalar_from_arg(ctx, ctx.eval_arg(&pair[1])));
    }

    let mut pairs: Vec<(&[Value], &Value)> = Vec::with_capacity(ranges.len());
    for (range, crit) in ranges.iter().zip(criteria_values.iter()) {
        pairs.push((range.as_slice(), crit));
    }

    match crate::functions::math::sumifs(&sum_range, &pairs, ctx.date_system()) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(e),
    }
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
    let criteria_range = match values_from_range_arg(ctx, ctx.eval_arg(&args[0])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let criteria = scalar_from_arg(ctx, ctx.eval_arg(&args[1]));
    let average_range = if args.len() == 3 {
        match values_from_range_arg(ctx, ctx.eval_arg(&args[2])) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    match crate::functions::math::averageif(
        &criteria_range,
        &criteria,
        average_range.as_deref(),
        ctx.date_system(),
    ) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(e),
    }
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

    let average_range = match values_from_range_arg(ctx, ctx.eval_arg(&args[0])) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let mut ranges: Vec<Vec<Value>> = Vec::new();
    let mut criteria_values: Vec<Value> = Vec::new();
    for pair in args[1..].chunks(2) {
        let range_values = match values_from_range_arg(ctx, ctx.eval_arg(&pair[0])) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        ranges.push(range_values);
        criteria_values.push(scalar_from_arg(ctx, ctx.eval_arg(&pair[1])));
    }

    let mut pairs: Vec<(&[Value], &Value)> = Vec::with_capacity(ranges.len());
    for (range, crit) in ranges.iter().zip(criteria_values.iter()) {
        pairs.push((range.as_slice(), crit));
    }

    match crate::functions::math::averageifs(&average_range, &pairs, ctx.date_system()) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "SUBTOTAL",
        min_args: 2,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Any],
        implementation: subtotal_fn,
    }
}

fn subtotal_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let function_num = match scalar_from_arg(ctx, ctx.eval_arg(&args[0])).coerce_to_i64() {
        Ok(v) => match i32::try_from(v) {
            Ok(v) => v,
            Err(_) => return Value::Error(ErrorKind::Value),
        },
        Err(e) => return Value::Error(e),
    };

    let mut values = Vec::new();
    for arg in &args[1..] {
        append_values_for_aggregate(ctx, ctx.eval_arg(arg), &mut values);
    }

    match crate::functions::math::subtotal(function_num, &values) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "AGGREGATE",
        min_args: 3,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Any],
        implementation: aggregate_fn,
    }
}

fn aggregate_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let function_num = match scalar_from_arg(ctx, ctx.eval_arg(&args[0])).coerce_to_i64() {
        Ok(v) => match i32::try_from(v) {
            Ok(v) => v,
            Err(_) => return Value::Error(ErrorKind::Value),
        },
        Err(e) => return Value::Error(e),
    };
    let options = match scalar_from_arg(ctx, ctx.eval_arg(&args[1])).coerce_to_i64() {
        Ok(v) => match i32::try_from(v) {
            Ok(v) => v,
            Err(_) => return Value::Error(ErrorKind::Value),
        },
        Err(e) => return Value::Error(e),
    };

    let mut values = Vec::new();
    for arg in &args[2..] {
        append_values_for_aggregate(ctx, ctx.eval_arg(arg), &mut values);
    }

    match crate::functions::math::aggregate(function_num, options, &values) {
        Ok(out) => Value::Number(out),
        Err(e) => Value::Error(e),
    }
}
