use std::collections::HashSet;

use crate::error::ExcelError;
use crate::eval::{CellAddr, CompiledExpr};
use crate::functions::{
    array_lift, ArgValue, ArraySupport, FunctionContext, FunctionSpec, Reference,
};
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
            let r = r.normalized();
            ctx.record_reference(&r);
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
        Value::Array(arr) => {
            Value::Array(Array::new(arr.rows, arr.cols, arr.iter().map(f).collect()))
        }
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
            let n = match other.coerce_to_number_with_ctx(ctx) {
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
            let n = match other.coerce_to_number_with_ctx(ctx) {
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
            let n = match other.coerce_to_number_with_ctx(ctx) {
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
            let n = match other.coerce_to_number_with_ctx(ctx) {
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
            let n = match other.coerce_to_number_with_ctx(ctx) {
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
            let n = match other.coerce_to_number_with_ctx(ctx) {
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
        let x = match x_num.coerce_to_number_with_ctx(ctx) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let y = match y_num.coerce_to_number_with_ctx(ctx) {
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
            let n = match other.coerce_to_number_with_ctx(ctx) {
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
            let n = match other.coerce_to_number_with_ctx(ctx) {
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
    // Excel treats omitted optional arguments as the function's default. In formulas like
    // `LOG(100,)` the base expression is parsed as a blank argument (`CompiledExpr::Blank`), which
    // should be treated the same as an omitted base (`LOG(100)`).
    if matches!(args.get(1), None | Some(CompiledExpr::Blank)) {
        return elementwise_unary(&number, |elem| match elem {
            Value::Error(e) => Value::Error(*e),
            other => {
                let n = match other.coerce_to_number_with_ctx(ctx) {
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
        let n = match num.coerce_to_number_with_ctx(ctx) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let b = match base.coerce_to_number_with_ctx(ctx) {
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
            let n = match other.coerce_to_number_with_ctx(ctx) {
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
            let n = match other.coerce_to_number_with_ctx(ctx) {
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
        let n = match num.coerce_to_number_with_ctx(ctx) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let p = match power.coerce_to_number_with_ctx(ctx) {
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
                Value::Text(s) => match Value::Text(s).coerce_to_number_with_ctx(ctx) {
                    Ok(n) => values.push(n),
                    Err(e) => return Value::Error(e),
                },
                Value::Entity(_) | Value::Record(_) => return Value::Error(ErrorKind::Value),
                Value::Reference(_) | Value::ReferenceUnion(_) => {
                    return Value::Error(ErrorKind::Value)
                }
                Value::Array(arr) => {
                    for v in arr.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => values.push(*n),
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                            Value::Bool(_)
                            | Value::Text(_)
                            | Value::Entity(_)
                            | Value::Record(_)
                            | Value::Blank
                            | Value::Reference(_)
                            | Value::ReferenceUnion(_)
                            | Value::Array(_)
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
                        Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                        // Excel quirk: logicals/text in references are ignored.
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Entity(_)
                        | Value::Record(_)
                        | Value::Blank
                        | Value::Reference(_)
                        | Value::ReferenceUnion(_)
                        | Value::Array(_)
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
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                            Value::Bool(_)
                            | Value::Text(_)
                            | Value::Entity(_)
                            | Value::Record(_)
                            | Value::Blank
                            | Value::Reference(_)
                            | Value::ReferenceUnion(_)
                            | Value::Array(_)
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
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: ceiling_fn,
    }
}

fn ceiling_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    let significance = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(number, significance, |number, significance| {
        let number = number.coerce_to_number_with_ctx(ctx)?;
        let significance = significance.coerce_to_number_with_ctx(ctx)?;
        match crate::functions::math::ceiling(number, significance) {
            Ok(out) => Ok(Value::Number(out)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "FLOOR",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: floor_fn,
    }
}

fn floor_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    let significance = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(number, significance, |number, significance| {
        let number = number.coerce_to_number_with_ctx(ctx)?;
        let significance = significance.coerce_to_number_with_ctx(ctx)?;
        match crate::functions::math::floor(number, significance) {
            Ok(out) => Ok(Value::Number(out)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "CEILING.MATH",
        min_args: 1,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: ceiling_math_fn,
    }
}

fn ceiling_math_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    let significance = match args.get(1) {
        Some(CompiledExpr::Blank) | None => Value::Number(1.0),
        Some(expr) => array_lift::eval_arg(ctx, expr),
    };
    let mode = match args.get(2) {
        Some(CompiledExpr::Blank) | None => Value::Number(0.0),
        Some(expr) => array_lift::eval_arg(ctx, expr),
    };

    array_lift::lift3(number, significance, mode, |number, significance, mode| {
        let number = number.coerce_to_number_with_ctx(ctx)?;
        let significance = significance.coerce_to_number_with_ctx(ctx)?;
        let mode = mode.coerce_to_number_with_ctx(ctx)?;
        match crate::functions::math::ceiling_math(number, Some(significance), Some(mode)) {
            Ok(out) => Ok(Value::Number(out)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "FLOOR.MATH",
        min_args: 1,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: floor_math_fn,
    }
}

fn floor_math_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    let significance = match args.get(1) {
        Some(CompiledExpr::Blank) | None => Value::Number(1.0),
        Some(expr) => array_lift::eval_arg(ctx, expr),
    };
    let mode = match args.get(2) {
        Some(CompiledExpr::Blank) | None => Value::Number(0.0),
        Some(expr) => array_lift::eval_arg(ctx, expr),
    };

    array_lift::lift3(number, significance, mode, |number, significance, mode| {
        let number = number.coerce_to_number_with_ctx(ctx)?;
        let significance = significance.coerce_to_number_with_ctx(ctx)?;
        let mode = mode.coerce_to_number_with_ctx(ctx)?;
        match crate::functions::math::floor_math(number, Some(significance), Some(mode)) {
            Ok(out) => Ok(Value::Number(out)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "CEILING.PRECISE",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: ceiling_precise_fn,
    }
}

fn ceiling_precise_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    let significance = match args.get(1) {
        Some(CompiledExpr::Blank) | None => Value::Number(1.0),
        Some(expr) => array_lift::eval_arg(ctx, expr),
    };

    array_lift::lift2(number, significance, |number, significance| {
        let number = number.coerce_to_number_with_ctx(ctx)?;
        let significance = significance.coerce_to_number_with_ctx(ctx)?;
        match crate::functions::math::ceiling_precise(number, Some(significance)) {
            Ok(out) => Ok(Value::Number(out)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "FLOOR.PRECISE",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: floor_precise_fn,
    }
}

fn floor_precise_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    let significance = match args.get(1) {
        Some(CompiledExpr::Blank) | None => Value::Number(1.0),
        Some(expr) => array_lift::eval_arg(ctx, expr),
    };

    array_lift::lift2(number, significance, |number, significance| {
        let number = number.coerce_to_number_with_ctx(ctx)?;
        let significance = significance.coerce_to_number_with_ctx(ctx)?;
        match crate::functions::math::floor_precise(number, Some(significance)) {
            Ok(out) => Ok(Value::Number(out)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "ISO.CEILING",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: iso_ceiling_fn,
    }
}

fn iso_ceiling_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    let significance = match args.get(1) {
        Some(CompiledExpr::Blank) | None => Value::Number(1.0),
        Some(expr) => array_lift::eval_arg(ctx, expr),
    };

    array_lift::lift2(number, significance, |number, significance| {
        let number = number.coerce_to_number_with_ctx(ctx)?;
        let significance = significance.coerce_to_number_with_ctx(ctx)?;
        match crate::functions::math::iso_ceiling(number, Some(significance)) {
            Ok(out) => Ok(Value::Number(out)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
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
    let function_num =
        match scalar_from_arg(ctx, ctx.eval_arg(&args[0])).coerce_to_i64_with_ctx(ctx) {
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
    let function_num =
        match scalar_from_arg(ctx, ctx.eval_arg(&args[0])).coerce_to_i64_with_ctx(ctx) {
            Ok(v) => match i32::try_from(v) {
                Ok(v) => v,
                Err(_) => return Value::Error(ErrorKind::Value),
            },
            Err(e) => return Value::Error(e),
        };
    let options = match scalar_from_arg(ctx, ctx.eval_arg(&args[1])).coerce_to_i64_with_ctx(ctx) {
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

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
