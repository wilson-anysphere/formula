use crate::eval::{CellAddr, CompiledExpr};
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{Array, ErrorKind, Value};

const VAR_ARGS: usize = 255;

inventory::submit! {
    FunctionSpec {
        name: "HSTACK",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: hstack_fn,
    }
}

fn hstack_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut arrays = Vec::with_capacity(args.len());
    let mut out_rows = 0usize;
    let mut out_cols = 0usize;

    for arg in args {
        let arr = match arg_to_array(ctx, arg) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        out_rows = out_rows.max(arr.rows);
        out_cols = out_cols.saturating_add(arr.cols);
        arrays.push(arr);
    }

    let total = match out_rows.checked_mul(out_cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Num),
    };
    let mut values = vec![Value::Error(ErrorKind::NA); total];

    let mut col_offset = 0usize;
    for arr in arrays {
        for row in 0..arr.rows {
            for col in 0..arr.cols {
                let dst_col = col_offset + col;
                let dst_idx = row * out_cols + dst_col;
                values[dst_idx] = arr.get(row, col).cloned().unwrap_or(Value::Blank);
            }
        }
        col_offset = col_offset.saturating_add(arr.cols);
    }

    Value::Array(Array::new(out_rows, out_cols, values))
}

inventory::submit! {
    FunctionSpec {
        name: "VSTACK",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: vstack_fn,
    }
}

fn vstack_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut arrays = Vec::with_capacity(args.len());
    let mut out_rows = 0usize;
    let mut out_cols = 0usize;

    for arg in args {
        let arr = match arg_to_array(ctx, arg) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        out_rows = out_rows.saturating_add(arr.rows);
        out_cols = out_cols.max(arr.cols);
        arrays.push(arr);
    }

    let total = match out_rows.checked_mul(out_cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Num),
    };
    let mut values = vec![Value::Error(ErrorKind::NA); total];

    let mut row_offset = 0usize;
    for arr in arrays {
        for row in 0..arr.rows {
            for col in 0..arr.cols {
                let dst_row = row_offset + row;
                let dst_idx = dst_row * out_cols + col;
                values[dst_idx] = arr.get(row, col).cloned().unwrap_or(Value::Blank);
            }
        }
        row_offset = row_offset.saturating_add(arr.rows);
    }

    Value::Array(Array::new(out_rows, out_cols, values))
}

inventory::submit! {
    FunctionSpec {
        name: "TOCOL",
        min_args: 1,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Bool],
        implementation: tocol_fn,
    }
}

fn tocol_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match arg_to_array(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let ignore = if args.len() >= 2 {
        match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        0
    };

    let scan_by_column = if args.len() >= 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_bool() {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        false
    };

    let ignore_mode = match ignore {
        0 | 1 | 2 | 3 => ignore,
        _ => return Value::Error(ErrorKind::Value),
    };

    let mut values = Vec::with_capacity(array.values.len());
    if scan_by_column {
        for col in 0..array.cols {
            for row in 0..array.rows {
                let v = array.get(row, col).cloned().unwrap_or(Value::Blank);
                if should_keep_value(&v, ignore_mode) {
                    values.push(v);
                }
            }
        }
    } else {
        for row in 0..array.rows {
            for col in 0..array.cols {
                let v = array.get(row, col).cloned().unwrap_or(Value::Blank);
                if should_keep_value(&v, ignore_mode) {
                    values.push(v);
                }
            }
        }
    }

    if values.is_empty() {
        return Value::Error(ErrorKind::Calc);
    }

    Value::Array(Array::new(values.len(), 1, values))
}

inventory::submit! {
    FunctionSpec {
        name: "TOROW",
        min_args: 1,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Bool],
        implementation: torow_fn,
    }
}

fn torow_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match arg_to_array(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let ignore = if args.len() >= 2 {
        match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        0
    };

    let scan_by_column = if args.len() >= 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_bool() {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        false
    };

    let ignore_mode = match ignore {
        0 | 1 | 2 | 3 => ignore,
        _ => return Value::Error(ErrorKind::Value),
    };

    let mut values = Vec::with_capacity(array.values.len());
    if scan_by_column {
        for col in 0..array.cols {
            for row in 0..array.rows {
                let v = array.get(row, col).cloned().unwrap_or(Value::Blank);
                if should_keep_value(&v, ignore_mode) {
                    values.push(v);
                }
            }
        }
    } else {
        for row in 0..array.rows {
            for col in 0..array.cols {
                let v = array.get(row, col).cloned().unwrap_or(Value::Blank);
                if should_keep_value(&v, ignore_mode) {
                    values.push(v);
                }
            }
        }
    }

    if values.is_empty() {
        return Value::Error(ErrorKind::Calc);
    }

    Value::Array(Array::new(1, values.len(), values))
}

inventory::submit! {
    FunctionSpec {
        name: "WRAPROWS",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Any],
        implementation: wraprows_fn,
    }
}

fn wraprows_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match arg_to_array(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let wrap_count = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    if wrap_count <= 0 {
        return Value::Error(ErrorKind::Value);
    }
    let wrap_cols = match usize::try_from(wrap_count) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Num),
    };

    let pad_with = if args.len() >= 3 {
        eval_scalar_arg(ctx, &args[2])
    } else {
        Value::Error(ErrorKind::NA)
    };

    let values_in = array.values;
    let out_rows = values_in
        .len()
        .checked_add(wrap_cols.saturating_sub(1))
        .map(|v| v / wrap_cols)
        .unwrap_or(usize::MAX);

    let total = match out_rows.checked_mul(wrap_cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Num),
    };
    let mut values = vec![pad_with.clone(); total];

    for (idx, v) in values_in.into_iter().enumerate() {
        let row = idx / wrap_cols;
        let col = idx % wrap_cols;
        values[row * wrap_cols + col] = v;
    }

    Value::Array(Array::new(out_rows, wrap_cols, values))
}

inventory::submit! {
    FunctionSpec {
        name: "WRAPCOLS",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Any],
        implementation: wrapcols_fn,
    }
}

fn wrapcols_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match arg_to_array(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let wrap_count = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    if wrap_count <= 0 {
        return Value::Error(ErrorKind::Value);
    }
    let wrap_rows = match usize::try_from(wrap_count) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Num),
    };

    let pad_with = if args.len() >= 3 {
        eval_scalar_arg(ctx, &args[2])
    } else {
        Value::Error(ErrorKind::NA)
    };

    let values_in = array.values;
    let out_cols = values_in
        .len()
        .checked_add(wrap_rows.saturating_sub(1))
        .map(|v| v / wrap_rows)
        .unwrap_or(usize::MAX);

    let total = match wrap_rows.checked_mul(out_cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Num),
    };
    let mut values = vec![pad_with.clone(); total];

    for (idx, v) in values_in.into_iter().enumerate() {
        let row = idx % wrap_rows;
        let col = idx / wrap_rows;
        values[row * out_cols + col] = v;
    }

    Value::Array(Array::new(wrap_rows, out_cols, values))
}

fn arg_to_array(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<Array, ErrorKind> {
    match ctx.eval_arg(expr) {
        ArgValue::Scalar(v) => match v {
            Value::Array(arr) => Ok(arr),
            Value::Error(e) => Err(e),
            other => Ok(Array::new(1, 1, vec![other])),
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            let rows = (r.end.row - r.start.row + 1) as usize;
            let cols = (r.end.col - r.start.col + 1) as usize;
            let mut values = Vec::with_capacity(rows.saturating_mul(cols));
            for row in r.start.row..=r.end.row {
                for col in r.start.col..=r.end.col {
                    values.push(ctx.get_cell_value(&r.sheet_id, CellAddr { row, col }));
                }
            }
            Ok(Array::new(rows, cols, values))
        }
        ArgValue::ReferenceUnion(_) => Err(ErrorKind::Value),
    }
}

fn is_blank_like(value: &Value) -> bool {
    match value {
        Value::Blank => true,
        Value::Text(s) => s.is_empty(),
        _ => false,
    }
}

fn should_keep_value(value: &Value, ignore_mode: i64) -> bool {
    let is_blank = is_blank_like(value);
    let is_error = matches!(value, Value::Error(_));

    match ignore_mode {
        0 => true,
        1 => !is_blank,
        2 => !is_error,
        3 => !is_blank && !is_error,
        _ => true,
    }
}
