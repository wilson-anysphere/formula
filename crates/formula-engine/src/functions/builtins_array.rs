use crate::eval::{CellAddr, CompiledExpr};
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{Array, ErrorKind, Value};

inventory::submit! {
    FunctionSpec {
        name: "TRANSPOSE",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: transpose_fn,
    }
}

fn transpose_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let input = match ctx.eval_arg(&args[0]) {
        ArgValue::Scalar(v) => match v {
            Value::Array(arr) => arr,
            Value::Error(e) => return Value::Error(e),
            other => Array::new(1, 1, vec![other]),
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            let rows = (r.end.row - r.start.row + 1) as usize;
            let cols = (r.end.col - r.start.col + 1) as usize;
            let mut values = Vec::with_capacity(rows.saturating_mul(cols));
            for row in r.start.row..=r.end.row {
                for col in r.start.col..=r.end.col {
                    values.push(ctx.get_cell_value(r.sheet_id, CellAddr { row, col }));
                }
            }
            Array::new(rows, cols, values)
        }
    };

    let mut out_values = Vec::with_capacity(input.rows.saturating_mul(input.cols));
    for r in 0..input.cols {
        for c in 0..input.rows {
            out_values.push(input.get(c, r).cloned().unwrap_or(Value::Blank));
        }
    }

    Value::Array(Array::new(input.cols, input.rows, out_values))
}

inventory::submit! {
    FunctionSpec {
        name: "SEQUENCE",
        min_args: 1,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: sequence_fn,
    }
}

fn sequence_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rows = match eval_scalar_arg(ctx, &args[0]).coerce_to_i64() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let cols = if args.len() >= 2 {
        match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        1
    };
    let start = if args.len() >= 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_number() {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        1.0
    };
    let step = if args.len() >= 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_number() {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        1.0
    };

    if rows <= 0 || cols <= 0 {
        return Value::Error(ErrorKind::Value);
    }

    let rows_usize = match usize::try_from(rows) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Num),
    };
    let cols_usize = match usize::try_from(cols) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Num),
    };

    let total = match rows_usize.checked_mul(cols_usize) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Num),
    };

    let mut values = Vec::with_capacity(total);
    for idx in 0..total {
        values.push(Value::Number(start + step * (idx as f64)));
    }

    Value::Array(Array::new(rows_usize, cols_usize, values))
}
