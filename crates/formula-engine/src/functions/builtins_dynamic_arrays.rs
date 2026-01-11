use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};

use crate::eval::{CellAddr, CompiledExpr};
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{Array, ErrorKind, Value};

inventory::submit! {
    FunctionSpec {
        name: "FILTER",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: filter_fn,
    }
}

fn filter_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match eval_array_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let include = match eval_array_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let filters_rows = include.rows == array.rows && include.cols == 1;
    let filters_cols = include.rows == 1 && include.cols == array.cols;

    if !filters_rows && !filters_cols {
        return Value::Error(ErrorKind::Value);
    }

    if filters_rows {
        let mut keep = Vec::with_capacity(array.rows);
        for row in 0..array.rows {
            let v = include.get(row, 0).cloned().unwrap_or(Value::Blank);
            match v.coerce_to_bool() {
                Ok(b) => keep.push(b),
                Err(e) => return Value::Error(e),
            }
        }

        let out_rows = keep.iter().filter(|&&b| b).count();
        if out_rows == 0 {
            return filter_if_empty(ctx, args.get(2));
        }

        let mut values = Vec::with_capacity(out_rows.saturating_mul(array.cols));
        for (row, keep_row) in keep.into_iter().enumerate() {
            if !keep_row {
                continue;
            }
            for col in 0..array.cols {
                values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
            }
        }

        return Value::Array(Array::new(out_rows, array.cols, values));
    }

    // Filter columns.
    let mut keep = Vec::with_capacity(array.cols);
    for col in 0..array.cols {
        let v = include.get(0, col).cloned().unwrap_or(Value::Blank);
        match v.coerce_to_bool() {
            Ok(b) => keep.push(b),
            Err(e) => return Value::Error(e),
        }
    }

    let out_cols = keep.iter().filter(|&&b| b).count();
    if out_cols == 0 {
        return filter_if_empty(ctx, args.get(2));
    }

    let mut values = Vec::with_capacity(array.rows.saturating_mul(out_cols));
    for row in 0..array.rows {
        for (col, keep_col) in keep.iter().copied().enumerate() {
            if keep_col {
                values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
            }
        }
    }

    Value::Array(Array::new(array.rows, out_cols, values))
}

fn filter_if_empty(ctx: &dyn FunctionContext, expr: Option<&CompiledExpr>) -> Value {
    let Some(expr) = expr else {
        return Value::Error(ErrorKind::Calc);
    };

    match arg_value_to_array(ctx, ctx.eval_arg(expr)) {
        Ok(arr) => Value::Array(arr),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "SORT",
        min_args: 1,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: sort_fn,
    }
}

fn sort_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match eval_array_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let sort_index = if args.len() >= 2 {
        match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        1
    };

    let sort_order = if args.len() >= 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_i64() {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        1
    };

    let by_col = if args.len() >= 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_bool() {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        false
    };

    let descending = match sort_order {
        1 => false,
        -1 => true,
        _ => return Value::Error(ErrorKind::Value),
    };

    if by_col {
        let max_index = i64::try_from(array.rows).unwrap_or(i64::MAX);
        if sort_index < 1 || sort_index > max_index {
            return Value::Error(ErrorKind::Value);
        }
        let key_row = (sort_index - 1) as usize;

        let mut keys = Vec::with_capacity(array.cols);
        for col in 0..array.cols {
            let v = array.get(key_row, col).unwrap_or(&Value::Blank);
            keys.push(sort_key(v));
        }

        let mut order: Vec<usize> = (0..array.cols).collect();
        order.sort_by(|&a, &b| {
            let ord = compare_sort_keys(&keys[a], &keys[b], descending);
            if ord == Ordering::Equal {
                a.cmp(&b)
            } else {
                ord
            }
        });

        let mut values = Vec::with_capacity(array.rows.saturating_mul(array.cols));
        for row in 0..array.rows {
            for &col in &order {
                values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
            }
        }

        return Value::Array(Array::new(array.rows, array.cols, values));
    }

    let max_index = i64::try_from(array.cols).unwrap_or(i64::MAX);
    if sort_index < 1 || sort_index > max_index {
        return Value::Error(ErrorKind::Value);
    }
    let key_col = (sort_index - 1) as usize;

    let mut keys = Vec::with_capacity(array.rows);
    for row in 0..array.rows {
        let v = array.get(row, key_col).unwrap_or(&Value::Blank);
        keys.push(sort_key(v));
    }

    let mut order: Vec<usize> = (0..array.rows).collect();
    order.sort_by(|&a, &b| {
        let ord = compare_sort_keys(&keys[a], &keys[b], descending);
        if ord == Ordering::Equal {
            a.cmp(&b)
        } else {
            ord
        }
    });

    let mut values = Vec::with_capacity(array.rows.saturating_mul(array.cols));
    for &row in &order {
        for col in 0..array.cols {
            values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
        }
    }

    Value::Array(Array::new(array.rows, array.cols, values))
}

#[derive(Debug, Clone)]
enum SortKeyValue {
    Number(f64),
    Text(String),
    Bool(bool),
    Error(ErrorKind),
    Blank,
}

impl SortKeyValue {
    fn kind_rank(&self) -> u8 {
        match self {
            SortKeyValue::Number(_) => 0,
            SortKeyValue::Text(_) => 1,
            SortKeyValue::Bool(_) => 2,
            SortKeyValue::Error(_) => 3,
            SortKeyValue::Blank => 4,
        }
    }
}

fn sort_key(value: &Value) -> SortKeyValue {
    match value {
        Value::Number(n) => SortKeyValue::Number(*n),
        Value::Text(s) => SortKeyValue::Text(s.to_lowercase()),
        Value::Bool(b) => SortKeyValue::Bool(*b),
        Value::Blank => SortKeyValue::Blank,
        Value::Error(e) => SortKeyValue::Error(*e),
        Value::Array(_) | Value::Lambda(_) | Value::Spill { .. } => {
            SortKeyValue::Error(ErrorKind::Value)
        }
    }
}

fn error_rank(error: ErrorKind) -> u8 {
    match error {
        ErrorKind::Null => 0,
        ErrorKind::Div0 => 1,
        ErrorKind::Value => 2,
        ErrorKind::Ref => 3,
        ErrorKind::Name => 4,
        ErrorKind::Num => 5,
        ErrorKind::NA => 6,
        ErrorKind::Spill => 7,
        ErrorKind::Calc => 8,
    }
}

fn compare_sort_keys(a: &SortKeyValue, b: &SortKeyValue, descending: bool) -> Ordering {
    let rank_cmp = a.kind_rank().cmp(&b.kind_rank());
    if rank_cmp != Ordering::Equal {
        // Excel keeps a fixed cross-type ordering (numbers, then text, then booleans, then errors,
        // with blanks last). Direction is applied within the same type.
        return rank_cmp;
    }

    let ord = match (a, b) {
        (SortKeyValue::Blank, SortKeyValue::Blank) => Ordering::Equal,
        (SortKeyValue::Number(a), SortKeyValue::Number(b)) => a.total_cmp(b),
        (SortKeyValue::Text(a), SortKeyValue::Text(b)) => a.cmp(b),
        (SortKeyValue::Bool(a), SortKeyValue::Bool(b)) => a.cmp(b),
        (SortKeyValue::Error(a), SortKeyValue::Error(b)) => error_rank(*a).cmp(&error_rank(*b)),
        _ => Ordering::Equal,
    };

    if descending {
        ord.reverse()
    } else {
        ord
    }
}

inventory::submit! {
    FunctionSpec {
        name: "UNIQUE",
        min_args: 1,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Bool, ValueType::Bool],
        implementation: unique_fn,
    }
}

fn unique_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match eval_array_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let by_col = if args.len() >= 2 {
        match eval_scalar_arg(ctx, &args[1]).coerce_to_bool() {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        false
    };

    let exactly_once = if args.len() >= 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_bool() {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        false
    };

    if by_col {
        unique_columns(array, exactly_once)
    } else {
        unique_rows(array, exactly_once)
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
enum UniqueKeyCell {
    Blank,
    Bool(bool),
    Number(u64),
    Text(String),
    Error(ErrorKind),
}

fn unique_key_cell(value: &Value) -> UniqueKeyCell {
    match value {
        Value::Blank => UniqueKeyCell::Blank,
        Value::Bool(b) => UniqueKeyCell::Bool(*b),
        Value::Number(n) => UniqueKeyCell::Number(canonical_number_bits(*n)),
        Value::Text(s) if s.is_empty() => UniqueKeyCell::Blank,
        Value::Text(s) => UniqueKeyCell::Text(s.to_lowercase()),
        Value::Error(e) => UniqueKeyCell::Error(*e),
        Value::Array(_) | Value::Lambda(_) | Value::Spill { .. } => UniqueKeyCell::Error(ErrorKind::Value),
    }
}

fn canonical_number_bits(n: f64) -> u64 {
    if n == 0.0 {
        return 0f64.to_bits();
    }
    if n.is_nan() {
        return f64::NAN.to_bits();
    }
    n.to_bits()
}

fn unique_rows(array: Array, exactly_once: bool) -> Value {
    let mut counts: HashMap<Vec<UniqueKeyCell>, usize> = HashMap::new();
    let mut keys_by_row: Vec<Vec<UniqueKeyCell>> = Vec::with_capacity(array.rows);

    for row in 0..array.rows {
        let mut key = Vec::with_capacity(array.cols);
        for col in 0..array.cols {
            let v = array.get(row, col).unwrap_or(&Value::Blank);
            key.push(unique_key_cell(v));
        }
        *counts.entry(key.clone()).or_insert(0) += 1;
        keys_by_row.push(key);
    }

    let mut selected: Vec<usize> = Vec::new();
    if exactly_once {
        for (row, key) in keys_by_row.iter().enumerate() {
            if counts.get(key).copied().unwrap_or(0) == 1 {
                selected.push(row);
            }
        }
    } else {
        let mut seen: HashSet<Vec<UniqueKeyCell>> = HashSet::new();
        for (row, key) in keys_by_row.iter().enumerate() {
            if seen.insert(key.clone()) {
                selected.push(row);
            }
        }
    }

    if selected.is_empty() {
        return Value::Error(ErrorKind::Calc);
    }

    let out_rows = selected.len();
    let mut values = Vec::with_capacity(out_rows.saturating_mul(array.cols));
    for row in selected {
        for col in 0..array.cols {
            values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
        }
    }

    Value::Array(Array::new(out_rows, array.cols, values))
}

fn unique_columns(array: Array, exactly_once: bool) -> Value {
    let mut counts: HashMap<Vec<UniqueKeyCell>, usize> = HashMap::new();
    let mut keys_by_col: Vec<Vec<UniqueKeyCell>> = Vec::with_capacity(array.cols);

    for col in 0..array.cols {
        let mut key = Vec::with_capacity(array.rows);
        for row in 0..array.rows {
            let v = array.get(row, col).unwrap_or(&Value::Blank);
            key.push(unique_key_cell(v));
        }
        *counts.entry(key.clone()).or_insert(0) += 1;
        keys_by_col.push(key);
    }

    let mut selected: Vec<usize> = Vec::new();
    if exactly_once {
        for (col, key) in keys_by_col.iter().enumerate() {
            if counts.get(key).copied().unwrap_or(0) == 1 {
                selected.push(col);
            }
        }
    } else {
        let mut seen: HashSet<Vec<UniqueKeyCell>> = HashSet::new();
        for (col, key) in keys_by_col.iter().enumerate() {
            if seen.insert(key.clone()) {
                selected.push(col);
            }
        }
    }

    if selected.is_empty() {
        return Value::Error(ErrorKind::Calc);
    }

    let mut values = Vec::with_capacity(array.rows.saturating_mul(selected.len()));
    for row in 0..array.rows {
        for &col in &selected {
            values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
        }
    }

    Value::Array(Array::new(array.rows, selected.len(), values))
}

fn eval_array_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<Array, ErrorKind> {
    arg_value_to_array(ctx, ctx.eval_arg(expr))
}

fn arg_value_to_array(ctx: &dyn FunctionContext, arg: ArgValue) -> Result<Array, ErrorKind> {
    match arg {
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
                    values.push(ctx.get_cell_value(r.sheet_id, CellAddr { row, col }));
                }
            }
            Ok(Array::new(rows, cols, values))
        }
        ArgValue::ReferenceUnion(_) => Err(ErrorKind::Value),
    }
}
