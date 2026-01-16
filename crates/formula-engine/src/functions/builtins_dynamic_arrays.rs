use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};

use crate::eval::{CellAddr, CompiledExpr, Expr, NameRef, SheetReference};
use crate::functions::{
    eval_scalar_arg, volatile_rand_u64_below, ArgValue, ArraySupport, FunctionContext, FunctionSpec,
};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{casefold, casefold_owned, Array, ErrorKind, Lambda, RecordValue, Value};

fn checked_array_cells(rows: usize, cols: usize) -> Result<usize, ErrorKind> {
    let total = rows.checked_mul(cols).ok_or(ErrorKind::Spill)?;
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        return Err(ErrorKind::Spill);
    }
    Ok(total)
}

fn try_vec_with_capacity<T>(len: usize) -> Result<Vec<T>, ErrorKind> {
    let mut out = Vec::new();
    out.try_reserve_exact(len).map_err(|_| ErrorKind::Num)?;
    Ok(out)
}

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
        let mut keep = match try_vec_with_capacity::<bool>(array.rows) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        for row in 0..array.rows {
            let v = include.get(row, 0).cloned().unwrap_or(Value::Blank);
            match v.coerce_to_bool_with_ctx(ctx) {
                Ok(b) => keep.push(b),
                Err(e) => return Value::Error(e),
            }
        }

        let out_rows = keep.iter().filter(|&&b| b).count();
        if out_rows == 0 {
            return filter_if_empty(ctx, args.get(2));
        }

        let total = match checked_array_cells(out_rows, array.cols) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let mut values = match try_vec_with_capacity::<Value>(total) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
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
    let mut keep = match try_vec_with_capacity::<bool>(array.cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for col in 0..array.cols {
        let v = include.get(0, col).cloned().unwrap_or(Value::Blank);
        match v.coerce_to_bool_with_ctx(ctx) {
            Ok(b) => keep.push(b),
            Err(e) => return Value::Error(e),
        }
    }

    let out_cols = keep.iter().filter(|&&b| b).count();
    if out_cols == 0 {
        return filter_if_empty(ctx, args.get(2));
    }

    let total = match checked_array_cells(array.rows, out_cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut values = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
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

    let by_col = match eval_optional_bool(ctx, args.get(3), false) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    if array.rows == 0 || array.cols == 0 {
        return Value::Array(array);
    }

    let key_count = if by_col { array.rows } else { array.cols };
    let sort_indices = match sort_vector_indices(ctx, args.get(1), key_count) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let descending_flags = match sort_vector_orders(ctx, args.get(2), sort_indices.len()) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    if by_col {
        let mut keys: Vec<Vec<SortKeyValue>> = match try_vec_with_capacity(sort_indices.len()) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        for &row_idx in &sort_indices {
            let mut out: Vec<SortKeyValue> = match try_vec_with_capacity(array.cols) {
                Ok(v) => v,
                Err(e) => return Value::Error(e),
            };
            for col in 0..array.cols {
                out.push(sort_key(
                    ctx,
                    array.get(row_idx, col).unwrap_or(&Value::Blank),
                ));
            }
            keys.push(out);
        }

        let mut order: Vec<usize> = (0..array.cols).collect();
        order.sort_by(|&a, &b| {
            for (key_idx, desc) in descending_flags.iter().copied().enumerate() {
                let ord = compare_sort_keys(&keys[key_idx][a], &keys[key_idx][b], desc);
                if ord != Ordering::Equal {
                    return ord;
                }
            }
            a.cmp(&b)
        });

        let total = match checked_array_cells(array.rows, array.cols) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let mut values = match try_vec_with_capacity::<Value>(total) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        for row in 0..array.rows {
            for &col in &order {
                values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
            }
        }

        return Value::Array(Array::new(array.rows, array.cols, values));
    }

    let mut keys: Vec<Vec<SortKeyValue>> = match try_vec_with_capacity(sort_indices.len()) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for &col_idx in &sort_indices {
        let mut out: Vec<SortKeyValue> = match try_vec_with_capacity(array.rows) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        for row in 0..array.rows {
            out.push(sort_key(
                ctx,
                array.get(row, col_idx).unwrap_or(&Value::Blank),
            ));
        }
        keys.push(out);
    }

    let mut order: Vec<usize> = (0..array.rows).collect();
    order.sort_by(|&a, &b| {
        for (key_idx, desc) in descending_flags.iter().copied().enumerate() {
            let ord = compare_sort_keys(&keys[key_idx][a], &keys[key_idx][b], desc);
            if ord != Ordering::Equal {
                return ord;
            }
        }
        a.cmp(&b)
    });

    let total = match checked_array_cells(array.rows, array.cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut values = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for &row in &order {
        for col in 0..array.cols {
            values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
        }
    }

    Value::Array(Array::new(array.rows, array.cols, values))
}

inventory::submit! {
    FunctionSpec {
        name: "SORTBY",
        min_args: 2,
        max_args: 255,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: sortby_fn,
    }
}

fn sortby_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match eval_array_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    #[derive(Clone, Copy)]
    enum VectorOrientation {
        Row,
        Column,
    }

    let vector_orientation = |arr: &Array, len: usize| -> Option<VectorOrientation> {
        if arr.rows == len && arr.cols == 1 {
            Some(VectorOrientation::Column)
        } else if arr.rows == 1 && arr.cols == len {
            Some(VectorOrientation::Row)
        } else {
            None
        }
    };

    let mut key_arrays = Vec::new();
    let mut descending_flags = Vec::new();

    // `SORTBY` arguments are `(array, by_array1, [sort_order1], [by_array2, [sort_order2]], ...)`.
    //
    // Excel disambiguates the optional `sort_orderN` arguments based on shape:
    // - If the next argument is a range/array, it is treated as the next `by_array`.
    // - Otherwise it is treated as `sort_order` for the current key.
    //
    // We must only evaluate each argument once to preserve volatility semantics. We also need to
    // preserve whether the argument was omitted (`Expr::Blank`) so that blank *values* can still
    // be distinguished from omitted optional `sort_order` arguments.
    let evaluated: Vec<(bool, ArgValue)> = args[1..]
        .iter()
        .map(|expr| (matches!(expr, Expr::Blank), ctx.eval_arg(expr)))
        .collect();

    let mut idx = 0usize;
    while idx < evaluated.len() {
        let (by_is_blank, by_value) = evaluated[idx].clone();
        idx += 1;

        if by_is_blank {
            return Value::Error(ErrorKind::Value);
        }

        let by_array = match arg_value_to_array(ctx, by_value) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };

        let mut desc = false;
        if idx < evaluated.len() && !arg_value_is_array_like(&evaluated[idx].1) {
            let (order_is_blank, order_arg) = evaluated[idx].clone();
            let order_value = match order_arg {
                ArgValue::Scalar(v) => v,
                ArgValue::Reference(r) => ctx.apply_implicit_intersection(&r),
                ArgValue::ReferenceUnion(_) => Value::Error(ErrorKind::Value),
            };
            if !order_is_blank {
                desc = match sort_descending_from_value(ctx, &order_value) {
                    Ok(v) => v,
                    Err(e) => return Value::Error(e),
                };
            }
            idx += 1;
        }

        key_arrays.push(by_array);
        descending_flags.push(desc);
    }

    let Some(first_key) = key_arrays.first() else {
        return Value::Error(ErrorKind::Value);
    };

    // Prefer row sorting when ambiguous (e.g. square arrays with 1xN keys).
    let sorts_rows = vector_orientation(first_key, array.rows).is_some();
    let sorts_cols = vector_orientation(first_key, array.cols).is_some();
    let sort_rows = if sorts_rows {
        true
    } else if sorts_cols {
        false
    } else {
        return Value::Error(ErrorKind::Value);
    };
    let axis_len = if sort_rows { array.rows } else { array.cols };

    let mut keys: Vec<Vec<SortKeyValue>> = match try_vec_with_capacity(key_arrays.len()) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for key in key_arrays {
        let Some(orientation) = vector_orientation(&key, axis_len) else {
            return Value::Error(ErrorKind::Value);
        };
        let mut out: Vec<SortKeyValue> = match try_vec_with_capacity(axis_len) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        for idx in 0..axis_len {
            let v = match orientation {
                VectorOrientation::Row => key.get(0, idx).unwrap_or(&Value::Blank),
                VectorOrientation::Column => key.get(idx, 0).unwrap_or(&Value::Blank),
            };
            out.push(sort_key(ctx, v));
        }
        keys.push(out);
    }

    let mut order: Vec<usize> = (0..axis_len).collect();
    order.sort_by(|&a, &b| {
        for (key_idx, desc) in descending_flags.iter().copied().enumerate() {
            let ord = compare_sort_keys(&keys[key_idx][a], &keys[key_idx][b], desc);
            if ord != Ordering::Equal {
                return ord;
            }
        }
        a.cmp(&b)
    });

    let total = match checked_array_cells(array.rows, array.cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut values = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    if sort_rows {
        for &row in &order {
            for col in 0..array.cols {
                values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
            }
        }
    } else {
        for row in 0..array.rows {
            for &col in &order {
                values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
            }
        }
    }

    Value::Array(Array::new(array.rows, array.cols, values))
}

#[derive(Debug, Clone)]
pub(super) enum SortKeyValue {
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

fn record_display_key_text(
    ctx: &dyn FunctionContext,
    record: &RecordValue,
) -> Result<Option<String>, ErrorKind> {
    let display = if let Some(display_field) = record.display_field.as_deref() {
        if let Some(value) = record.get_field_case_insensitive(display_field) {
            value.coerce_to_string_with_ctx(ctx)?
        } else {
            record.display.clone()
        }
    } else {
        record.display.clone()
    };

    if display.is_empty() {
        Ok(None)
    } else {
        Ok(Some(casefold_owned(display)))
    }
}

pub(super) fn sort_key(ctx: &dyn FunctionContext, value: &Value) -> SortKeyValue {
    match value {
        Value::Number(n) => SortKeyValue::Number(*n),
        Value::Text(s) => SortKeyValue::Text(casefold(s)),
        Value::Bool(b) => SortKeyValue::Bool(*b),
        Value::Entity(entity) if entity.display.is_empty() => SortKeyValue::Blank,
        Value::Entity(entity) => SortKeyValue::Text(casefold(&entity.display)),
        Value::Record(record) => match record_display_key_text(ctx, record) {
            Ok(None) => SortKeyValue::Blank,
            Ok(Some(s)) => SortKeyValue::Text(s),
            Err(e) => SortKeyValue::Error(e),
        },
        Value::Blank => SortKeyValue::Blank,
        Value::Error(e) => SortKeyValue::Error(*e),
        other => {
            // Treat any future scalar-ish values as their display string for stable ordering.
            // For non-scalar values, fall back to the existing #VALUE! behavior.
            let display = match other.coerce_to_string() {
                Ok(s) => s,
                Err(_) => return SortKeyValue::Error(ErrorKind::Value),
            };
            if display.is_empty() {
                SortKeyValue::Blank
            } else {
                SortKeyValue::Text(casefold_owned(display))
            }
        }
    }
}

fn error_rank(error: ErrorKind) -> u8 {
    error.code()
}

pub(super) fn compare_sort_keys(a: &SortKeyValue, b: &SortKeyValue, descending: bool) -> Ordering {
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
        match eval_scalar_arg(ctx, &args[1]).coerce_to_bool_with_ctx(ctx) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        false
    };

    let exactly_once = if args.len() >= 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_bool_with_ctx(ctx) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        false
    };

    if by_col {
        unique_columns(ctx, array, exactly_once)
    } else {
        unique_rows(ctx, array, exactly_once)
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

fn unique_key_cell(ctx: &dyn FunctionContext, value: &Value) -> UniqueKeyCell {
    match value {
        Value::Blank => UniqueKeyCell::Blank,
        Value::Bool(b) => UniqueKeyCell::Bool(*b),
        Value::Number(n) => UniqueKeyCell::Number(canonical_number_bits(*n)),
        Value::Text(s) if s.is_empty() => UniqueKeyCell::Blank,
        Value::Text(s) => UniqueKeyCell::Text(casefold(s)),
        Value::Entity(entity) if entity.display.is_empty() => UniqueKeyCell::Blank,
        Value::Entity(entity) => UniqueKeyCell::Text(casefold(&entity.display)),
        Value::Record(record) => match record_display_key_text(ctx, record) {
            Ok(None) => UniqueKeyCell::Blank,
            Ok(Some(s)) => UniqueKeyCell::Text(s),
            Err(e) => UniqueKeyCell::Error(e),
        },
        Value::Error(e) => UniqueKeyCell::Error(*e),
        Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Array(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => UniqueKeyCell::Error(ErrorKind::Value),
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

fn unique_rows(ctx: &dyn FunctionContext, array: Array, exactly_once: bool) -> Value {
    let mut counts: HashMap<Vec<UniqueKeyCell>, usize> = HashMap::new();
    let mut keys_by_row: Vec<Vec<UniqueKeyCell>> = match try_vec_with_capacity(array.rows) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    for row in 0..array.rows {
        let mut key: Vec<UniqueKeyCell> = match try_vec_with_capacity(array.cols) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        for col in 0..array.cols {
            let v = array.get(row, col).unwrap_or(&Value::Blank);
            key.push(unique_key_cell(ctx, v));
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
    let total = match checked_array_cells(out_rows, array.cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut values = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for row in selected {
        for col in 0..array.cols {
            values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
        }
    }

    Value::Array(Array::new(out_rows, array.cols, values))
}

fn unique_columns(ctx: &dyn FunctionContext, array: Array, exactly_once: bool) -> Value {
    let mut counts: HashMap<Vec<UniqueKeyCell>, usize> = HashMap::new();
    let mut keys_by_col: Vec<Vec<UniqueKeyCell>> = match try_vec_with_capacity(array.cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    for col in 0..array.cols {
        let mut key: Vec<UniqueKeyCell> = match try_vec_with_capacity(array.rows) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        for row in 0..array.rows {
            let v = array.get(row, col).unwrap_or(&Value::Blank);
            key.push(unique_key_cell(ctx, v));
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

    let total = match checked_array_cells(array.rows, selected.len()) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut values = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for row in 0..array.rows {
        for &col in &selected {
            values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
        }
    }

    Value::Array(Array::new(array.rows, selected.len(), values))
}

inventory::submit! {
    FunctionSpec {
        name: "TAKE",
        min_args: 1,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: take_fn,
    }
}

fn take_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match eval_array_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let default_rows = i64::try_from(array.rows).unwrap_or(i64::MAX);
    let rows = match eval_optional_i64(ctx, args.get(1), default_rows) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let cols = match eval_optional_i64(
        ctx,
        args.get(2),
        i64::try_from(array.cols).unwrap_or(i64::MAX),
    ) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let (row_start, out_rows) = take_span(array.rows, rows);
    let (col_start, out_cols) = take_span(array.cols, cols);
    if out_rows == 0 || out_cols == 0 {
        return Value::Error(ErrorKind::Calc);
    }

    let total = match checked_array_cells(out_rows, out_cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut values = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for r in 0..out_rows {
        for c in 0..out_cols {
            values.push(
                array
                    .get(row_start.saturating_add(r), col_start.saturating_add(c))
                    .cloned()
                    .unwrap_or(Value::Blank),
            );
        }
    }

    Value::Array(Array::new(out_rows, out_cols, values))
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}

inventory::submit! {
    FunctionSpec {
        name: "DROP",
        min_args: 1,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: drop_fn,
    }
}

fn drop_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match eval_array_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let rows = match eval_optional_i64(ctx, args.get(1), 0) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let cols = match eval_optional_i64(ctx, args.get(2), 0) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let (row_start, out_rows) = drop_span(array.rows, rows);
    let (col_start, out_cols) = drop_span(array.cols, cols);
    if out_rows == 0 || out_cols == 0 {
        return Value::Error(ErrorKind::Calc);
    }

    let total = match checked_array_cells(out_rows, out_cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut values = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for r in 0..out_rows {
        for c in 0..out_cols {
            values.push(
                array
                    .get(row_start.saturating_add(r), col_start.saturating_add(c))
                    .cloned()
                    .unwrap_or(Value::Blank),
            );
        }
    }

    Value::Array(Array::new(out_rows, out_cols, values))
}

inventory::submit! {
    FunctionSpec {
        name: "CHOOSECOLS",
        min_args: 2,
        max_args: 255,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Number],
        implementation: choosecols_fn,
    }
}

fn choosecols_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match eval_array_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let mut cols = Vec::with_capacity(args.len().saturating_sub(1));
    for expr in &args[1..] {
        let idx = match eval_scalar_arg(ctx, expr).coerce_to_i64_with_ctx(ctx) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let Some(col) = normalize_index(idx, array.cols) else {
            return Value::Error(ErrorKind::Value);
        };
        cols.push(col);
    }

    let total = match checked_array_cells(array.rows, cols.len()) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut values = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for row in 0..array.rows {
        for &col in &cols {
            values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
        }
    }

    Value::Array(Array::new(array.rows, cols.len(), values))
}

inventory::submit! {
    FunctionSpec {
        name: "CHOOSEROWS",
        min_args: 2,
        max_args: 255,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Number],
        implementation: chooserows_fn,
    }
}

fn chooserows_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match eval_array_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let mut rows = Vec::with_capacity(args.len().saturating_sub(1));
    for expr in &args[1..] {
        let idx = match eval_scalar_arg(ctx, expr).coerce_to_i64_with_ctx(ctx) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let Some(row) = normalize_index(idx, array.rows) else {
            return Value::Error(ErrorKind::Value);
        };
        rows.push(row);
    }

    let total = match checked_array_cells(rows.len(), array.cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut values = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for &row in &rows {
        for col in 0..array.cols {
            values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
        }
    }

    Value::Array(Array::new(rows.len(), array.cols, values))
}

inventory::submit! {
    FunctionSpec {
        name: "HSTACK",
        min_args: 1,
        max_args: 255,
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
    for expr in args {
        let array = match eval_array_arg(ctx, expr) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        arrays.push(array);
    }

    let out_rows = arrays.iter().map(|a| a.rows).max().unwrap_or(0);
    let out_cols: usize = arrays.iter().map(|a| a.cols).sum();

    let total = match checked_array_cells(out_rows, out_cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut values = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    values.resize(total, Value::Error(ErrorKind::NA));
    let mut col_offset = 0usize;
    for array in arrays {
        for row in 0..array.rows {
            for col in 0..array.cols {
                let idx = row
                    .saturating_mul(out_cols)
                    .saturating_add(col_offset.saturating_add(col));
                if let Some(dst) = values.get_mut(idx) {
                    *dst = array.get(row, col).cloned().unwrap_or(Value::Blank);
                }
            }
        }
        col_offset = col_offset.saturating_add(array.cols);
    }

    Value::Array(Array::new(out_rows, out_cols, values))
}

inventory::submit! {
    FunctionSpec {
        name: "VSTACK",
        min_args: 1,
        max_args: 255,
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
    for expr in args {
        let array = match eval_array_arg(ctx, expr) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        arrays.push(array);
    }

    let out_rows: usize = arrays.iter().map(|a| a.rows).sum();
    let out_cols = arrays.iter().map(|a| a.cols).max().unwrap_or(0);

    let total = match checked_array_cells(out_rows, out_cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut values = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    values.resize(total, Value::Error(ErrorKind::NA));
    let mut row_offset = 0usize;
    for array in arrays {
        for row in 0..array.rows {
            let dst_row = row_offset.saturating_add(row);
            for col in 0..array.cols {
                let idx = dst_row.saturating_mul(out_cols).saturating_add(col);
                if let Some(dst) = values.get_mut(idx) {
                    *dst = array.get(row, col).cloned().unwrap_or(Value::Blank);
                }
            }
        }
        row_offset = row_offset.saturating_add(array.rows);
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
    to_vector_fn(ctx, args, true)
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
    to_vector_fn(ctx, args, false)
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
    wrap_vector_fn(ctx, args, true)
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
    wrap_vector_fn(ctx, args, false)
}

inventory::submit! {
    FunctionSpec {
        name: "RANDARRAY",
        min_args: 0,
        max_args: 5,
        volatility: Volatility::Volatile,
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
        implementation: randarray_fn,
    }
}

fn randarray_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rows = match eval_optional_i64(ctx, args.get(0), 1) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let cols = match eval_optional_i64(ctx, args.get(1), 1) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let min = match eval_optional_number(ctx, args.get(2), 0.0) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let max = match eval_optional_number(ctx, args.get(3), 1.0) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let whole_number = match eval_optional_bool(ctx, args.get(4), false) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    if rows <= 0 || cols <= 0 {
        return Value::Error(ErrorKind::Value);
    }

    if !min.is_finite() || !max.is_finite() {
        return Value::Error(ErrorKind::Num);
    }
    if min > max {
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
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }

    // Use a fallible reservation to avoid aborting the process if a user asks for an
    // unreasonably large array (or if memory is exhausted).
    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    if whole_number {
        let low_f = min.ceil();
        let high_f = max.floor();
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

        for _ in 0..total {
            let offset = volatile_rand_u64_below(ctx, span) as i64;
            values.push(Value::Number((low + offset) as f64));
        }
    } else {
        if min == max {
            values.resize(total, Value::Number(min));
            return Value::Array(Array::new(rows_usize, cols_usize, values));
        }
        for _ in 0..total {
            let r = ctx.volatile_rand();
            // Generate a uniform float in [min, max) while avoiding potential overflow in
            // `(max - min)` for very large finite bounds.
            //
            // `r` is guaranteed to be in [0,1), so the convex combination should be < max in exact
            // arithmetic; clamp defensively to preserve the half-open interval under floating-point
            // rounding.
            let mut value = (1.0 - r) * min + r * max;
            if value < min {
                value = min;
            }
            if value >= max {
                value = max.next_down();
            }
            values.push(Value::Number(value));
        }
    }

    Value::Array(Array::new(rows_usize, cols_usize, values))
}

inventory::submit! {
    FunctionSpec {
        name: "EXPAND",
        min_args: 2,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Any],
        implementation: expand_fn,
    }
}

fn expand_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match eval_array_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let rows = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let default_cols = i64::try_from(array.cols).unwrap_or(i64::MAX);
    let cols = match eval_optional_i64(ctx, args.get(2), default_cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let pad_with = match eval_optional_pad_with(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    if rows <= 0 || cols <= 0 {
        return Value::Error(ErrorKind::Value);
    }

    let out_rows = match usize::try_from(rows) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Num),
    };
    let out_cols = match usize::try_from(cols) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Num),
    };

    if out_rows < array.rows || out_cols < array.cols {
        return Value::Error(ErrorKind::Value);
    }

    let total = match out_rows.checked_mul(out_cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Num),
    };
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }

    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for row in 0..out_rows {
        for col in 0..out_cols {
            if row < array.rows && col < array.cols {
                values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
            } else {
                values.push(pad_with.clone());
            }
        }
    }

    Value::Array(Array::new(out_rows, out_cols, values))
}

inventory::submit! {
    FunctionSpec {
        name: "MAP",
        min_args: 2,
        max_args: 255,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: map_fn,
    }
}

fn map_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.len() < 2 {
        return Value::Error(ErrorKind::Value);
    }

    let lambda_expr = args.last().expect("checked args is non-empty");
    let call_name = lambda_call_name(lambda_expr, "__MAP_LAMBDA_CALL");
    let lambda = match eval_lambda_arg(ctx, lambda_expr) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let mut arrays = Vec::with_capacity(args.len().saturating_sub(1));
    for expr in &args[..args.len() - 1] {
        match eval_array_arg(ctx, expr) {
            Ok(v) => arrays.push(v),
            Err(e) => return Value::Error(e),
        }
    }

    if lambda.params.len() != arrays.len() {
        return Value::Error(ErrorKind::Value);
    }

    let (out_rows, out_cols) = match broadcast_shape(&arrays) {
        Some(shape) => shape,
        None => return Value::Error(ErrorKind::Value),
    };

    let call = prepare_lambda_call(&call_name, arrays.len());

    let total = match checked_array_cells(out_rows, out_cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut out = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for row in 0..out_rows {
        for col in 0..out_cols {
            let mut arg_values = Vec::with_capacity(arrays.len());
            for arr in &arrays {
                let v = if arr.rows == 1 && arr.cols == 1 {
                    arr.get(0, 0).cloned().unwrap_or(Value::Blank)
                } else {
                    arr.get(row, col).cloned().unwrap_or(Value::Blank)
                };
                arg_values.push(v);
            }

            let value = invoke_lambda(ctx, &lambda, &call, &arg_values);
            out.push(value);
        }
    }

    Value::Array(Array::new(out_rows, out_cols, out))
}

inventory::submit! {
    FunctionSpec {
        name: "BYROW",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: byrow_fn,
    }
}

fn byrow_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match eval_array_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let call_name = lambda_call_name(&args[1], "__BYROW_LAMBDA_CALL");
    let lambda = match eval_lambda_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    if lambda.params.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }

    let call = prepare_lambda_call(&call_name, 1);
    let mut values = match try_vec_with_capacity::<Value>(array.rows) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for row in 0..array.rows {
        let mut row_values = match try_vec_with_capacity::<Value>(array.cols) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        for col in 0..array.cols {
            row_values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
        }
        let arg = Value::Array(Array::new(1, array.cols, row_values));
        values.push(invoke_lambda(ctx, &lambda, &call, &[arg]));
    }

    Value::Array(Array::new(array.rows, 1, values))
}

inventory::submit! {
    FunctionSpec {
        name: "BYCOL",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: bycol_fn,
    }
}

fn bycol_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match eval_array_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let call_name = lambda_call_name(&args[1], "__BYCOL_LAMBDA_CALL");
    let lambda = match eval_lambda_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    if lambda.params.len() != 1 {
        return Value::Error(ErrorKind::Value);
    }

    let call = prepare_lambda_call(&call_name, 1);
    let mut values = match try_vec_with_capacity::<Value>(array.cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for col in 0..array.cols {
        let mut col_values = match try_vec_with_capacity::<Value>(array.rows) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        for row in 0..array.rows {
            col_values.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
        }
        let arg = Value::Array(Array::new(array.rows, 1, col_values));
        values.push(invoke_lambda(ctx, &lambda, &call, &[arg]));
    }

    Value::Array(Array::new(1, array.cols, values))
}

inventory::submit! {
    FunctionSpec {
        name: "MAKEARRAY",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Any],
        implementation: makearray_fn,
    }
}

fn makearray_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rows = match eval_scalar_arg(ctx, &args[0]).coerce_to_i64_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let cols = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let call_name = lambda_call_name(&args[2], "__MAKEARRAY_LAMBDA_CALL");
    let lambda = match eval_lambda_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    if lambda.params.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }

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
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }

    let call = prepare_lambda_call(&call_name, 2);
    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for row in 0..rows_usize {
        for col in 0..cols_usize {
            let arg_values = [
                Value::Number((row as f64) + 1.0),
                Value::Number((col as f64) + 1.0),
            ];
            values.push(invoke_lambda(ctx, &lambda, &call, &arg_values));
        }
    }

    Value::Array(Array::new(rows_usize, cols_usize, values))
}

inventory::submit! {
    FunctionSpec {
        name: "REDUCE",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: reduce_fn,
    }
}

fn reduce_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let (initial, array_expr, lambda_expr) = match args {
        [array, lambda] => (None, array, lambda),
        [initial, array, lambda] => (Some(eval_scalar_arg(ctx, initial)), array, lambda),
        _ => return Value::Error(ErrorKind::Value),
    };

    let array = match eval_array_arg(ctx, array_expr) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let call_name = lambda_call_name(lambda_expr, "__REDUCE_LAMBDA_CALL");
    let lambda = match eval_lambda_arg(ctx, lambda_expr) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    if lambda.params.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }

    let call = prepare_lambda_call(&call_name, 2);
    let mut iter = array.values.iter();
    let mut acc = match initial {
        Some(initial) => initial,
        None => match iter.next() {
            Some(v) => v.clone(),
            None => return Value::Error(ErrorKind::Calc),
        },
    };

    for cell in iter {
        let args = [acc.clone(), cell.clone()];
        acc = invoke_lambda_value(ctx, &lambda, &call, &args);
    }

    acc
}

inventory::submit! {
    FunctionSpec {
        name: "SCAN",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: scan_fn,
    }
}

fn scan_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let (initial, array_expr, lambda_expr) = match args {
        [array, lambda] => (None, array, lambda),
        [initial, array, lambda] => (Some(initial), array, lambda),
        _ => return Value::Error(ErrorKind::Value),
    };

    let array = match eval_array_arg(ctx, array_expr) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let call_name = lambda_call_name(lambda_expr, "__SCAN_LAMBDA_CALL");
    let lambda = match eval_lambda_arg(ctx, lambda_expr) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    if lambda.params.len() != 2 {
        return Value::Error(ErrorKind::Value);
    }

    let call = prepare_lambda_call(&call_name, 2);
    let mut values = match try_vec_with_capacity::<Value>(array.values.len()) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match initial {
        Some(initial_expr) => {
            let initial = eval_scalar_arg(ctx, initial_expr);
            let mut acc = match scalarize_value(initial) {
                Ok(v) => v,
                Err(e) => return Value::Error(e),
            };

            for cell in &array.values {
                let args = [acc.clone(), cell.clone()];
                acc = invoke_lambda(ctx, &lambda, &call, &args);
                values.push(acc.clone());
            }
        }
        None => {
            let Some((first, rest)) = array.values.split_first() else {
                return Value::Error(ErrorKind::Calc);
            };
            let mut acc = match scalarize_value(first.clone()) {
                Ok(v) => v,
                Err(e) => Value::Error(e),
            };
            values.push(acc.clone());

            for cell in rest {
                let args = [acc.clone(), cell.clone()];
                acc = invoke_lambda(ctx, &lambda, &call, &args);
                values.push(acc.clone());
            }
        }
    }

    Value::Array(Array::new(array.rows, array.cols, values))
}

fn eval_array_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<Array, ErrorKind> {
    arg_value_to_array(ctx, ctx.eval_arg(expr))
}

pub(super) fn arg_value_to_array(
    ctx: &dyn FunctionContext,
    arg: ArgValue,
) -> Result<Array, ErrorKind> {
    match arg {
        ArgValue::Scalar(v) => match v {
            Value::Array(arr) => Ok(arr),
            Value::Error(e) => Err(e),
            other => Ok(Array::new(1, 1, vec![other])),
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            ctx.record_reference(&r);
            let rows = (r.end.row - r.start.row + 1) as usize;
            let cols = (r.end.col - r.start.col + 1) as usize;
            let total = checked_array_cells(rows, cols)?;
            let mut values = try_vec_with_capacity::<Value>(total)?;
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

fn eval_optional_pad_with(
    ctx: &dyn FunctionContext,
    expr: Option<&CompiledExpr>,
) -> Result<Value, ErrorKind> {
    let Some(expr) = expr else {
        return Ok(Value::Error(ErrorKind::NA));
    };
    if matches!(expr, Expr::Blank) {
        return Ok(Value::Error(ErrorKind::NA));
    }
    let v = eval_scalar_arg(ctx, expr);
    match v {
        Value::Array(_) | Value::Spill { .. } => Err(ErrorKind::Value),
        other => Ok(other),
    }
}

fn eval_lambda_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<Lambda, ErrorKind> {
    match eval_scalar_arg(ctx, expr) {
        Value::Lambda(lambda) => Ok(lambda),
        Value::Error(e) => Err(e),
        _ => Err(ErrorKind::Value),
    }
}

fn broadcast_shape(arrays: &[Array]) -> Option<(usize, usize)> {
    let mut rows = 1usize;
    let mut cols = 1usize;
    let mut found = false;
    for arr in arrays {
        if arr.rows != 1 || arr.cols != 1 {
            rows = arr.rows;
            cols = arr.cols;
            found = true;
            break;
        }
    }

    if !found {
        return Some((1, 1));
    }

    for arr in arrays {
        if (arr.rows == 1 && arr.cols == 1) || (arr.rows == rows && arr.cols == cols) {
            continue;
        }
        return None;
    }

    Some((rows, cols))
}

#[derive(Debug, Clone)]
struct LambdaCall {
    name: String,
    arg_names: Vec<String>,
    expr: CompiledExpr,
}

fn prepare_lambda_call(call_name: &str, arg_count: usize) -> LambdaCall {
    let mut arg_names = Vec::with_capacity(arg_count);
    let mut args = Vec::with_capacity(arg_count);
    for idx in 0..arg_count {
        let name = format!("__ARG{idx}");
        args.push(Expr::NameRef(NameRef {
            sheet: SheetReference::Current,
            name: name.clone(),
        }));
        arg_names.push(name);
    }

    LambdaCall {
        name: call_name.to_string(),
        arg_names,
        expr: Expr::Call {
            callee: Box::new(Expr::NameRef(NameRef {
                sheet: SheetReference::Current,
                name: call_name.to_string(),
            })),
            args,
        },
    }
}

fn lambda_call_name(expr: &CompiledExpr, fallback: &str) -> String {
    match expr {
        Expr::NameRef(nref) if matches!(nref.sheet, SheetReference::Current) => nref.name.clone(),
        _ => fallback.to_string(),
    }
}

fn invoke_lambda(
    ctx: &dyn FunctionContext,
    lambda: &Lambda,
    call: &LambdaCall,
    args: &[Value],
) -> Value {
    let value = invoke_lambda_value(ctx, lambda, call, args);
    match scalarize_value(value) {
        Ok(v) => v,
        Err(e) => Value::Error(e),
    }
}

fn invoke_lambda_value(
    ctx: &dyn FunctionContext,
    lambda: &Lambda,
    call: &LambdaCall,
    args: &[Value],
) -> Value {
    if args.len() != call.arg_names.len() {
        return Value::Error(ErrorKind::Value);
    }

    let mut bindings: HashMap<String, Value> = HashMap::with_capacity(call.arg_names.len() + 1);
    bindings.insert(call.name.clone(), Value::Lambda(lambda.clone()));
    for (name, value) in call.arg_names.iter().zip(args.iter()) {
        bindings.insert(name.clone(), value.clone());
    }

    ctx.eval_formula_with_bindings(&call.expr, &bindings)
}

fn scalarize_value(value: Value) -> Result<Value, ErrorKind> {
    match value {
        Value::Array(arr) => {
            if arr.rows == 1 && arr.cols == 1 {
                Ok(arr.top_left())
            } else {
                Err(ErrorKind::Value)
            }
        }
        Value::Spill { .. } => Err(ErrorKind::Value),
        other => Ok(other),
    }
}

fn sort_vector_indices(
    ctx: &dyn FunctionContext,
    expr: Option<&CompiledExpr>,
    key_count: usize,
) -> Result<Vec<usize>, ErrorKind> {
    if key_count == 0 {
        return Err(ErrorKind::Value);
    }

    let Some(expr) = expr else {
        return Ok(vec![0]);
    };
    if matches!(expr, Expr::Blank) {
        return Ok(vec![0]);
    }

    let arr = arg_value_to_array(ctx, ctx.eval_arg(expr))?;
    if arr.rows != 1 && arr.cols != 1 {
        return Err(ErrorKind::Value);
    }
    if arr.values.is_empty() {
        return Err(ErrorKind::Value);
    }

    let max_index = i64::try_from(key_count).unwrap_or(i64::MAX);
    let mut indices = try_vec_with_capacity::<usize>(arr.values.len())?;
    for v in &arr.values {
        let idx = v.coerce_to_i64_with_ctx(ctx)?;
        if idx < 1 || idx > max_index {
            return Err(ErrorKind::Value);
        }
        indices.push((idx - 1) as usize);
    }
    Ok(indices)
}

fn sort_vector_orders(
    ctx: &dyn FunctionContext,
    expr: Option<&CompiledExpr>,
    key_len: usize,
) -> Result<Vec<bool>, ErrorKind> {
    let Some(expr) = expr else {
        return Ok(vec![false; key_len]);
    };
    if matches!(expr, Expr::Blank) {
        return Ok(vec![false; key_len]);
    }

    let arr = arg_value_to_array(ctx, ctx.eval_arg(expr))?;
    if arr.rows != 1 && arr.cols != 1 {
        return Err(ErrorKind::Value);
    }
    if arr.values.is_empty() {
        return Err(ErrorKind::Value);
    }

    let mut orders = try_vec_with_capacity::<bool>(arr.values.len())?;
    for v in &arr.values {
        orders.push(sort_descending_from_value(ctx, v)?);
    }

    if orders.len() == 1 {
        Ok(vec![orders[0]; key_len])
    } else if orders.len() == key_len {
        Ok(orders)
    } else {
        Err(ErrorKind::Value)
    }
}

fn eval_optional_i64(
    ctx: &dyn FunctionContext,
    expr: Option<&CompiledExpr>,
    default: i64,
) -> Result<i64, ErrorKind> {
    let Some(expr) = expr else {
        return Ok(default);
    };
    if matches!(expr, Expr::Blank) {
        return Ok(default);
    }
    let v = eval_scalar_arg(ctx, expr);
    match v {
        Value::Error(e) => Err(e),
        other => other.coerce_to_i64_with_ctx(ctx),
    }
}

fn eval_optional_number(
    ctx: &dyn FunctionContext,
    expr: Option<&CompiledExpr>,
    default: f64,
) -> Result<f64, ErrorKind> {
    let Some(expr) = expr else {
        return Ok(default);
    };
    if matches!(expr, Expr::Blank) {
        return Ok(default);
    }
    let v = eval_scalar_arg(ctx, expr);
    match v {
        Value::Error(e) => Err(e),
        other => other.coerce_to_number_with_ctx(ctx),
    }
}

fn eval_optional_bool(
    ctx: &dyn FunctionContext,
    expr: Option<&CompiledExpr>,
    default: bool,
) -> Result<bool, ErrorKind> {
    let Some(expr) = expr else {
        return Ok(default);
    };
    if matches!(expr, Expr::Blank) {
        return Ok(default);
    }
    let v = eval_scalar_arg(ctx, expr);
    match v {
        Value::Error(e) => Err(e),
        other => other.coerce_to_bool_with_ctx(ctx),
    }
}

fn sort_descending_from_value(ctx: &dyn FunctionContext, value: &Value) -> Result<bool, ErrorKind> {
    match value {
        Value::Error(e) => Err(*e),
        Value::Array(_) | Value::Lambda(_) | Value::Spill { .. } => Err(ErrorKind::Value),
        other => match other.coerce_to_i64_with_ctx(ctx) {
            Ok(1) => Ok(false),
            Ok(-1) => Ok(true),
            Ok(_) => Err(ErrorKind::Value),
            Err(e) => Err(e),
        },
    }
}

fn arg_value_is_array_like(value: &ArgValue) -> bool {
    matches!(value, ArgValue::Reference(r) if !r.is_single_cell())
        || matches!(value, ArgValue::ReferenceUnion(_))
        || matches!(value, ArgValue::Scalar(Value::Array(_)))
}

fn take_span(len: usize, n: i64) -> (usize, usize) {
    if len == 0 {
        return (0, 0);
    }
    if n == 0 {
        return (0, 0);
    }

    let len_u64 = len as u64;
    let mag = n.unsigned_abs().min(len_u64);
    let count = usize::try_from(mag).unwrap_or(len);

    if n.is_negative() {
        (len.saturating_sub(count), count)
    } else {
        (0, count)
    }
}

fn drop_span(len: usize, n: i64) -> (usize, usize) {
    if len == 0 {
        return (0, 0);
    }
    if n == 0 {
        return (0, len);
    }

    let len_u64 = len as u64;
    let mag = n.unsigned_abs().min(len_u64);
    let drop = usize::try_from(mag).unwrap_or(len);

    if n.is_negative() {
        (0, len.saturating_sub(drop))
    } else {
        (drop.min(len), len.saturating_sub(drop))
    }
}

fn normalize_index(index: i64, len: usize) -> Option<usize> {
    if len == 0 || index == 0 {
        return None;
    }

    let len_i64 = i64::try_from(len).ok()?;
    let pos = if index > 0 {
        index - 1
    } else {
        len_i64 + index
    };
    if pos < 0 || pos >= len_i64 {
        return None;
    }
    usize::try_from(pos).ok()
}

fn to_vector_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr], to_col: bool) -> Value {
    let array = match eval_array_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let ignore = match eval_optional_i64(ctx, args.get(1), 0) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let ignore_blanks = ignore == 1 || ignore == 3;
    let ignore_errors = ignore == 2 || ignore == 3;
    if ignore < 0 || ignore > 3 {
        return Value::Error(ErrorKind::Value);
    }

    let scan_by_column = match eval_optional_bool(ctx, args.get(2), false) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let total = match checked_array_cells(array.rows, array.cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut values = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let push_cell = |v: Value, out: &mut Vec<Value>| {
        if ignore_blanks && matches!(v, Value::Blank) {
            return;
        }
        if ignore_errors && matches!(v, Value::Error(_)) {
            return;
        }
        out.push(v);
    };

    if scan_by_column {
        for col in 0..array.cols {
            for row in 0..array.rows {
                push_cell(
                    array.get(row, col).cloned().unwrap_or(Value::Blank),
                    &mut values,
                );
            }
        }
    } else {
        for row in 0..array.rows {
            for col in 0..array.cols {
                push_cell(
                    array.get(row, col).cloned().unwrap_or(Value::Blank),
                    &mut values,
                );
            }
        }
    }

    if values.is_empty() {
        return Value::Error(ErrorKind::Calc);
    }

    if to_col {
        Value::Array(Array::new(values.len(), 1, values))
    } else {
        Value::Array(Array::new(1, values.len(), values))
    }
}

fn wrap_vector_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr], wrap_rows: bool) -> Value {
    let array = match eval_array_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let wrap_count = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    if wrap_count <= 0 {
        return Value::Error(ErrorKind::Value);
    }

    let wrap_count = match usize::try_from(wrap_count) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Num),
    };

    let pad_with = match eval_optional_pad_with(ctx, args.get(2)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let total = match checked_array_cells(array.rows, array.cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut flat = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for row in 0..array.rows {
        for col in 0..array.cols {
            flat.push(array.get(row, col).cloned().unwrap_or(Value::Blank));
        }
    }

    if flat.is_empty() {
        return Value::Error(ErrorKind::Calc);
    }

    if wrap_rows {
        let out_cols = wrap_count;
        let out_rows = flat.len().div_ceil(out_cols);

        let total = match checked_array_cells(out_rows, out_cols) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let mut values = match try_vec_with_capacity::<Value>(total) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        for idx in 0..total {
            if let Some(v) = flat.get(idx).cloned() {
                values.push(v);
            } else {
                values.push(pad_with.clone());
            }
        }

        return Value::Array(Array::new(out_rows, out_cols, values));
    }

    let out_rows = wrap_count;
    let out_cols = flat.len().div_ceil(out_rows);

    let total = match checked_array_cells(out_rows, out_cols) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let mut values = match try_vec_with_capacity::<Value>(total) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    for row in 0..out_rows {
        for col in 0..out_cols {
            let idx = col.saturating_mul(out_rows).saturating_add(row);
            if let Some(v) = flat.get(idx).cloned() {
                values.push(v);
            } else {
                values.push(pad_with.clone());
            }
        }
    }

    Value::Array(Array::new(out_rows, out_cols, values))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::date::ExcelDateSystem;
    use crate::functions::{Reference, SheetId};
    use crate::value::{EntityValue, RecordValue};
    use chrono::TimeZone;
    struct DummyContext;

    impl FunctionContext for DummyContext {
        fn eval_arg(&self, _expr: &CompiledExpr) -> ArgValue {
            unreachable!("not needed for sort_key tests")
        }

        fn eval_scalar(&self, _expr: &CompiledExpr) -> Value {
            unreachable!("not needed for sort_key tests")
        }

        fn eval_formula(&self, _expr: &CompiledExpr) -> Value {
            unreachable!("not needed for sort_key tests")
        }

        fn eval_formula_with_bindings(
            &self,
            _expr: &CompiledExpr,
            _bindings: &std::collections::HashMap<String, Value>,
        ) -> Value {
            unreachable!("not needed for sort_key tests")
        }

        fn capture_lexical_env(&self) -> std::collections::HashMap<String, Value> {
            std::collections::HashMap::new()
        }

        fn apply_implicit_intersection(&self, _reference: &Reference) -> Value {
            Value::Blank
        }

        fn get_cell_value(&self, _sheet_id: &SheetId, _addr: CellAddr) -> Value {
            Value::Blank
        }

        fn iter_reference_cells<'a>(
            &'a self,
            _reference: &'a Reference,
        ) -> Box<dyn Iterator<Item = CellAddr> + 'a> {
            Box::new(std::iter::empty())
        }

        fn now_utc(&self) -> chrono::DateTime<chrono::Utc> {
            chrono::Utc.timestamp_opt(0, 0).single().unwrap()
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

    fn stable_sort_order(values: &[Value]) -> Vec<usize> {
        let ctx = DummyContext;
        let keys: Vec<_> = values.iter().map(|v| sort_key(&ctx, v)).collect();
        let mut order: Vec<usize> = (0..values.len()).collect();
        order.sort_by(|&a, &b| {
            let ord = compare_sort_keys(&keys[a], &keys[b], false);
            if ord == Ordering::Equal {
                a.cmp(&b)
            } else {
                ord
            }
        });
        order
    }

    #[test]
    fn compare_sort_keys_ranks_field_error_deterministically() {
        let values = vec![
            Value::Error(ErrorKind::Field),
            Value::Error(ErrorKind::Div0),
            Value::Error(ErrorKind::Calc),
        ];

        let order = stable_sort_order(&values);
        let sorted: Vec<ErrorKind> = order
            .into_iter()
            .map(|idx| match values[idx] {
                Value::Error(e) => e,
                _ => unreachable!("expected error value"),
            })
            .collect();

        assert_eq!(
            sorted,
            vec![ErrorKind::Div0, ErrorKind::Calc, ErrorKind::Field]
        );
    }

    #[test]
    fn compare_sort_keys_sorts_entities_and_records_by_display_case_insensitive_and_stable() {
        let values = vec![
            Value::Record(RecordValue::new("b")),
            Value::Entity(EntityValue::new("A")),
            Value::Record(RecordValue::new("a")),
            // Simulate missing display: empty string sorts like blank.
            Value::Entity(EntityValue::new("")),
        ];

        let order = stable_sort_order(&values);

        // Sorted by display string case-insensitively; ties preserve original order.
        assert_eq!(order, vec![1, 2, 0, 3]);
    }
}
