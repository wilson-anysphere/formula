use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec, Reference};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::functions::lookup;
use crate::value::{ErrorKind, Value};

inventory::submit! {
    FunctionSpec {
        name: "VLOOKUP",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Any],
        implementation: vlookup_fn,
    }
}

fn vlookup_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let lookup_value = eval_scalar_arg(ctx, &args[0]);
    if let Value::Error(e) = lookup_value {
        return Value::Error(e);
    }

    let table_ref = match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(r) => r,
        ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };
    let table = table_ref.normalized();

    let col_index = match eval_scalar_arg(ctx, &args[2]).coerce_to_i64() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    if col_index < 1 {
        return Value::Error(ErrorKind::Value);
    }
    let cols = (table.end.col - table.start.col + 1) as i64;
    if col_index > cols {
        return Value::Error(ErrorKind::Ref);
    }

    let approx = if args.len() == 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_bool() {
            Ok(b) => b,
            Err(e) => return Value::Error(e),
        }
    } else {
        true
    };

    let row_offset = if approx {
        match approximate_match_in_first_col(ctx, &lookup_value, table) {
            Some(r) => r,
            None => return Value::Error(ErrorKind::NA),
        }
    } else {
        match exact_match_in_first_col(ctx, &lookup_value, table) {
            Some(r) => r,
            None => return Value::Error(ErrorKind::NA),
        }
    };

    let result_addr = crate::eval::CellAddr {
        row: table.start.row + row_offset,
        col: table.start.col + (col_index as u32) - 1,
    };
    ctx.get_cell_value(table.sheet_id, result_addr)
}

inventory::submit! {
    FunctionSpec {
        name: "HLOOKUP",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Any],
        implementation: hlookup_fn,
    }
}

fn hlookup_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let lookup_value = eval_scalar_arg(ctx, &args[0]);
    if let Value::Error(e) = lookup_value {
        return Value::Error(e);
    }

    let table_ref = match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(r) => r,
        ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };
    let table = table_ref.normalized();

    let row_index = match eval_scalar_arg(ctx, &args[2]).coerce_to_i64() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    if row_index < 1 {
        return Value::Error(ErrorKind::Value);
    }
    let rows = (table.end.row - table.start.row + 1) as i64;
    if row_index > rows {
        return Value::Error(ErrorKind::Ref);
    }

    let approx = if args.len() == 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_bool() {
            Ok(b) => b,
            Err(e) => return Value::Error(e),
        }
    } else {
        true
    };

    let col_offset = if approx {
        match approximate_match_in_first_row(ctx, &lookup_value, table) {
            Some(c) => c,
            None => return Value::Error(ErrorKind::NA),
        }
    } else {
        match exact_match_in_first_row(ctx, &lookup_value, table) {
            Some(c) => c,
            None => return Value::Error(ErrorKind::NA),
        }
    };

    let result_addr = crate::eval::CellAddr {
        row: table.start.row + (row_index as u32) - 1,
        col: table.start.col + col_offset,
    };
    ctx.get_cell_value(table.sheet_id, result_addr)
}

inventory::submit! {
    FunctionSpec {
        name: "INDEX",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: index_fn,
    }
}

fn index_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let array = match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(r) => r.normalized(),
        ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };
    let row = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let col = if args.len() == 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_i64() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        }
    } else {
        1
    };
    if row < 1 || col < 1 {
        return Value::Error(ErrorKind::Value);
    }
    let rows = (array.end.row - array.start.row + 1) as i64;
    let cols = (array.end.col - array.start.col + 1) as i64;
    if row > rows || col > cols {
        return Value::Error(ErrorKind::Ref);
    }
    let addr = crate::eval::CellAddr {
        row: array.start.row + (row as u32) - 1,
        col: array.start.col + (col as u32) - 1,
    };
    ctx.get_cell_value(array.sheet_id, addr)
}

inventory::submit! {
    FunctionSpec {
        name: "MATCH",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number],
        implementation: match_fn,
    }
}

fn match_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let lookup = eval_scalar_arg(ctx, &args[0]);
    if let Value::Error(e) = lookup {
        return Value::Error(e);
    }

    let array = match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(r) => r.normalized(),
        ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };

    let match_type = if args.len() == 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_i64() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        }
    } else {
        1
    };

    let values = match flatten_1d(array) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::NA),
    };

    let pos = match match_type {
        0 => exact_match_1d(ctx, &lookup, array.sheet_id, &values),
        1 => approximate_match_1d(ctx, &lookup, array.sheet_id, &values, true),
        -1 => approximate_match_1d(ctx, &lookup, array.sheet_id, &values, false),
        _ => return Value::Error(ErrorKind::NA),
    };

    match pos {
        Some(p) => Value::Number((p + 1) as f64),
        None => Value::Error(ErrorKind::NA),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "XMATCH",
        min_args: 2,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: xmatch_fn,
    }
}

fn xmatch_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let lookup_value = eval_scalar_arg(ctx, &args[0]);
    if let Value::Error(e) = lookup_value {
        return Value::Error(e);
    }

    let lookup_range = match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(r) => r.normalized(),
        ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };

    // Only the most common XMATCH mode is implemented for now:
    // - match_mode = 0 (exact)
    // - search_mode = 1 (first-to-last)
    if let Some(expr) = args.get(2) {
        let mode = match eval_scalar_arg(ctx, expr).coerce_to_i64() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        if mode != 0 {
            return Value::Error(ErrorKind::Value);
        }
    }
    if let Some(expr) = args.get(3) {
        let mode = match eval_scalar_arg(ctx, expr).coerce_to_i64() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        if mode != 1 {
            return Value::Error(ErrorKind::Value);
        }
    }

    let values = match flatten_1d_values(ctx, lookup_range) {
        Some(values) => values,
        None => return Value::Error(ErrorKind::Value),
    };

    match lookup::xmatch(&lookup_value, &values) {
        Ok(pos) => Value::Number(pos as f64),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "XLOOKUP",
        min_args: 3,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: xlookup_fn,
    }
}

fn xlookup_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let lookup_value = eval_scalar_arg(ctx, &args[0]);
    if let Value::Error(e) = lookup_value {
        return Value::Error(e);
    }

    let lookup_range = match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(r) => r.normalized(),
        ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };
    let return_range = match ctx.eval_arg(&args[2]) {
        ArgValue::Reference(r) => r.normalized(),
        ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };

    let if_not_found = args.get(3).map(|expr| eval_scalar_arg(ctx, expr));

    // Only the most common XLOOKUP mode is implemented for now:
    // - match_mode = 0 (exact)
    // - search_mode = 1 (first-to-last)
    if let Some(expr) = args.get(4) {
        let mode = match eval_scalar_arg(ctx, expr).coerce_to_i64() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        if mode != 0 {
            return Value::Error(ErrorKind::Value);
        }
    }
    if let Some(expr) = args.get(5) {
        let mode = match eval_scalar_arg(ctx, expr).coerce_to_i64() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        if mode != 1 {
            return Value::Error(ErrorKind::Value);
        }
    }

    let lookup_values = match flatten_1d_values(ctx, lookup_range) {
        Some(values) => values,
        None => return Value::Error(ErrorKind::Value),
    };
    let return_values = match flatten_1d_values(ctx, return_range) {
        Some(values) => values,
        None => return Value::Error(ErrorKind::Value),
    };

    match lookup::xlookup(&lookup_value, &lookup_values, &return_values, if_not_found) {
        Ok(v) => v,
        Err(e) => Value::Error(e),
    }
}

fn flatten_1d(r: Reference) -> Option<Vec<crate::eval::CellAddr>> {
    if r.start.row == r.end.row {
        let cols = r.start.col..=r.end.col;
        Some(cols.map(|col| crate::eval::CellAddr { row: r.start.row, col }).collect())
    } else if r.start.col == r.end.col {
        let rows = r.start.row..=r.end.row;
        Some(rows.map(|row| crate::eval::CellAddr { row, col: r.start.col }).collect())
    } else {
        None
    }
}

fn flatten_1d_values(ctx: &dyn FunctionContext, r: Reference) -> Option<Vec<Value>> {
    let sheet_id = r.sheet_id;
    let addrs = flatten_1d(r)?;
    Some(
        addrs
            .into_iter()
            .map(|addr| ctx.get_cell_value(sheet_id, addr))
            .collect(),
    )
}

fn exact_match_in_first_col(ctx: &dyn FunctionContext, lookup: &Value, table: Reference) -> Option<u32> {
    let rows = table.start.row..=table.end.row;
    for (idx, row) in rows.enumerate() {
        let addr = crate::eval::CellAddr { row, col: table.start.col };
        let key = ctx.get_cell_value(table.sheet_id, addr);
        if excel_eq(lookup, &key) {
            return Some(idx as u32);
        }
    }
    None
}

fn exact_match_in_first_row(ctx: &dyn FunctionContext, lookup: &Value, table: Reference) -> Option<u32> {
    let cols = table.start.col..=table.end.col;
    for (idx, col) in cols.enumerate() {
        let addr = crate::eval::CellAddr { row: table.start.row, col };
        let key = ctx.get_cell_value(table.sheet_id, addr);
        if excel_eq(lookup, &key) {
            return Some(idx as u32);
        }
    }
    None
}

fn approximate_match_in_first_col(ctx: &dyn FunctionContext, lookup: &Value, table: Reference) -> Option<u32> {
    let mut best: Option<u32> = None;
    let rows = table.start.row..=table.end.row;
    for (idx, row) in rows.enumerate() {
        let addr = crate::eval::CellAddr { row, col: table.start.col };
        let key = ctx.get_cell_value(table.sheet_id, addr);
        if excel_le(&key, lookup)? {
            best = Some(idx as u32);
        } else {
            break;
        }
    }
    best
}

fn approximate_match_in_first_row(ctx: &dyn FunctionContext, lookup: &Value, table: Reference) -> Option<u32> {
    let mut best: Option<u32> = None;
    let cols = table.start.col..=table.end.col;
    for (idx, col) in cols.enumerate() {
        let addr = crate::eval::CellAddr { row: table.start.row, col };
        let key = ctx.get_cell_value(table.sheet_id, addr);
        if excel_le(&key, lookup)? {
            best = Some(idx as u32);
        } else {
            break;
        }
    }
    best
}

fn exact_match_1d(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    sheet_id: usize,
    values: &[crate::eval::CellAddr],
) -> Option<usize> {
    for (idx, addr) in values.iter().enumerate() {
        let v = ctx.get_cell_value(sheet_id, *addr);
        if excel_eq(lookup, &v) {
            return Some(idx);
        }
    }
    None
}

fn approximate_match_1d(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    sheet_id: usize,
    values: &[crate::eval::CellAddr],
    ascending: bool,
) -> Option<usize> {
    let mut best: Option<usize> = None;
    for (idx, addr) in values.iter().enumerate() {
        let v = ctx.get_cell_value(sheet_id, *addr);
        let ok = if ascending {
            excel_le(&v, lookup)?
        } else {
            excel_ge(&v, lookup)?
        };
        if ok {
            best = Some(idx);
        } else {
            break;
        }
    }
    best
}

fn excel_eq(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Number(x), Value::Number(y)) => x == y,
        (Value::Text(x), Value::Text(y)) => x.eq_ignore_ascii_case(y),
        (Value::Bool(x), Value::Bool(y)) => x == y,
        (Value::Blank, Value::Blank) => true,
        (Value::Error(x), Value::Error(y)) => x == y,
        _ => false,
    }
}

fn excel_le(a: &Value, b: &Value) -> Option<bool> {
    excel_cmp(a, b).map(|o| o <= 0)
}

fn excel_ge(a: &Value, b: &Value) -> Option<bool> {
    excel_cmp(a, b).map(|o| o >= 0)
}

fn excel_cmp(a: &Value, b: &Value) -> Option<i32> {
    match (a, b) {
        (Value::Number(x), Value::Number(y)) => match x.partial_cmp(y)? {
            std::cmp::Ordering::Less => Some(-1),
            std::cmp::Ordering::Equal => Some(0),
            std::cmp::Ordering::Greater => Some(1),
        },
        (Value::Text(x), Value::Text(y)) => {
            let ord = x.to_ascii_lowercase().cmp(&y.to_ascii_lowercase());
            Some(match ord {
                std::cmp::Ordering::Less => -1,
                std::cmp::Ordering::Equal => 0,
                std::cmp::Ordering::Greater => 1,
            })
        }
        _ => None,
    }
}
