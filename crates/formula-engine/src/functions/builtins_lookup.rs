use crate::eval::CompiledExpr;
use crate::functions::lookup;
use crate::functions::wildcard::WildcardPattern;
use crate::functions::{
    eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec, Reference, SheetId,
};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{Array, ErrorKind, Value};
use std::borrow::Cow;
use std::collections::HashMap;

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

    let col_index = match eval_scalar_arg(ctx, &args[2]).coerce_to_i64() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    if col_index < 1 {
        return Value::Error(ErrorKind::Value);
    }
    let approx = if args.len() == 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_bool() {
            Ok(b) => b,
            Err(e) => return Value::Error(e),
        }
    } else {
        true
    };

    match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(table_ref) => {
            let table = table_ref.normalized();
            // Record dereference for dynamic dependency tracing (e.g. VLOOKUP(…, OFFSET(...), …)).
            ctx.record_reference(&table);
            let cols = (table.end.col - table.start.col + 1) as i64;
            if col_index > cols {
                return Value::Error(ErrorKind::Ref);
            }

            let row_offset = if approx {
                match approximate_match_in_first_col(ctx, &lookup_value, &table) {
                    Some(r) => r,
                    None => return Value::Error(ErrorKind::NA),
                }
            } else {
                match exact_match_in_first_col(ctx, &lookup_value, &table) {
                    Some(r) => r,
                    None => return Value::Error(ErrorKind::NA),
                }
            };

            let result_addr = crate::eval::CellAddr {
                row: table.start.row + row_offset,
                col: table.start.col + (col_index as u32) - 1,
            };
            ctx.get_cell_value(&table.sheet_id, result_addr)
        }
        ArgValue::Scalar(Value::Array(table)) => {
            let cols = i64::try_from(table.cols).unwrap_or(i64::MAX);
            if col_index > cols {
                return Value::Error(ErrorKind::Ref);
            }

            let row_offset = if approx {
                match approximate_match_in_first_col_array(&lookup_value, &table) {
                    Some(r) => r,
                    None => return Value::Error(ErrorKind::NA),
                }
            } else {
                match exact_match_in_first_col_array(&lookup_value, &table) {
                    Some(r) => r,
                    None => return Value::Error(ErrorKind::NA),
                }
            };

            table
                .get(row_offset as usize, (col_index - 1) as usize)
                .cloned()
                .unwrap_or(Value::Blank)
        }
        ArgValue::Scalar(Value::Error(e)) => Value::Error(e),
        ArgValue::ReferenceUnion(_) | ArgValue::Scalar(_) => Value::Error(ErrorKind::Value),
    }
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

    let row_index = match eval_scalar_arg(ctx, &args[2]).coerce_to_i64() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    if row_index < 1 {
        return Value::Error(ErrorKind::Value);
    }
    let approx = if args.len() == 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_bool() {
            Ok(b) => b,
            Err(e) => return Value::Error(e),
        }
    } else {
        true
    };

    match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(table_ref) => {
            let table = table_ref.normalized();
            // Record dereference for dynamic dependency tracing (e.g. HLOOKUP(…, OFFSET(...), …)).
            ctx.record_reference(&table);
            let rows = (table.end.row - table.start.row + 1) as i64;
            if row_index > rows {
                return Value::Error(ErrorKind::Ref);
            }

            let col_offset = if approx {
                match approximate_match_in_first_row(ctx, &lookup_value, &table) {
                    Some(c) => c,
                    None => return Value::Error(ErrorKind::NA),
                }
            } else {
                match exact_match_in_first_row(ctx, &lookup_value, &table) {
                    Some(c) => c,
                    None => return Value::Error(ErrorKind::NA),
                }
            };

            let result_addr = crate::eval::CellAddr {
                row: table.start.row + (row_index as u32) - 1,
                col: table.start.col + col_offset,
            };
            ctx.get_cell_value(&table.sheet_id, result_addr)
        }
        ArgValue::Scalar(Value::Array(table)) => {
            let rows = i64::try_from(table.rows).unwrap_or(i64::MAX);
            if row_index > rows {
                return Value::Error(ErrorKind::Ref);
            }

            let col_offset = if approx {
                match approximate_match_in_first_row_array(&lookup_value, &table) {
                    Some(c) => c,
                    None => return Value::Error(ErrorKind::NA),
                }
            } else {
                match exact_match_in_first_row_array(&lookup_value, &table) {
                    Some(c) => c,
                    None => return Value::Error(ErrorKind::NA),
                }
            };

            table
                .get((row_index - 1) as usize, col_offset as usize)
                .cloned()
                .unwrap_or(Value::Blank)
        }
        ArgValue::Scalar(Value::Error(e)) => Value::Error(e),
        ArgValue::ReferenceUnion(_) | ArgValue::Scalar(_) => Value::Error(ErrorKind::Value),
    }
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
    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(r) => {
            let array = r.normalized();
            // Record dereference for dynamic dependency tracing (e.g. INDEX(OFFSET(...), …)).
            ctx.record_reference(&array);
            let rows = (array.end.row - array.start.row + 1) as i64;
            let cols = (array.end.col - array.start.col + 1) as i64;
            if row > rows || col > cols {
                return Value::Error(ErrorKind::Ref);
            }
            let addr = crate::eval::CellAddr {
                row: array.start.row + (row as u32) - 1,
                col: array.start.col + (col as u32) - 1,
            };
            ctx.get_cell_value(&array.sheet_id, addr)
        }
        ArgValue::Scalar(Value::Array(arr)) => {
            let rows = i64::try_from(arr.rows).unwrap_or(i64::MAX);
            let cols = i64::try_from(arr.cols).unwrap_or(i64::MAX);
            if row > rows || col > cols {
                return Value::Error(ErrorKind::Ref);
            }
            arr.get((row - 1) as usize, (col - 1) as usize)
                .cloned()
                .unwrap_or(Value::Blank)
        }
        ArgValue::Scalar(Value::Error(e)) => Value::Error(e),
        ArgValue::ReferenceUnion(_) | ArgValue::Scalar(_) => Value::Error(ErrorKind::Value),
    }
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

    let match_type = if args.len() == 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_i64() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        }
    } else {
        1
    };

    let pos = match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(r) => {
            let array = r.normalized();
            let values = match flatten_1d(&array) {
                Some(v) => v,
                None => return Value::Error(ErrorKind::NA),
            };
            // Record dereference for dynamic dependency tracing (e.g. MATCH(…, OFFSET(...), …)).
            ctx.record_reference(&array);
            match match_type {
                0 => exact_match_1d(ctx, &lookup, &array.sheet_id, &values),
                1 => approximate_match_1d(ctx, &lookup, &array.sheet_id, &values, true),
                -1 => approximate_match_1d(ctx, &lookup, &array.sheet_id, &values, false),
                _ => return Value::Error(ErrorKind::NA),
            }
        }
        ArgValue::Scalar(Value::Array(arr)) => {
            if arr.rows != 1 && arr.cols != 1 {
                return Value::Error(ErrorKind::NA);
            }
            match match_type {
                0 => exact_match_values(&lookup, &arr.values),
                1 => approximate_match_values(&lookup, &arr.values, true),
                -1 => approximate_match_values(&lookup, &arr.values, false),
                _ => return Value::Error(ErrorKind::NA),
            }
        }
        ArgValue::Scalar(Value::Error(e)) => return Value::Error(e),
        ArgValue::ReferenceUnion(_) | ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
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
    let match_mode = match args.get(2) {
        Some(expr) if matches!(expr, CompiledExpr::Blank) => lookup::MatchMode::Exact,
        Some(expr) => match eval_scalar_arg(ctx, expr).coerce_to_i64() {
            Ok(n) => match lookup::MatchMode::try_from(n) {
                Ok(m) => m,
                Err(e) => return Value::Error(e),
            },
            Err(e) => return Value::Error(e),
        },
        None => lookup::MatchMode::Exact,
    };
    let search_mode = match args.get(3) {
        Some(expr) if matches!(expr, CompiledExpr::Blank) => lookup::SearchMode::FirstToLast,
        Some(expr) => match eval_scalar_arg(ctx, expr).coerce_to_i64() {
            Ok(n) => match lookup::SearchMode::try_from(n) {
                Ok(m) => m,
                Err(e) => return Value::Error(e),
            },
            Err(e) => return Value::Error(e),
        },
        None => lookup::SearchMode::FirstToLast,
    };

    match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(r) => {
            let r = r.normalized();
            let (shape, len) = match reference_1d_shape_len(&r) {
                Some(v) => v,
                None => return Value::Error(ErrorKind::Value),
            };
            // Record that the lookup_array reference was dereferenced so dynamic reference
            // arguments (OFFSET/INDIRECT) participate in dependency tracing.
            ctx.record_reference(&r);
            let sheet_id = &r.sheet_id;
            let start = r.start;

            let pos = lookup::xmatch_with_modes_accessor(
                &lookup_value,
                len,
                |idx| match shape {
                    XlookupVectorShape::Vertical => ctx.get_cell_value(
                        sheet_id,
                        crate::eval::CellAddr {
                            row: start.row + idx as u32,
                            col: start.col,
                        },
                    ),
                    XlookupVectorShape::Horizontal => ctx.get_cell_value(
                        sheet_id,
                        crate::eval::CellAddr {
                            row: start.row,
                            col: start.col + idx as u32,
                        },
                    ),
                },
                match_mode,
                search_mode,
            );

            match pos {
                Ok(p) => Value::Number(p as f64),
                Err(e) => Value::Error(e),
            }
        }
        ArgValue::Scalar(Value::Array(arr)) => {
            if arr.rows != 1 && arr.cols != 1 {
                return Value::Error(ErrorKind::Value);
            }
            match lookup::xmatch_with_modes(&lookup_value, &arr.values, match_mode, search_mode) {
                Ok(pos) => Value::Number(pos as f64),
                Err(e) => Value::Error(e),
            }
        }
        ArgValue::Scalar(Value::Error(e)) => Value::Error(e),
        ArgValue::ReferenceUnion(_) | ArgValue::Scalar(_) => Value::Error(ErrorKind::Value),
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
        // XLOOKUP's optional match/search modes are numeric, so expose them in
        // the lightweight metadata consumed by editor tooling.
        arg_types: &[
            ValueType::Any,
            ValueType::Any,
            ValueType::Any,
            ValueType::Any,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: xlookup_fn,
    }
}

fn xlookup_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let lookup_value = eval_scalar_arg(ctx, &args[0]);
    if let Value::Error(e) = lookup_value {
        return Value::Error(e);
    }

    let if_not_found = match args.get(3) {
        Some(expr) if matches!(expr, CompiledExpr::Blank) => None,
        Some(expr) => Some(eval_scalar_arg(ctx, expr)),
        None => None,
    };

    let match_mode = match args.get(4) {
        Some(expr) if matches!(expr, CompiledExpr::Blank) => lookup::MatchMode::Exact,
        Some(expr) => match eval_scalar_arg(ctx, expr).coerce_to_i64() {
            Ok(n) => match lookup::MatchMode::try_from(n) {
                Ok(m) => m,
                Err(e) => return Value::Error(e),
            },
            Err(e) => return Value::Error(e),
        },
        None => lookup::MatchMode::Exact,
    };
    let search_mode = match args.get(5) {
        Some(expr) if matches!(expr, CompiledExpr::Blank) => lookup::SearchMode::FirstToLast,
        Some(expr) => match eval_scalar_arg(ctx, expr).coerce_to_i64() {
            Ok(n) => match lookup::SearchMode::try_from(n) {
                Ok(m) => m,
                Err(e) => return Value::Error(e),
            },
            Err(e) => return Value::Error(e),
        },
        None => lookup::SearchMode::FirstToLast,
    };

    enum XlookupLookupArray {
        Values {
            shape: XlookupVectorShape,
            values: Vec<Value>,
        },
        Reference {
            shape: XlookupVectorShape,
            reference: Reference,
            len: usize,
        },
    }

    impl XlookupLookupArray {
        fn shape(&self) -> XlookupVectorShape {
            match self {
                XlookupLookupArray::Values { shape, .. }
                | XlookupLookupArray::Reference { shape, .. } => *shape,
            }
        }

        fn len(&self) -> usize {
            match self {
                XlookupLookupArray::Values { values, .. } => values.len(),
                XlookupLookupArray::Reference { len, .. } => *len,
            }
        }

        fn xmatch(
            &self,
            ctx: &dyn FunctionContext,
            lookup_value: &Value,
            match_mode: lookup::MatchMode,
            search_mode: lookup::SearchMode,
        ) -> Result<i32, ErrorKind> {
            match self {
                XlookupLookupArray::Values { values, .. } => {
                    lookup::xmatch_with_modes(lookup_value, values, match_mode, search_mode)
                }
                XlookupLookupArray::Reference {
                    shape,
                    reference,
                    len,
                } => {
                    let sheet_id = &reference.sheet_id;
                    let start = reference.start;
                    lookup::xmatch_with_modes_accessor(
                        lookup_value,
                        *len,
                        |idx| match shape {
                            XlookupVectorShape::Vertical => ctx.get_cell_value(
                                sheet_id,
                                crate::eval::CellAddr {
                                    row: start.row + idx as u32,
                                    col: start.col,
                                },
                            ),
                            XlookupVectorShape::Horizontal => ctx.get_cell_value(
                                sheet_id,
                                crate::eval::CellAddr {
                                    row: start.row,
                                    col: start.col + idx as u32,
                                },
                            ),
                        },
                        match_mode,
                        search_mode,
                    )
                }
            }
        }
    }

    let lookup_array = match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(r) => {
            let r = r.normalized();
            let (shape, len) = match reference_1d_shape_len(&r) {
                Some(v) => v,
                None => return Value::Error(ErrorKind::Value),
            };
            // Record that the lookup_array reference was dereferenced so dynamic reference
            // arguments (OFFSET/INDIRECT) participate in dependency tracing.
            ctx.record_reference(&r);
            XlookupLookupArray::Reference {
                shape,
                reference: r,
                len,
            }
        }
        ArgValue::Scalar(Value::Array(arr)) => match array_1d_with_shape(arr) {
            Ok((shape, values)) => XlookupLookupArray::Values { shape, values },
            Err(e) => return Value::Error(e),
        },
        ArgValue::Scalar(Value::Error(e)) => return Value::Error(e),
        ArgValue::ReferenceUnion(_) | ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };

    enum XlookupReturnArray {
        Array(Array),
        Reference(Reference),
    }

    impl XlookupReturnArray {
        fn rows(&self) -> usize {
            match self {
                XlookupReturnArray::Array(arr) => arr.rows,
                XlookupReturnArray::Reference(r) => (r.end.row - r.start.row + 1) as usize,
            }
        }

        fn cols(&self) -> usize {
            match self {
                XlookupReturnArray::Array(arr) => arr.cols,
                XlookupReturnArray::Reference(r) => (r.end.col - r.start.col + 1) as usize,
            }
        }

        fn get(&self, ctx: &dyn FunctionContext, row: usize, col: usize) -> Value {
            match self {
                XlookupReturnArray::Array(arr) => arr.get(row, col).cloned().unwrap_or(Value::Blank),
                XlookupReturnArray::Reference(r) => ctx.get_cell_value(
                    &r.sheet_id,
                    crate::eval::CellAddr {
                        row: r.start.row + row as u32,
                        col: r.start.col + col as u32,
                    },
                ),
            }
        }
    }

    let return_array = match ctx.eval_arg(&args[2]) {
        ArgValue::Reference(r) => XlookupReturnArray::Reference(r.normalized()),
        ArgValue::Scalar(Value::Array(arr)) => XlookupReturnArray::Array(arr),
        ArgValue::Scalar(Value::Error(e)) => return Value::Error(e),
        ArgValue::ReferenceUnion(_) | ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };

    let lookup_shape = lookup_array.shape();
    let lookup_len = lookup_array.len();
    if lookup_len == 0 {
        return match if_not_found {
            Some(v) => v,
            None => Value::Error(ErrorKind::NA),
        };
    }

    // Validate return_array shape:
    // - vertical lookup_array (Nx1) requires return_array.rows == N; return spills horizontally.
    // - horizontal lookup_array (1xN) requires return_array.cols == N; return spills vertically.
    match lookup_shape {
        XlookupVectorShape::Vertical => {
            if return_array.rows() != lookup_len {
                return Value::Error(ErrorKind::Value);
            }
        }
        XlookupVectorShape::Horizontal => {
            if return_array.cols() != lookup_len {
                return Value::Error(ErrorKind::Value);
            }
        }
    }

    let match_pos = match lookup_array.xmatch(ctx, &lookup_value, match_mode, search_mode) {
        Ok(pos) => pos,
        Err(ErrorKind::NA) => {
            return match if_not_found {
                Some(v) => v,
                None => Value::Error(ErrorKind::NA),
            };
        }
        Err(e) => return Value::Error(e),
    };

    let idx = match usize::try_from(match_pos - 1) {
        Ok(v) => v,
        Err(_) => return Value::Error(ErrorKind::Value),
    };

    if let XlookupReturnArray::Reference(r) = &return_array {
        // Record dereference for dynamic dependency tracing (e.g. XLOOKUP with OFFSET return_array).
        ctx.record_reference(r);
    }

    match lookup_shape {
        XlookupVectorShape::Vertical => {
            // Return the matched row.
            let cols = return_array.cols();
            if cols == 1 {
                return return_array.get(ctx, idx, 0);
            }
            let mut values = Vec::with_capacity(cols);
            for col in 0..cols {
                values.push(return_array.get(ctx, idx, col));
            }
            Value::Array(Array::new(1, cols, values))
        }
        XlookupVectorShape::Horizontal => {
            // Return the matched column.
            let rows = return_array.rows();
            if rows == 1 {
                return return_array.get(ctx, 0, idx);
            }
            let mut values = Vec::with_capacity(rows);
            for row in 0..rows {
                values.push(return_array.get(ctx, row, idx));
            }
            Value::Array(Array::new(rows, 1, values))
        }
    }
}

inventory::submit! {
    FunctionSpec {
        name: "GETPIVOTDATA",
        min_args: 2,
        max_args: 255,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        // Signature: GETPIVOTDATA(data_field, pivot_table, [field1, item1], ...)
        arg_types: &[ValueType::Any],
        implementation: getpivotdata_fn,
    }
}

struct PivotLayout {
    header_row: u32,
    top_left_col: u32,
    row_fields: HashMap<String, u32>,
    data_col: u32,
}

fn getpivotdata_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let data_field_value = eval_scalar_arg(ctx, &args[0]);
    if let Value::Error(e) = data_field_value {
        return Value::Error(e);
    }
    let data_field = match data_field_value.coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    if data_field.is_empty() {
        return Value::Error(ErrorKind::Value);
    }

    let pivot_ref = match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(r) => r.normalized(),
        ArgValue::ReferenceUnion(_) | ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };

    // Remaining args must be (field, item) pairs.
    if (args.len() - 2) % 2 != 0 {
        return Value::Error(ErrorKind::Value);
    }

    let layout = match find_pivot_layout(ctx, &pivot_ref, &data_field) {
        Ok(l) => l,
        Err(e) => return Value::Error(e),
    };

    let mut criteria = Vec::new();
    for pair_idx in (2..args.len()).step_by(2) {
        let field_value = eval_scalar_arg(ctx, &args[pair_idx]);
        if let Value::Error(e) = field_value {
            return Value::Error(e);
        }
        let item_value = eval_scalar_arg(ctx, &args[pair_idx + 1]);
        if let Value::Error(e) = item_value {
            return Value::Error(e);
        }
        let field = match field_value.coerce_to_string() {
            Ok(s) => s,
            Err(e) => return Value::Error(e),
        };
        if field.is_empty() {
            return Value::Error(ErrorKind::Value);
        }
        let col = match layout.row_fields.get(&field.to_ascii_uppercase()) {
            Some(c) => *c,
            None => return Value::Error(ErrorKind::Ref),
        };
        criteria.push((col, item_value));
    }

    if criteria.is_empty() {
        return getpivotdata_grand_total(ctx, &pivot_ref.sheet_id, &layout);
    }

    match getpivotdata_find_row(ctx, &pivot_ref.sheet_id, &layout, &criteria) {
        Ok(row) => {
            let addr = crate::eval::CellAddr {
                row,
                col: layout.data_col,
            };
            ctx.get_cell_value(&pivot_ref.sheet_id, addr)
        }
        Err(e) => Value::Error(e),
    }
}

fn find_pivot_layout(
    ctx: &dyn FunctionContext,
    pivot_ref: &Reference,
    data_field: &str,
) -> Result<PivotLayout, ErrorKind> {
    // Heuristics (MVP, pivot-engine-compatible):
    //
    // 1) Starting from the provided pivot_table cell, scan upward for a header cell that
    //    matches `data_field` (case-insensitive). This is assumed to be the value-field
    //    caption written by our pivot engine.
    // 2) From that header cell, scan left across the same row while cells are non-empty
    //    text to locate the pivot table's top-left corner.
    // 3) The row-field columns are the leading columns where the first data row contains
    //    text (row labels). The first non-text cell below the header indicates the start
    //    of the value area.
    //
    // Limitations:
    // - Only supports pivot-engine Tabular layout (not Compact/"Row Labels").
    // - Requires exactly one base value column (plus an optional "Grand Total - ..." column).
    // - Does not support column fields.
    const MAX_SCAN_ROWS: u32 = 10_000;
    const MAX_SCAN_COLS: u32 = 64;

    let sheet_id = &pivot_ref.sheet_id;
    let anchor = pivot_ref.start;

    let mut header_row: Option<u32> = None;
    let mut data_col: Option<u32> = None;

    let max_up = anchor.row.min(MAX_SCAN_ROWS);
    let col_start = anchor.col.saturating_sub(MAX_SCAN_COLS);
    let col_end = anchor.col.saturating_add(MAX_SCAN_COLS);

    for delta in 0..=max_up {
        let row = anchor.row - delta;
        for col in col_start..=col_end {
            let v = ctx.get_cell_value(sheet_id, crate::eval::CellAddr { row, col });
            if value_text_eq(&v, data_field) {
                let below = ctx.get_cell_value(
                    sheet_id,
                    crate::eval::CellAddr {
                        row: row.saturating_add(1),
                        col,
                    },
                );
                // Pivot-engine value cells are numbers/blanks; if we see a text value directly
                // below, this is unlikely to be the pivot header row.
                if !matches!(below, Value::Text(_)) {
                    header_row = Some(row);
                    data_col = Some(col);
                    break;
                }
            }
        }
        if header_row.is_some() {
            break;
        }
    }

    let header_row = header_row.ok_or(ErrorKind::Ref)?;
    let data_col = data_col.ok_or(ErrorKind::Ref)?;

    // Find top-left header cell by scanning left across the header row.
    let mut top_left_col = data_col;
    while top_left_col > 0 {
        let left = ctx.get_cell_value(
            sheet_id,
            crate::eval::CellAddr {
                row: header_row,
                col: top_left_col - 1,
            },
        );
        match left {
            Value::Text(ref s) if !s.is_empty() => {
                top_left_col -= 1;
            }
            _ => break,
        }
    }

    // Disallow Compact layout pivots produced by our engine.
    let first_header = ctx.get_cell_value(
        sheet_id,
        crate::eval::CellAddr {
            row: header_row,
            col: top_left_col,
        },
    );
    if matches!(first_header, Value::Text(ref s) if s.eq_ignore_ascii_case("Row Labels")) {
        return Err(ErrorKind::Ref);
    }

    // Determine where the value area begins by inspecting the first data row.
    let first_data_row = header_row.saturating_add(1);
    let mut first_value_col: Option<u32> = None;
    for col in top_left_col..=data_col {
        let below = ctx.get_cell_value(
            sheet_id,
            crate::eval::CellAddr {
                row: first_data_row,
                col,
            },
        );
        if !matches!(below, Value::Text(_)) {
            first_value_col = Some(col);
            break;
        }
    }
    let first_value_col = first_value_col.ok_or(ErrorKind::Ref)?;
    if first_value_col == top_left_col {
        // Pivot with no row fields isn't supported in the MVP.
        return Err(ErrorKind::Ref);
    }

    // Ensure the requested data_field refers to the base value column.
    if data_col != first_value_col {
        return Err(ErrorKind::Ref);
    }

    // Collect value headers (base + optional grand total).
    let mut value_headers: Vec<String> = Vec::new();
    for col in first_value_col..=first_value_col.saturating_add(MAX_SCAN_COLS) {
        let v = ctx.get_cell_value(
            sheet_id,
            crate::eval::CellAddr {
                row: header_row,
                col,
            },
        );
        match v {
            Value::Text(s) if !s.is_empty() => value_headers.push(s),
            _ => break,
        }
    }
    if value_headers.is_empty() {
        return Err(ErrorKind::Ref);
    }

    let base_headers: Vec<&str> = value_headers
        .iter()
        .filter_map(|h| {
            (!h.to_ascii_lowercase().starts_with("grand total - ")).then_some(h.as_str())
        })
        .collect();

    if base_headers.len() != 1 {
        // Multiple value fields and/or column fields are not supported in the MVP.
        return Err(ErrorKind::Ref);
    }

    let base_header = base_headers[0];
    if !base_header.eq_ignore_ascii_case(data_field) {
        return Err(ErrorKind::Ref);
    }

    // Allow at most a single pivot-engine grand-total column.
    let expected_gt = format!("Grand Total - {base_header}");
    for h in &value_headers {
        if h.eq_ignore_ascii_case(base_header) {
            continue;
        }
        if h.eq_ignore_ascii_case(&expected_gt) {
            continue;
        }
        return Err(ErrorKind::Ref);
    }

    // Row field headers are everything between top-left and the first value column.
    let mut row_fields = HashMap::new();
    for col in top_left_col..first_value_col {
        let v = ctx.get_cell_value(
            sheet_id,
            crate::eval::CellAddr {
                row: header_row,
                col,
            },
        );
        let name = match v {
            Value::Text(s) if !s.is_empty() => s,
            _ => return Err(ErrorKind::Ref),
        };
        row_fields.insert(name.to_ascii_uppercase(), col);
    }

    Ok(PivotLayout {
        header_row,
        top_left_col,
        row_fields,
        data_col,
    })
}

fn getpivotdata_grand_total(
    ctx: &dyn FunctionContext,
    sheet_id: &SheetId,
    layout: &PivotLayout,
) -> Value {
    const MAX_SCAN_ROWS: u32 = 10_000;

    for delta in 1..=MAX_SCAN_ROWS {
        let row = match layout.header_row.checked_add(delta) {
            Some(r) => r,
            None => break,
        };
        let label = ctx.get_cell_value(
            sheet_id,
            crate::eval::CellAddr {
                row,
                col: layout.top_left_col,
            },
        );
        if matches!(label, Value::Text(ref s) if s.eq_ignore_ascii_case("Grand Total")) {
            return ctx.get_cell_value(
                sheet_id,
                crate::eval::CellAddr {
                    row,
                    col: layout.data_col,
                },
            );
        }

        // Stop scanning once we hit a fully blank row in the pivot area.
        let first = ctx.get_cell_value(
            sheet_id,
            crate::eval::CellAddr {
                row,
                col: layout.top_left_col,
            },
        );
        let value = ctx.get_cell_value(
            sheet_id,
            crate::eval::CellAddr {
                row,
                col: layout.data_col,
            },
        );
        if matches!(first, Value::Blank) && matches!(value, Value::Blank) {
            break;
        }
    }

    Value::Error(ErrorKind::Ref)
}

fn getpivotdata_find_row(
    ctx: &dyn FunctionContext,
    sheet_id: &SheetId,
    layout: &PivotLayout,
    criteria: &[(u32, Value)],
) -> Result<u32, ErrorKind> {
    const MAX_SCAN_ROWS: u32 = 10_000;

    let mut found: Option<u32> = None;

    for delta in 1..=MAX_SCAN_ROWS {
        let row = layout.header_row + delta;

        // Stop scanning once we hit a fully blank row in the pivot area.
        let first = ctx.get_cell_value(
            sheet_id,
            crate::eval::CellAddr {
                row,
                col: layout.top_left_col,
            },
        );
        let value = ctx.get_cell_value(
            sheet_id,
            crate::eval::CellAddr {
                row,
                col: layout.data_col,
            },
        );
        if matches!(first, Value::Blank) && matches!(value, Value::Blank) {
            break;
        }

        let mut row_matches = true;
        for (col, item) in criteria {
            let cell = ctx.get_cell_value(sheet_id, crate::eval::CellAddr { row, col: *col });
            if !pivot_item_matches(&cell, item)? {
                row_matches = false;
                break;
            }
        }

        if row_matches {
            if found.is_some() {
                // Ambiguous match (insufficient criteria).
                return Err(ErrorKind::Ref);
            }
            found = Some(row);
        }
    }

    found.ok_or(ErrorKind::NA)
}

fn value_text_eq(v: &Value, s: &str) -> bool {
    match v {
        Value::Text(t) => t.eq_ignore_ascii_case(s),
        _ => false,
    }
}

fn pivot_item_matches(cell: &Value, item: &Value) -> Result<bool, ErrorKind> {
    match (cell, item) {
        (Value::Error(e), _) => Err(*e),
        (_, Value::Error(e)) => Err(*e),
        (Value::Text(cell_text), _) => {
            let item_text = item.coerce_to_string()?;
            Ok(cell_text.eq_ignore_ascii_case(&item_text))
        }
        (Value::Blank, _) => {
            let item_text = item.coerce_to_string()?;
            Ok(item_text.is_empty())
        }
        _ => Ok(excel_eq(cell, item)),
    }
}

fn flatten_1d(r: &Reference) -> Option<Vec<crate::eval::CellAddr>> {
    if r.start.row == r.end.row {
        let cols = r.start.col..=r.end.col;
        Some(
            cols.map(|col| crate::eval::CellAddr {
                row: r.start.row,
                col,
            })
            .collect(),
        )
    } else if r.start.col == r.end.col {
        let rows = r.start.row..=r.end.row;
        Some(
            rows.map(|row| crate::eval::CellAddr {
                row,
                col: r.start.col,
            })
            .collect(),
        )
    } else {
        None
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum XlookupVectorShape {
    Horizontal,
    Vertical,
}

fn reference_1d_shape_len(r: &Reference) -> Option<(XlookupVectorShape, usize)> {
    if r.start.row == r.end.row && r.start.col == r.end.col {
        return Some((XlookupVectorShape::Vertical, 1));
    }
    if r.start.row == r.end.row {
        let len = (r.end.col - r.start.col + 1) as usize;
        return Some((XlookupVectorShape::Horizontal, len));
    }
    if r.start.col == r.end.col {
        let len = (r.end.row - r.start.row + 1) as usize;
        return Some((XlookupVectorShape::Vertical, len));
    }
    None
}
fn array_1d_with_shape(arr: Array) -> Result<(XlookupVectorShape, Vec<Value>), ErrorKind> {
    if arr.rows == 1 && arr.cols == 1 {
        return Ok((XlookupVectorShape::Vertical, arr.values));
    }
    if arr.rows == 1 {
        return Ok((XlookupVectorShape::Horizontal, arr.values));
    }
    if arr.cols == 1 {
        return Ok((XlookupVectorShape::Vertical, arr.values));
    }
    Err(ErrorKind::Value)
}

fn wildcard_pattern_for_lookup(lookup: &Value) -> Option<WildcardPattern> {
    let Value::Text(pattern) = lookup else {
        return None;
    };
    if !pattern.contains('*') && !pattern.contains('?') && !pattern.contains('~') {
        return None;
    }
    Some(WildcardPattern::new(pattern))
}

fn exact_match_values(lookup: &Value, values: &[Value]) -> Option<usize> {
    if let Some(pattern) = wildcard_pattern_for_lookup(lookup) {
        for (idx, candidate) in values.iter().enumerate() {
            let text = match candidate {
                Value::Error(_) => continue,
                Value::Text(s) => Cow::Borrowed(s.as_str()),
                other => match other.coerce_to_string() {
                    Ok(s) => Cow::Owned(s),
                    Err(_) => continue,
                },
            };
            if pattern.matches(text.as_ref()) {
                return Some(idx);
            }
        }
        return None;
    }

    values.iter().position(|v| excel_eq(lookup, v))
}

fn approximate_match_values(lookup: &Value, values: &[Value], ascending: bool) -> Option<usize> {
    let mut best: Option<usize> = None;
    for (idx, v) in values.iter().enumerate() {
        let ok = if ascending {
            excel_le(v, lookup)?
        } else {
            excel_ge(v, lookup)?
        };
        if ok {
            best = Some(idx);
        } else {
            break;
        }
    }
    best
}

fn exact_match_in_first_col(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    table: &Reference,
) -> Option<u32> {
    let wildcard_pattern = wildcard_pattern_for_lookup(lookup);
    let rows = table.start.row..=table.end.row;
    for (idx, row) in rows.enumerate() {
        let addr = crate::eval::CellAddr {
            row,
            col: table.start.col,
        };
        let key = ctx.get_cell_value(&table.sheet_id, addr);
        if let Some(pattern) = &wildcard_pattern {
            let text = match &key {
                Value::Error(_) => continue,
                Value::Text(s) => Cow::Borrowed(s.as_str()),
                other => match other.coerce_to_string() {
                    Ok(s) => Cow::Owned(s),
                    Err(_) => continue,
                },
            };
            if pattern.matches(text.as_ref()) {
                return Some(idx as u32);
            }
        } else if excel_eq(lookup, &key) {
            return Some(idx as u32);
        }
    }
    None
}

fn exact_match_in_first_col_array(lookup: &Value, table: &Array) -> Option<u32> {
    let wildcard_pattern = wildcard_pattern_for_lookup(lookup);
    for row in 0..table.rows {
        let key = table.get(row, 0).unwrap_or(&Value::Blank);
        if let Some(pattern) = &wildcard_pattern {
            let text = match key {
                Value::Error(_) => continue,
                Value::Text(s) => Cow::Borrowed(s.as_str()),
                other => match other.coerce_to_string() {
                    Ok(s) => Cow::Owned(s),
                    Err(_) => continue,
                },
            };
            if pattern.matches(text.as_ref()) {
                return Some(row as u32);
            }
        } else if excel_eq(lookup, key) {
            return Some(row as u32);
        }
    }
    None
}

fn exact_match_in_first_row(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    table: &Reference,
) -> Option<u32> {
    let wildcard_pattern = wildcard_pattern_for_lookup(lookup);
    let cols = table.start.col..=table.end.col;
    for (idx, col) in cols.enumerate() {
        let addr = crate::eval::CellAddr {
            row: table.start.row,
            col,
        };
        let key = ctx.get_cell_value(&table.sheet_id, addr);
        if let Some(pattern) = &wildcard_pattern {
            let text = match &key {
                Value::Error(_) => continue,
                Value::Text(s) => Cow::Borrowed(s.as_str()),
                other => match other.coerce_to_string() {
                    Ok(s) => Cow::Owned(s),
                    Err(_) => continue,
                },
            };
            if pattern.matches(text.as_ref()) {
                return Some(idx as u32);
            }
        } else if excel_eq(lookup, &key) {
            return Some(idx as u32);
        }
    }
    None
}

fn exact_match_in_first_row_array(lookup: &Value, table: &Array) -> Option<u32> {
    let wildcard_pattern = wildcard_pattern_for_lookup(lookup);
    for col in 0..table.cols {
        let key = table.get(0, col).unwrap_or(&Value::Blank);
        if let Some(pattern) = &wildcard_pattern {
            let text = match key {
                Value::Error(_) => continue,
                Value::Text(s) => Cow::Borrowed(s.as_str()),
                other => match other.coerce_to_string() {
                    Ok(s) => Cow::Owned(s),
                    Err(_) => continue,
                },
            };
            if pattern.matches(text.as_ref()) {
                return Some(col as u32);
            }
        } else if excel_eq(lookup, key) {
            return Some(col as u32);
        }
    }
    None
}

fn approximate_match_in_first_col(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    table: &Reference,
) -> Option<u32> {
    let mut best: Option<u32> = None;
    let rows = table.start.row..=table.end.row;
    for (idx, row) in rows.enumerate() {
        let addr = crate::eval::CellAddr {
            row,
            col: table.start.col,
        };
        let key = ctx.get_cell_value(&table.sheet_id, addr);
        if excel_le(&key, lookup)? {
            best = Some(idx as u32);
        } else {
            break;
        }
    }
    best
}

fn approximate_match_in_first_col_array(lookup: &Value, table: &Array) -> Option<u32> {
    let mut best: Option<u32> = None;
    for row in 0..table.rows {
        let key = table.get(row, 0).unwrap_or(&Value::Blank);
        if excel_le(key, lookup)? {
            best = Some(row as u32);
        } else {
            break;
        }
    }
    best
}

fn approximate_match_in_first_row(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    table: &Reference,
) -> Option<u32> {
    let mut best: Option<u32> = None;
    let cols = table.start.col..=table.end.col;
    for (idx, col) in cols.enumerate() {
        let addr = crate::eval::CellAddr {
            row: table.start.row,
            col,
        };
        let key = ctx.get_cell_value(&table.sheet_id, addr);
        if excel_le(&key, lookup)? {
            best = Some(idx as u32);
        } else {
            break;
        }
    }
    best
}

fn approximate_match_in_first_row_array(lookup: &Value, table: &Array) -> Option<u32> {
    let mut best: Option<u32> = None;
    for col in 0..table.cols {
        let key = table.get(0, col).unwrap_or(&Value::Blank);
        if excel_le(key, lookup)? {
            best = Some(col as u32);
        } else {
            break;
        }
    }
    best
}

fn exact_match_1d(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    sheet_id: &SheetId,
    values: &[crate::eval::CellAddr],
) -> Option<usize> {
    let wildcard_pattern = wildcard_pattern_for_lookup(lookup);
    for (idx, addr) in values.iter().enumerate() {
        let v = ctx.get_cell_value(sheet_id, *addr);
        if let Some(pattern) = &wildcard_pattern {
            let text = match &v {
                Value::Error(_) => continue,
                Value::Text(s) => Cow::Borrowed(s.as_str()),
                other => match other.coerce_to_string() {
                    Ok(s) => Cow::Owned(s),
                    Err(_) => continue,
                },
            };
            if pattern.matches(text.as_ref()) {
                return Some(idx);
            }
        } else if excel_eq(lookup, &v) {
            return Some(idx);
        }
    }
    None
}

fn approximate_match_1d(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    sheet_id: &SheetId,
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

fn text_eq_case_insensitive(a: &str, b: &str) -> bool {
    if a.is_ascii() && b.is_ascii() {
        return a.eq_ignore_ascii_case(b);
    }

    a.chars()
        .flat_map(|c| c.to_uppercase())
        .eq(b.chars().flat_map(|c| c.to_uppercase()))
}

fn cmp_ascii_case_insensitive(a: &str, b: &str) -> std::cmp::Ordering {
    let mut a_iter = a.as_bytes().iter();
    let mut b_iter = b.as_bytes().iter();
    loop {
        match (a_iter.next(), b_iter.next()) {
            (Some(&ac), Some(&bc)) => {
                let ac = ac.to_ascii_uppercase();
                let bc = bc.to_ascii_uppercase();
                match ac.cmp(&bc) {
                    std::cmp::Ordering::Equal => continue,
                    ord => return ord,
                }
            }
            (None, Some(_)) => return std::cmp::Ordering::Less,
            (Some(_), None) => return std::cmp::Ordering::Greater,
            (None, None) => return std::cmp::Ordering::Equal,
        }
    }
}

fn cmp_case_insensitive(a: &str, b: &str) -> std::cmp::Ordering {
    if a.is_ascii() && b.is_ascii() {
        return cmp_ascii_case_insensitive(a, b);
    }

    let mut a_iter = a.chars().flat_map(|c| c.to_uppercase());
    let mut b_iter = b.chars().flat_map(|c| c.to_uppercase());
    loop {
        match (a_iter.next(), b_iter.next()) {
            (Some(ac), Some(bc)) => match ac.cmp(&bc) {
                std::cmp::Ordering::Equal => continue,
                ord => return ord,
            },
            (None, Some(_)) => return std::cmp::Ordering::Less,
            (Some(_), None) => return std::cmp::Ordering::Greater,
            (None, None) => return std::cmp::Ordering::Equal,
        }
    }
}

fn excel_eq(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Number(x), Value::Number(y)) => x == y,
        (Value::Text(x), Value::Text(y)) => text_eq_case_insensitive(x, y),
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
    fn ordering_to_i32(ord: std::cmp::Ordering) -> i32 {
        match ord {
            std::cmp::Ordering::Less => -1,
            std::cmp::Ordering::Equal => 0,
            std::cmp::Ordering::Greater => 1,
        }
    }

    fn type_rank(v: &Value) -> Option<u8> {
        match v {
            Value::Number(_) => Some(0),
            Value::Text(_) => Some(1),
            Value::Bool(_) => Some(2),
            Value::Blank => Some(3),
            Value::Error(_) => Some(4),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => None,
        }
    }

    match (a, b) {
        // Blank compares like the other type (Excel semantics).
        (Value::Blank, Value::Number(y)) => match 0.0_f64.partial_cmp(y)? {
            std::cmp::Ordering::Less => Some(-1),
            std::cmp::Ordering::Equal => Some(0),
            std::cmp::Ordering::Greater => Some(1),
        },
        (Value::Number(x), Value::Blank) => match x.partial_cmp(&0.0_f64)? {
            std::cmp::Ordering::Less => Some(-1),
            std::cmp::Ordering::Equal => Some(0),
            std::cmp::Ordering::Greater => Some(1),
        },
        (Value::Blank, Value::Text(y)) => Some(match cmp_case_insensitive("", y) {
            std::cmp::Ordering::Less => -1,
            std::cmp::Ordering::Equal => 0,
            std::cmp::Ordering::Greater => 1,
        }),
        (Value::Text(x), Value::Blank) => Some(match cmp_case_insensitive(x, "") {
            std::cmp::Ordering::Less => -1,
            std::cmp::Ordering::Equal => 0,
            std::cmp::Ordering::Greater => 1,
        }),
        (Value::Blank, Value::Bool(y)) => Some(match false.cmp(y) {
            std::cmp::Ordering::Less => -1,
            std::cmp::Ordering::Equal => 0,
            std::cmp::Ordering::Greater => 1,
        }),
        (Value::Bool(x), Value::Blank) => Some(match x.cmp(&false) {
            std::cmp::Ordering::Less => -1,
            std::cmp::Ordering::Equal => 0,
            std::cmp::Ordering::Greater => 1,
        }),
        _ => {
            let ra = type_rank(a)?;
            let rb = type_rank(b)?;
            if ra != rb {
                return Some(ordering_to_i32(ra.cmp(&rb)));
            }

            match (a, b) {
                (Value::Number(x), Value::Number(y)) => Some(ordering_to_i32(x.partial_cmp(y)?)),
                (Value::Text(x), Value::Text(y)) => Some(ordering_to_i32(cmp_case_insensitive(x, y))),
                (Value::Bool(x), Value::Bool(y)) => Some(ordering_to_i32(x.cmp(y))),
                (Value::Blank, Value::Blank) => Some(0),
                (Value::Error(x), Value::Error(y)) => Some(ordering_to_i32(x.as_code().cmp(y.as_code()))),
                _ => None,
            }
        }
    }
}
