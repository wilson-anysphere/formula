use crate::coercion::datetime::parse_value_text;
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
    if matches!(lookup_value, Value::Lambda(_)) {
        return Value::Error(ErrorKind::Value);
    }

    let col_index = match eval_scalar_arg(ctx, &args[2]).coerce_to_i64_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    if col_index < 1 {
        return Value::Error(ErrorKind::Value);
    }
    let approx = if args.len() == 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_bool_with_ctx(ctx) {
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
                match exact_match_in_first_col_array(ctx, &lookup_value, &table) {
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
    if matches!(lookup_value, Value::Lambda(_)) {
        return Value::Error(ErrorKind::Value);
    }

    let row_index = match eval_scalar_arg(ctx, &args[2]).coerce_to_i64_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    if row_index < 1 {
        return Value::Error(ErrorKind::Value);
    }
    let approx = if args.len() == 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_bool_with_ctx(ctx) {
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
                match exact_match_in_first_row_array(ctx, &lookup_value, &table) {
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
        name: "LOOKUP",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: lookup_fn,
    }
}

fn lookup_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    #[derive(Debug, Clone)]
    enum LookupVectorArg {
        Values(Vec<Value>),
        Reference {
            shape: XlookupVectorShape,
            reference: Reference,
            len: usize,
        },
    }

    impl LookupVectorArg {
        fn len(&self) -> usize {
            match self {
                LookupVectorArg::Values(values) => values.len(),
                LookupVectorArg::Reference { len, .. } => *len,
            }
        }

        fn get(&self, ctx: &dyn FunctionContext, idx: usize) -> Value {
            match self {
                LookupVectorArg::Values(values) => values.get(idx).cloned().unwrap_or(Value::Blank),
                LookupVectorArg::Reference {
                    shape,
                    reference,
                    ..
                } => {
                    let start = reference.start;
                    match *shape {
                        XlookupVectorShape::Vertical => ctx.get_cell_value(
                            &reference.sheet_id,
                            crate::eval::CellAddr {
                                row: start.row + idx as u32,
                                col: start.col,
                            },
                        ),
                        XlookupVectorShape::Horizontal => ctx.get_cell_value(
                            &reference.sheet_id,
                            crate::eval::CellAddr {
                                row: start.row,
                                col: start.col + idx as u32,
                            },
                        ),
                    }
                }
            }
        }

        fn xmatch_approx(&self, ctx: &dyn FunctionContext, lookup_value: &Value) -> Result<i32, ErrorKind> {
            match self {
                LookupVectorArg::Values(values) => lookup::xmatch_with_modes(
                    lookup_value,
                    values,
                    lookup::MatchMode::ExactOrNextSmaller,
                    lookup::SearchMode::BinaryAscending,
                ),
                LookupVectorArg::Reference {
                    shape,
                    reference,
                    len,
                } => {
                    let sheet_id = &reference.sheet_id;
                    let start = reference.start;
                    lookup::xmatch_with_modes_accessor(
                        lookup_value,
                        *len,
                        |idx| match *shape {
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
                        lookup::MatchMode::ExactOrNextSmaller,
                        lookup::SearchMode::BinaryAscending,
                    )
                }
            }
        }
    }

    fn eval_vector_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<LookupVectorArg, ErrorKind> {
        match ctx.eval_arg(expr) {
            ArgValue::Reference(r) => {
                let r = r.normalized();
                let (shape, len) = reference_1d_shape_len(&r).ok_or(ErrorKind::Value)?;
                // Record dereference for dynamic dependency tracing (e.g. LOOKUP with OFFSET vectors).
                ctx.record_reference(&r);
                Ok(LookupVectorArg::Reference {
                    shape,
                    reference: r,
                    len,
                })
            }
            ArgValue::Scalar(Value::Array(arr)) => {
                if arr.rows != 1 && arr.cols != 1 {
                    return Err(ErrorKind::Value);
                }
                Ok(LookupVectorArg::Values(arr.values))
            }
            ArgValue::Scalar(Value::Error(e)) => Err(e),
            ArgValue::ReferenceUnion(_) | ArgValue::Scalar(_) => Err(ErrorKind::Value),
        }
    }

    fn lookup_array_ref(ctx: &dyn FunctionContext, lookup_value: &Value, reference: &Reference) -> Value {
        let r = reference.normalized();
        // Record dereference for dynamic dependency tracing.
        ctx.record_reference(&r);

        let rows = (r.end.row - r.start.row + 1) as usize;
        let cols = (r.end.col - r.start.col + 1) as usize;
        if rows == 0 || cols == 0 {
            return Value::Error(ErrorKind::NA);
        }

        let search_first_col = rows >= cols;
        let len = if search_first_col { rows } else { cols };
        let sheet_id = &r.sheet_id;
        let start = r.start;
        let end = r.end;
        let pos = lookup::xmatch_with_modes_accessor(
            lookup_value,
            len,
            |idx| {
                if search_first_col {
                    ctx.get_cell_value(
                        sheet_id,
                        crate::eval::CellAddr {
                            row: start.row + idx as u32,
                            col: start.col,
                        },
                    )
                } else {
                    ctx.get_cell_value(
                        sheet_id,
                        crate::eval::CellAddr {
                            row: start.row,
                            col: start.col + idx as u32,
                        },
                    )
                }
            },
            lookup::MatchMode::ExactOrNextSmaller,
            lookup::SearchMode::BinaryAscending,
        );

        let pos = match pos {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let idx = match usize::try_from(pos - 1) {
            Ok(v) => v,
            Err(_) => return Value::Error(ErrorKind::Value),
        };

        if search_first_col {
            ctx.get_cell_value(
                sheet_id,
                crate::eval::CellAddr {
                    row: start.row + idx as u32,
                    col: end.col,
                },
            )
        } else {
            ctx.get_cell_value(
                sheet_id,
                crate::eval::CellAddr {
                    row: end.row,
                    col: start.col + idx as u32,
                },
            )
        }
    }

    let lookup_value = eval_scalar_arg(ctx, &args[0]);
    if let Value::Error(e) = lookup_value {
        return Value::Error(e);
    }
    if matches!(lookup_value, Value::Lambda(_)) {
        return Value::Error(ErrorKind::Value);
    }

    // Vector form: LOOKUP(lookup_value, lookup_vector, [result_vector])
    if args.len() == 3 {
        let lookup_vector = match eval_vector_arg(ctx, &args[1]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let result_vector = match eval_vector_arg(ctx, &args[2]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };

        if lookup_vector.len() != result_vector.len() {
            return Value::Error(ErrorKind::Value);
        }

        let pos = match lookup_vector.xmatch_approx(ctx, &lookup_value) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let idx = match usize::try_from(pos - 1) {
            Ok(v) => v,
            Err(_) => return Value::Error(ErrorKind::Value),
        };
        return result_vector.get(ctx, idx);
    }

    // 2-arg form: could be vector-form or array-form depending on the shape.
    match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(r) => {
            let r_norm = r.normalized();
            if r_norm.start.row == r_norm.end.row || r_norm.start.col == r_norm.end.col {
                // 1D vector.
                let (shape, len) = match reference_1d_shape_len(&r_norm) {
                    Some(v) => v,
                    None => return Value::Error(ErrorKind::Value),
                };
                // Record dereference for dynamic dependency tracing.
                ctx.record_reference(&r_norm);
                let vector = LookupVectorArg::Reference {
                    shape,
                    reference: r_norm,
                    len,
                };
                let pos = match vector.xmatch_approx(ctx, &lookup_value) {
                    Ok(v) => v,
                    Err(e) => return Value::Error(e),
                };
                let idx = match usize::try_from(pos - 1) {
                    Ok(v) => v,
                    Err(_) => return Value::Error(ErrorKind::Value),
                };
                vector.get(ctx, idx)
            } else {
                // 2D array form.
                lookup_array_ref(ctx, &lookup_value, &r_norm)
            }
        }
        ArgValue::Scalar(Value::Array(arr)) => {
            if arr.rows != 1 && arr.cols != 1 {
                return match lookup::lookup_array(&lookup_value, &arr) {
                    Ok(v) => v,
                    Err(e) => Value::Error(e),
                };
            }

            match lookup::lookup_vector(&lookup_value, &arr.values, None) {
                Ok(v) => v,
                Err(e) => Value::Error(e),
            }
        }
        ArgValue::Scalar(Value::Error(e)) => Value::Error(e),
        ArgValue::ReferenceUnion(_) | ArgValue::Scalar(_) => Value::Error(ErrorKind::Value),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "INDEX",
        min_args: 2,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: index_fn,
    }
}

fn index_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    fn index_reference(reference: &Reference, row: i64, col: i64) -> Result<Reference, ErrorKind> {
        let array = reference.normalized();
        let rows = (array.end.row - array.start.row + 1) as i64;
        let cols = (array.end.col - array.start.col + 1) as i64;

        match (row, col) {
            (0, 0) => Ok(array),
            (0, c) => {
                if c < 1 || c > cols {
                    return Err(ErrorKind::Ref);
                }
                let col = array.start.col + (c as u32) - 1;
                Ok(Reference {
                    sheet_id: array.sheet_id,
                    start: crate::eval::CellAddr {
                        row: array.start.row,
                        col,
                    },
                    end: crate::eval::CellAddr {
                        row: array.end.row,
                        col,
                    },
                })
            }
            (r, 0) => {
                if r < 1 || r > rows {
                    return Err(ErrorKind::Ref);
                }
                let row = array.start.row + (r as u32) - 1;
                Ok(Reference {
                    sheet_id: array.sheet_id,
                    start: crate::eval::CellAddr {
                        row,
                        col: array.start.col,
                    },
                    end: crate::eval::CellAddr {
                        row,
                        col: array.end.col,
                    },
                })
            }
            (r, c) => {
                if r < 1 || r > rows || c < 1 || c > cols {
                    return Err(ErrorKind::Ref);
                }
                let addr = crate::eval::CellAddr {
                    row: array.start.row + (r as u32) - 1,
                    col: array.start.col + (c as u32) - 1,
                };
                Ok(Reference {
                    sheet_id: array.sheet_id,
                    start: addr,
                    end: addr,
                })
            }
        }
    }

    fn index_array(arr: Array, row: i64, col: i64) -> Result<Value, ErrorKind> {
        let rows = i64::try_from(arr.rows).unwrap_or(i64::MAX);
        let cols = i64::try_from(arr.cols).unwrap_or(i64::MAX);
        match (row, col) {
            (0, 0) => Ok(Value::Array(arr)),
            (0, c) => {
                if c < 1 || c > cols {
                    return Err(ErrorKind::Ref);
                }
                let c = (c - 1) as usize;
                if arr.rows > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
                    return Err(ErrorKind::Spill);
                }
                let mut values = Vec::new();
                if values.try_reserve_exact(arr.rows).is_err() {
                    return Err(ErrorKind::Num);
                }
                for r in 0..arr.rows {
                    values.push(arr.get(r, c).cloned().unwrap_or(Value::Blank));
                }
                Ok(Value::Array(Array::new(arr.rows, 1, values)))
            }
            (r, 0) => {
                if r < 1 || r > rows {
                    return Err(ErrorKind::Ref);
                }
                let r = (r - 1) as usize;
                if arr.cols > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
                    return Err(ErrorKind::Spill);
                }
                let mut values = Vec::new();
                if values.try_reserve_exact(arr.cols).is_err() {
                    return Err(ErrorKind::Num);
                }
                for c in 0..arr.cols {
                    values.push(arr.get(r, c).cloned().unwrap_or(Value::Blank));
                }
                Ok(Value::Array(Array::new(1, arr.cols, values)))
            }
            (r, c) => {
                if r < 1 || r > rows || c < 1 || c > cols {
                    return Err(ErrorKind::Ref);
                }
                Ok(arr
                    .get((r - 1) as usize, (c - 1) as usize)
                    .cloned()
                    .unwrap_or(Value::Blank))
            }
        }
    }

    let row = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let col = match args.get(2) {
        Some(expr) => match eval_scalar_arg(ctx, expr).coerce_to_i64_with_ctx(ctx) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        },
        None => 1,
    };
    if row < 0 || col < 0 {
        return Value::Error(ErrorKind::Value);
    }

    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(r) => {
            let area_num = match args.get(3) {
                Some(expr) => match eval_scalar_arg(ctx, expr).coerce_to_i64_with_ctx(ctx) {
                    Ok(n) => n,
                    Err(e) => return Value::Error(e),
                },
                None => 1,
            };
            if area_num != 1 {
                return Value::Error(ErrorKind::Ref);
            }
            // Record the referenced input range for dynamic dependency tracing (e.g.
            // INDEX(OFFSET(...), ...)). The evaluator will separately record any dereferenced
            // output cell(s) when the result is consumed.
            ctx.record_reference(&r);
            match index_reference(&r, row, col) {
                Ok(reference) => Value::Reference(reference),
                Err(e) => Value::Error(e),
            }
        }
        ArgValue::ReferenceUnion(ranges) => {
            let area_num = match args.get(3) {
                Some(expr) => match eval_scalar_arg(ctx, expr).coerce_to_i64_with_ctx(ctx) {
                    Ok(n) => n,
                    Err(e) => return Value::Error(e),
                },
                None => 1,
            };
            if area_num < 1 {
                return Value::Error(ErrorKind::Ref);
            }
            let idx = match usize::try_from(area_num - 1) {
                Ok(v) => v,
                Err(_) => return Value::Error(ErrorKind::Ref),
            };
            let Some(r) = ranges.get(idx) else {
                return Value::Error(ErrorKind::Ref);
            };
            // Record the referenced input range for dynamic dependency tracing.
            ctx.record_reference(r);
            match index_reference(r, row, col) {
                Ok(reference) => Value::Reference(reference),
                Err(e) => Value::Error(e),
            }
        }
        ArgValue::Scalar(Value::Array(arr)) => {
            if args.len() == 4 {
                return Value::Error(ErrorKind::Value);
            }
            match index_array(arr, row, col) {
                Ok(v) => v,
                Err(e) => Value::Error(e),
            }
        }
        ArgValue::Scalar(Value::Error(e)) => Value::Error(e),
        ArgValue::Scalar(_) => Value::Error(ErrorKind::Value),
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
    if matches!(lookup, Value::Lambda(_)) {
        return Value::Error(ErrorKind::Value);
    }

    let match_type = if args.len() == 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_i64_with_ctx(ctx) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        }
    } else {
        1
    };

    let pos = match ctx.eval_arg(&args[1]) {
        ArgValue::Reference(r) => {
            let array = r.normalized();
            let (shape, _) = match reference_1d_shape_len(&array) {
                Some(v) => v,
                None => return Value::Error(ErrorKind::NA),
            };
            // Record dereference for dynamic dependency tracing (e.g. MATCH(…, OFFSET(...), …)).
            ctx.record_reference(&array);
            match (shape, match_type) {
                (XlookupVectorShape::Vertical, 0) => {
                    exact_match_in_first_col(ctx, &lookup, &array).map(|v| v as usize)
                }
                (XlookupVectorShape::Horizontal, 0) => {
                    exact_match_in_first_row(ctx, &lookup, &array).map(|v| v as usize)
                }
                (XlookupVectorShape::Vertical, 1) => {
                    approximate_match_in_first_col(ctx, &lookup, &array).map(|v| v as usize)
                }
                (XlookupVectorShape::Horizontal, 1) => {
                    approximate_match_in_first_row(ctx, &lookup, &array).map(|v| v as usize)
                }
                (XlookupVectorShape::Vertical, -1) => {
                    approximate_match_in_first_col_desc(ctx, &lookup, &array).map(|v| v as usize)
                }
                (XlookupVectorShape::Horizontal, -1) => {
                    approximate_match_in_first_row_desc(ctx, &lookup, &array).map(|v| v as usize)
                }
                _ => return Value::Error(ErrorKind::NA),
            }
        }
        ArgValue::Scalar(Value::Array(arr)) => {
            if arr.rows != 1 && arr.cols != 1 {
                return Value::Error(ErrorKind::NA);
            }
            match match_type {
                0 => exact_match_values(ctx, &lookup, &arr.values),
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
    if matches!(lookup_value, Value::Lambda(_)) {
        return Value::Error(ErrorKind::Value);
    }
    let match_mode = match args.get(2) {
        Some(expr) if matches!(expr, CompiledExpr::Blank) => lookup::MatchMode::Exact,
        Some(expr) => match eval_scalar_arg(ctx, expr).coerce_to_i64_with_ctx(ctx) {
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
        Some(expr) => match eval_scalar_arg(ctx, expr).coerce_to_i64_with_ctx(ctx) {
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

            let pos = lookup::xmatch_with_modes_accessor_with_locale(
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
                ctx.value_locale(),
                ctx.date_system(),
                ctx.now_utc(),
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
            match lookup::xmatch_with_modes_with_locale(
                &lookup_value,
                &arr.values,
                match_mode,
                search_mode,
                ctx.value_locale(),
                ctx.date_system(),
                ctx.now_utc(),
            ) {
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
    if matches!(lookup_value, Value::Lambda(_)) {
        return Value::Error(ErrorKind::Value);
    }

    let if_not_found = match args.get(3) {
        Some(expr) if matches!(expr, CompiledExpr::Blank) => None,
        Some(expr) => Some(eval_scalar_arg(ctx, expr)),
        None => None,
    };

    let match_mode = match args.get(4) {
        Some(expr) if matches!(expr, CompiledExpr::Blank) => lookup::MatchMode::Exact,
        Some(expr) => match eval_scalar_arg(ctx, expr).coerce_to_i64_with_ctx(ctx) {
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
        Some(expr) => match eval_scalar_arg(ctx, expr).coerce_to_i64_with_ctx(ctx) {
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
                    lookup::xmatch_with_modes_with_locale(
                        lookup_value,
                        values,
                        match_mode,
                        search_mode,
                        ctx.value_locale(),
                        ctx.date_system(),
                        ctx.now_utc(),
                    )
                }
                XlookupLookupArray::Reference {
                    shape,
                    reference,
                    len,
                } => {
                    let sheet_id = &reference.sheet_id;
                    let start = reference.start;
                    lookup::xmatch_with_modes_accessor_with_locale(
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
                        ctx.value_locale(),
                        ctx.date_system(),
                        ctx.now_utc(),
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
                XlookupReturnArray::Array(arr) => {
                    arr.get(row, col).cloned().unwrap_or(Value::Blank)
                }
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
            if cols > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut values = Vec::new();
            if values.try_reserve_exact(cols).is_err() {
                return Value::Error(ErrorKind::Num);
            }
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
            if rows > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
                return Value::Error(ErrorKind::Spill);
            }
            let mut values = Vec::new();
            if values.try_reserve_exact(rows).is_err() {
                return Value::Error(ErrorKind::Num);
            }
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
    let data_field = match data_field_value.coerce_to_string_with_ctx(ctx) {
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
        let field = match field_value.coerce_to_string_with_ctx(ctx) {
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
    // - Supports pivot-engine Tabular and Compact layouts.
    //   - In Compact layout, the row-axis is represented by a single "Row Labels" column. For
    //     pivots with multiple row fields, our pivot engine renders the combined key as
    //     "Field1 / Field2 / …" in that column. GETPIVOTDATA can match against that combined
    //     display string, but does not attempt to interpret Excel-style indentation or repeated
    //     labels.
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
            if !pivot_item_matches(ctx, &cell, item)? {
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

fn pivot_item_matches(ctx: &dyn FunctionContext, cell: &Value, item: &Value) -> Result<bool, ErrorKind> {
    match (cell, item) {
        (Value::Error(e), _) => Err(*e),
        (_, Value::Error(e)) => Err(*e),
        (Value::Text(cell_text), _) => {
            let item_text = item.coerce_to_string_with_ctx(ctx)?;
            Ok(cell_text.eq_ignore_ascii_case(&item_text))
        }
        (Value::Blank, _) => {
            let item_text = item.coerce_to_string_with_ctx(ctx)?;
            Ok(item_text.is_empty())
        }
        _ => Ok(excel_eq(ctx, cell, item)),
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

fn exact_match_values(ctx: &dyn FunctionContext, lookup: &Value, values: &[Value]) -> Option<usize> {
    if let Some(pattern) = wildcard_pattern_for_lookup(lookup) {
        for (idx, candidate) in values.iter().enumerate() {
            let text = match candidate {
                Value::Error(_) => continue,
                Value::Text(s) => Cow::Borrowed(s.as_str()),
                other => match other.coerce_to_string_with_ctx(ctx) {
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

    values.iter().position(|v| excel_eq(ctx, lookup, v))
}

fn approximate_match_values(lookup: &Value, values: &[Value], ascending: bool) -> Option<usize> {
    if values.is_empty() {
        return None;
    }

    // Excel's approximate matching behaves like a binary search over a sorted array:
    // - ascending: last index where value <= lookup
    // - descending: last index where value >= lookup
    let mut lo = 0usize;
    let mut hi = values.len();
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let v = &values[mid];
        let ok = if ascending {
            excel_le(v, lookup)?
        } else {
            excel_ge(v, lookup)?
        };
        if ok {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1)
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
                other => match other.coerce_to_string_with_ctx(ctx) {
                    Ok(s) => Cow::Owned(s),
                    Err(_) => continue,
                },
            };
            if pattern.matches(text.as_ref()) {
                return Some(idx as u32);
            }
        } else if excel_eq(ctx, lookup, &key) {
            return Some(idx as u32);
        }
    }
    None
}

fn exact_match_in_first_col_array(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    table: &Array,
) -> Option<u32> {
    let wildcard_pattern = wildcard_pattern_for_lookup(lookup);
    for row in 0..table.rows {
        let key = table.get(row, 0).unwrap_or(&Value::Blank);
        if let Some(pattern) = &wildcard_pattern {
            let text = match key {
                Value::Error(_) => continue,
                Value::Text(s) => Cow::Borrowed(s.as_str()),
                other => match other.coerce_to_string_with_ctx(ctx) {
                    Ok(s) => Cow::Owned(s),
                    Err(_) => continue,
                },
            };
            if pattern.matches(text.as_ref()) {
                return Some(row as u32);
            }
        } else if excel_eq(ctx, lookup, key) {
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
                other => match other.coerce_to_string_with_ctx(ctx) {
                    Ok(s) => Cow::Owned(s),
                    Err(_) => continue,
                },
            };
            if pattern.matches(text.as_ref()) {
                return Some(idx as u32);
            }
        } else if excel_eq(ctx, lookup, &key) {
            return Some(idx as u32);
        }
    }
    None
}

fn exact_match_in_first_row_array(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    table: &Array,
) -> Option<u32> {
    let wildcard_pattern = wildcard_pattern_for_lookup(lookup);
    for col in 0..table.cols {
        let key = table.get(0, col).unwrap_or(&Value::Blank);
        if let Some(pattern) = &wildcard_pattern {
            let text = match key {
                Value::Error(_) => continue,
                Value::Text(s) => Cow::Borrowed(s.as_str()),
                other => match other.coerce_to_string_with_ctx(ctx) {
                    Ok(s) => Cow::Owned(s),
                    Err(_) => continue,
                },
            };
            if pattern.matches(text.as_ref()) {
                return Some(col as u32);
            }
        } else if excel_eq(ctx, lookup, key) {
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
    let len = (table.end.row - table.start.row + 1) as usize;
    if len == 0 {
        return None;
    }

    // Find the insertion point after the last key that is <= lookup.
    let mut lo = 0usize;
    let mut hi = len;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let addr = crate::eval::CellAddr {
            row: table.start.row + mid as u32,
            col: table.start.col,
        };
        let key = ctx.get_cell_value(&table.sheet_id, addr);
        if excel_le(&key, lookup)? {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1).map(|idx| idx as u32)
}

fn approximate_match_in_first_col_desc(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    table: &Reference,
) -> Option<u32> {
    let len = (table.end.row - table.start.row + 1) as usize;
    if len == 0 {
        return None;
    }

    // Descending approximate match: insertion point after the last key that is >= lookup.
    let mut lo = 0usize;
    let mut hi = len;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let addr = crate::eval::CellAddr {
            row: table.start.row + mid as u32,
            col: table.start.col,
        };
        let key = ctx.get_cell_value(&table.sheet_id, addr);
        if excel_ge(&key, lookup)? {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1).map(|idx| idx as u32)
}

fn approximate_match_in_first_col_array(lookup: &Value, table: &Array) -> Option<u32> {
    if table.rows == 0 {
        return None;
    }

    let mut lo = 0usize;
    let mut hi = table.rows;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let key = table.get(mid, 0).unwrap_or(&Value::Blank);
        if excel_le(key, lookup)? {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1).map(|idx| idx as u32)
}

fn approximate_match_in_first_row(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    table: &Reference,
) -> Option<u32> {
    let len = (table.end.col - table.start.col + 1) as usize;
    if len == 0 {
        return None;
    }

    let mut lo = 0usize;
    let mut hi = len;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let addr = crate::eval::CellAddr {
            row: table.start.row,
            col: table.start.col + mid as u32,
        };
        let key = ctx.get_cell_value(&table.sheet_id, addr);
        if excel_le(&key, lookup)? {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1).map(|idx| idx as u32)
}

fn approximate_match_in_first_row_desc(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    table: &Reference,
) -> Option<u32> {
    let len = (table.end.col - table.start.col + 1) as usize;
    if len == 0 {
        return None;
    }

    let mut lo = 0usize;
    let mut hi = len;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let addr = crate::eval::CellAddr {
            row: table.start.row,
            col: table.start.col + mid as u32,
        };
        let key = ctx.get_cell_value(&table.sheet_id, addr);
        if excel_ge(&key, lookup)? {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1).map(|idx| idx as u32)
}

fn approximate_match_in_first_row_array(lookup: &Value, table: &Array) -> Option<u32> {
    if table.cols == 0 {
        return None;
    }

    let mut lo = 0usize;
    let mut hi = table.cols;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let key = table.get(0, mid).unwrap_or(&Value::Blank);
        if excel_le(key, lookup)? {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1).map(|idx| idx as u32)
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

fn excel_eq(ctx: &dyn FunctionContext, a: &Value, b: &Value) -> bool {
    fn text_like_str(v: &Value) -> Option<&str> {
        match v {
            Value::Text(s) => Some(s.as_str()),
            Value::Entity(v) => Some(v.display.as_str()),
            Value::Record(v) => Some(v.display.as_str()),
            _ => None,
        }
    }

    match (a, b) {
        (Value::Number(x), Value::Number(y)) => x == y,
        (a, b) if text_like_str(a).is_some() && text_like_str(b).is_some() => {
            text_eq_case_insensitive(text_like_str(a).unwrap(), text_like_str(b).unwrap())
        }
        (Value::Bool(x), Value::Bool(y)) => x == y,
        (Value::Blank, Value::Blank) => true,
        (Value::Error(x), Value::Error(y)) => x == y,
        (Value::Number(num), other) | (other, Value::Number(num)) if text_like_str(other).is_some() => {
            let trimmed = text_like_str(other).unwrap().trim();
            if trimmed.is_empty() {
                false
            } else {
                parse_value_text(
                    trimmed,
                    ctx.value_locale(),
                    ctx.now_utc(),
                    ctx.date_system(),
                )
                .is_ok_and(|parsed| parsed == *num)
            }
        }
        (Value::Bool(b), Value::Number(n)) | (Value::Number(n), Value::Bool(b)) => {
            (*n == 0.0 && !b) || (*n == 1.0 && *b)
        }
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

    fn text_like_str(v: &Value) -> Option<std::borrow::Cow<'_, str>> {
        match v {
            Value::Text(s) => Some(std::borrow::Cow::Borrowed(s)),
            Value::Entity(v) => Some(std::borrow::Cow::Borrowed(v.display.as_str())),
            Value::Record(v) => Some(std::borrow::Cow::Borrowed(v.display.as_str())),
            _ => None,
        }
    }

    fn type_rank(v: &Value) -> Option<u8> {
        match v {
            Value::Number(_) => Some(0),
            Value::Text(_) | Value::Entity(_) | Value::Record(_) => Some(1),
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
        (Value::Blank, other) if text_like_str(other).is_some() => {
            let other = text_like_str(other)?;
            Some(match cmp_case_insensitive("", other.as_ref()) {
                std::cmp::Ordering::Less => -1,
                std::cmp::Ordering::Equal => 0,
                std::cmp::Ordering::Greater => 1,
            })
        }
        (other, Value::Blank) if text_like_str(other).is_some() => {
            let other = text_like_str(other)?;
            Some(match cmp_case_insensitive(other.as_ref(), "") {
                std::cmp::Ordering::Less => -1,
                std::cmp::Ordering::Equal => 0,
                std::cmp::Ordering::Greater => 1,
            })
        }
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
                (a, b) if text_like_str(a).is_some() && text_like_str(b).is_some() => {
                    let a = text_like_str(a)?;
                    let b = text_like_str(b)?;
                    Some(ordering_to_i32(cmp_case_insensitive(a.as_ref(), b.as_ref())))
                }
                (Value::Bool(x), Value::Bool(y)) => Some(ordering_to_i32(x.cmp(y))),
                (Value::Blank, Value::Blank) => Some(0),
                (Value::Error(x), Value::Error(y)) => {
                    Some(ordering_to_i32(x.code().cmp(&y.code())))
                }
                _ => None,
            }
        }
    }
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
