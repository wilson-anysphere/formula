use crate::coercion::datetime::parse_value_text;
use crate::eval::CompiledExpr;
use crate::functions::lookup;
use crate::functions::wildcard::WildcardPattern;
use crate::functions::{
    eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec, Reference, SheetId,
};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::pivot::{
    AggregationType as PivotAggregationType, PivotKeyPart, PivotValue as PivotEngineValue,
};
use crate::value::{casefold_owned, with_casefolded_key, Array, ErrorKind, Value};
use chrono::Datelike;
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
                match approximate_match_in_first_col_array(ctx, &lookup_value, &table) {
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
                match approximate_match_in_first_row_array(ctx, &lookup_value, &table) {
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
                    shape, reference, ..
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

        fn xmatch_approx(
            &self,
            ctx: &dyn FunctionContext,
            lookup_value: &Value,
        ) -> Result<i32, ErrorKind> {
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

    fn eval_vector_arg(
        ctx: &dyn FunctionContext,
        expr: &CompiledExpr,
    ) -> Result<LookupVectorArg, ErrorKind> {
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

    fn lookup_array_ref(
        ctx: &dyn FunctionContext,
        lookup_value: &Value,
        reference: &Reference,
    ) -> Value {
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
                1 => approximate_match_values(ctx, &lookup, &arr.values, true),
                -1 => approximate_match_values(ctx, &lookup, &arr.values, false),
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
                XlookupLookupArray::Values { values, .. } => lookup::xmatch_with_modes_with_locale(
                    lookup_value,
                    values,
                    match_mode,
                    search_mode,
                    ctx.value_locale(),
                    ctx.date_system(),
                    ctx.now_utc(),
                ),
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PivotValueColKind {
    Regular,
    GrandTotal,
}

#[derive(Debug, Clone)]
struct PivotValueCol {
    col: u32,
    /// Column key item strings for column-field pivots (e.g. `["A"]` or `["A", "2024"]`).
    ///
    /// Note: the current pivot output flattens column keys into header text and does not include
    /// the source field names. We therefore treat column criteria as matching by item string.
    column_items: Vec<String>,
    kind: PivotValueColKind,
}

struct PivotLayout {
    header_row: u32,
    top_left_col: u32,
    row_fields: HashMap<String, u32>,
    value_cols: Vec<PivotValueCol>,
    /// Map of casefolded rendered header -> index into `value_cols`.
    value_col_by_header: HashMap<String, usize>,
    /// Map of casefolded value field name -> indices into `value_cols`.
    value_cols_by_value_name: HashMap<String, Vec<usize>>,
    has_column_fields: bool,
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
        criteria.push((field, item_value));
    }

    // Prefer the pivot registry (engine-driven metadata) when available.
    if let Some(registry) = ctx.pivot_registry() {
        if let SheetId::Local(sheet_id) = pivot_ref.sheet_id.clone() {
            if let Some(entry) = registry.find_by_cell(sheet_id, pivot_ref.start) {
                // Record a dynamic dependency on the full rendered pivot destination so recalculation
                // can observe pivot refreshes even when the pivot_table argument is a single cell.
                ctx.record_reference(&Reference {
                    sheet_id: SheetId::Local(entry.sheet_id),
                    start: entry.destination.start,
                    end: entry.destination.end,
                });
                return match getpivotdata_from_registry(ctx, entry, &data_field, &criteria) {
                    Ok(v) => v,
                    Err(e) => Value::Error(e),
                };
            }
        }
    }

    // Fallback: heuristic scan of the rendered pivot grid (best-effort).
    let layout = match find_pivot_layout(ctx, &pivot_ref, &data_field) {
        Ok(l) => l,
        Err(e) => return Value::Error(e),
    };

    let mut row_criteria = Vec::new();
    let mut col_criteria = Vec::new();
    for (field, item_value) in criteria {
        let row_col = with_casefolded_key(field.trim(), |key| layout.row_fields.get(key).copied());
        if let Some(col) = row_col {
            row_criteria.push((col, item_value));
            continue;
        }

        // Column-field pivots do not render the column field name(s) in the output grid; only
        // the column item values appear as prefixes in the value headers (e.g. `"A - Sum of Sales"`).
        //
        // For scan-based GETPIVOTDATA, treat any non-row-field criteria as a column-item constraint,
        // but only if we detected column fields in the pivot output.
        if !layout.has_column_fields {
            return Value::Error(ErrorKind::Ref);
        }
        col_criteria.push(item_value);
    }

    let value_col = match select_pivot_value_col(ctx, &layout, &data_field, &col_criteria) {
        Ok(c) => c,
        Err(e) => return Value::Error(e),
    };

    if row_criteria.is_empty() {
        return getpivotdata_grand_total(ctx, &pivot_ref.sheet_id, &layout, value_col);
    }

    match getpivotdata_find_row(ctx, &pivot_ref.sheet_id, &layout, &row_criteria, value_col) {
        Ok(row) => {
            let addr = crate::eval::CellAddr {
                row,
                col: value_col,
            };
            ctx.get_cell_value(&pivot_ref.sheet_id, addr)
        }
        Err(e) => Value::Error(e),
    }
}

#[derive(Debug, Clone)]
struct PivotAccumulator {
    count: u64,
    count_numbers: u64,
    sum: f64,
    product: f64,
    min: f64,
    max: f64,
    mean: f64,
    m2: f64,
}

impl PivotAccumulator {
    fn new() -> Self {
        Self {
            count: 0,
            count_numbers: 0,
            sum: 0.0,
            product: 1.0,
            min: f64::INFINITY,
            max: f64::NEG_INFINITY,
            mean: 0.0,
            m2: 0.0,
        }
    }

    fn update(&mut self, value: &PivotEngineValue) {
        if !matches!(value, PivotEngineValue::Blank) {
            self.count += 1;
        }

        let PivotEngineValue::Number(x) = value else {
            return;
        };
        self.count_numbers += 1;
        self.sum += x;
        self.product *= x;
        if *x < self.min {
            self.min = *x;
        }
        if *x > self.max {
            self.max = *x;
        }

        // Welford variance
        let n = self.count_numbers as f64;
        let delta = x - self.mean;
        self.mean += delta / n;
        let delta2 = x - self.mean;
        self.m2 += delta * delta2;
    }

    fn finalize(&self, agg: PivotAggregationType) -> PivotEngineValue {
        match agg {
            PivotAggregationType::Count => PivotEngineValue::Number(self.count as f64),
            PivotAggregationType::CountNumbers => {
                PivotEngineValue::Number(self.count_numbers as f64)
            }
            PivotAggregationType::Sum => PivotEngineValue::Number(self.sum),
            PivotAggregationType::Product => {
                if self.count_numbers == 0 {
                    PivotEngineValue::Blank
                } else {
                    PivotEngineValue::Number(self.product)
                }
            }
            PivotAggregationType::Average => {
                if self.count_numbers == 0 {
                    PivotEngineValue::Blank
                } else {
                    PivotEngineValue::Number(self.sum / self.count_numbers as f64)
                }
            }
            PivotAggregationType::Min => {
                if self.count_numbers == 0 {
                    PivotEngineValue::Blank
                } else {
                    PivotEngineValue::Number(self.min)
                }
            }
            PivotAggregationType::Max => {
                if self.count_numbers == 0 {
                    PivotEngineValue::Blank
                } else {
                    PivotEngineValue::Number(self.max)
                }
            }
            PivotAggregationType::Var => {
                if self.count_numbers < 2 {
                    PivotEngineValue::Blank
                } else {
                    PivotEngineValue::Number(self.m2 / (self.count_numbers as f64 - 1.0))
                }
            }
            PivotAggregationType::VarP => {
                if self.count_numbers == 0 {
                    PivotEngineValue::Blank
                } else {
                    PivotEngineValue::Number(self.m2 / (self.count_numbers as f64))
                }
            }
            PivotAggregationType::StdDev => {
                if self.count_numbers < 2 {
                    PivotEngineValue::Blank
                } else {
                    PivotEngineValue::Number((self.m2 / (self.count_numbers as f64 - 1.0)).sqrt())
                }
            }
            PivotAggregationType::StdDevP => {
                if self.count_numbers == 0 {
                    PivotEngineValue::Blank
                } else {
                    PivotEngineValue::Number((self.m2 / (self.count_numbers as f64)).sqrt())
                }
            }
        }
    }
}

fn pivot_engine_value_to_key_part(value: &PivotEngineValue) -> PivotKeyPart {
    match value {
        PivotEngineValue::Blank => PivotKeyPart::Blank,
        PivotEngineValue::Number(n) => {
            // Match pivot engine's canonicalization: treat 0.0 and -0.0 as the same item.
            let bits = if *n == 0.0 {
                0.0_f64.to_bits()
            } else if n.is_nan() {
                f64::NAN.to_bits()
            } else {
                n.to_bits()
            };
            PivotKeyPart::Number(bits)
        }
        PivotEngineValue::Date(d) => PivotKeyPart::Date(*d),
        PivotEngineValue::Text(s) => PivotKeyPart::Text(s.clone()),
        PivotEngineValue::Bool(b) => PivotKeyPart::Bool(*b),
    }
}

fn pivot_engine_value_as_value(
    ctx: &dyn FunctionContext,
    value: &PivotEngineValue,
) -> Result<Value, ErrorKind> {
    Ok(match value {
        PivotEngineValue::Blank => Value::Blank,
        PivotEngineValue::Number(n) => Value::Number(*n),
        PivotEngineValue::Text(s) => Value::Text(s.clone()),
        PivotEngineValue::Bool(b) => Value::Bool(*b),
        PivotEngineValue::Date(d) => {
            let date = crate::date::ExcelDate::new(d.year(), d.month() as u8, d.day() as u8);
            let serial =
                crate::date::ymd_to_serial(date, ctx.date_system()).map_err(|_| ErrorKind::Num)?;
            Value::Number(serial as f64)
        }
    })
}

fn getpivotdata_from_registry(
    ctx: &dyn FunctionContext,
    entry: &crate::pivot_registry::PivotRegistryEntry,
    data_field: &str,
    criteria: &[(String, Value)],
) -> Result<Value, ErrorKind> {
    let pivot = entry.pivot.as_ref();
    let data_field = data_field.trim();
    let value_field_idx = with_casefolded_key(data_field, |key| {
        entry.value_field_indices.get(key).copied()
    })
    .ok_or(ErrorKind::Ref)?;
    let value_field = pivot
        .config
        .value_fields
        .get(value_field_idx)
        .ok_or(ErrorKind::Ref)?;
    let value_src_idx = entry
        .value_field_source_indices
        .get(value_field_idx)
        .copied()
        .ok_or(ErrorKind::Ref)?;

    let mut criteria_indices: Vec<(usize, Value)> = Vec::new();
    if criteria_indices.try_reserve_exact(criteria.len()).is_err() {
        debug_assert!(false, "GETPIVOTDATA criteria allocation failed");
        return Err(ErrorKind::Num);
    }
    for (field, item) in criteria {
        let field_trimmed = field.trim();
        if field_trimmed.is_empty() {
            return Err(ErrorKind::Value);
        }
        let normalized = crate::pivot_registry::normalize_pivot_cache_field_name(field_trimmed);
        let (idx, cache_name) = with_casefolded_key(normalized.as_ref(), |key| {
            (
                entry.field_indices.get(key).copied(),
                entry.cache_field_names.get(key),
            )
        });
        let idx = idx.ok_or(ErrorKind::Ref)?;

        // Validate the requested item exists for this field so we can return `#N/A` for unknown
        // items (Excel-compatible; matches the legacy scan behavior).
        let cache_name = cache_name.ok_or(ErrorKind::Ref)?;
        let uniques = pivot
            .cache
            .unique_values
            .get(cache_name)
            .map(|v| v.as_slice())
            .unwrap_or(&[]);
        let mut item_exists = false;
        for pv in uniques {
            let cell = pivot_engine_value_as_value(ctx, pv)?;
            if pivot_item_matches(ctx, &cell, item)? {
                item_exists = true;
                break;
            }
        }
        if !item_exists {
            return Err(ErrorKind::NA);
        }

        criteria_indices.push((idx, item.clone()));
    }

    let mut acc = PivotAccumulator::new();
    let mut pivot_filter_indices: Vec<(usize, &std::collections::HashSet<PivotKeyPart>)> =
        Vec::new();
    if pivot_filter_indices
        .try_reserve(pivot.config.filter_fields.len())
        .is_err()
    {
        debug_assert!(false, "GETPIVOTDATA filter allocation failed");
        return Err(ErrorKind::Num);
    }
    for filter in &pivot.config.filter_fields {
        let Some(allowed) = filter.allowed.as_ref() else {
            continue;
        };
        let field_name = filter.source_field.canonical_name();
        let field_name = crate::pivot_registry::normalize_pivot_cache_field_name(field_name.as_ref());
        let Some(idx) = with_casefolded_key(field_name.as_ref(), |key| {
            entry.field_indices.get(key).copied()
        }) else {
            continue;
        };
        pivot_filter_indices.push((idx, allowed));
    }

    'records: for record in &pivot.cache.records {
        // Apply pivot-config filter fields (report filters / slicers).
        for &(idx, allowed) in &pivot_filter_indices {
            let pv = record.get(idx).unwrap_or(&PivotEngineValue::Blank);
            if !allowed.contains(&pivot_engine_value_to_key_part(pv)) {
                continue 'records;
            }
        }

        for (idx, item) in &criteria_indices {
            let pv = record.get(*idx).unwrap_or(&PivotEngineValue::Blank);
            let cell = pivot_engine_value_as_value(ctx, pv)?;
            if !pivot_item_matches(ctx, &cell, item)? {
                continue 'records;
            }
        }

        let pv = record
            .get(value_src_idx)
            .unwrap_or(&PivotEngineValue::Blank);
        acc.update(pv);
    }

    // NOTE: `GETPIVOTDATA` returns the value field aggregation, not any formatting.
    // `show_as` transformations (percent of total, running total, etc) are not applied here yet.
    let out_pivot_value = acc.finalize(value_field.aggregation);
    pivot_engine_value_as_value(ctx, &out_pivot_value)
}

fn find_pivot_layout(
    ctx: &dyn FunctionContext,
    pivot_ref: &Reference,
    data_field: &str,
) -> Result<PivotLayout, ErrorKind> {
    // Heuristics (scan-based, pivot-engine-compatible):
    //
    // 1) Starting from the provided pivot_table cell, scan upward for a header row containing
    //    a value header that matches `data_field`. Matching is case-insensitive and accepts:
    //      - exact header match (e.g. `"Sum of Sales"` or `"A - Sum of Sales"`)
    //      - column-field headers that end with `" - <data_field>"` (e.g. `"A - Sum of Sales"`)
    //      - grand total headers `"Grand Total - <data_field>"`.
    // 2) From that header cell, scan left across the same row while cells are non-empty text
    //    to locate the pivot table's top-left corner.
    // 3) The row-field columns are the leading columns where the first data row contains text
    //    (row labels). The first non-text cell below the header indicates the start of the
    //    value area.
    // 4) Scan right across the header row to detect the full value area width and build a
    //    mapping from rendered value headers to value columns.
    //
    // MVP safeguards / limitations:
    // - Scan-based heuristic: returns #REF! when the referenced range does not resemble an
    //   engine-produced pivot output.
    // - Supports pivot-engine Tabular layout and a limited form of Compact layout:
    //   - In Compact layout, the row-axis is represented by a single "Row Labels" column. For
    //     pivots with multiple row fields, the pivot engine renders the combined key as
    //     "Field1 / Field2 / …" in that column. GETPIVOTDATA can match against that combined
    //     display string, but does not attempt to interpret Excel-style indentation or repeated
    //     labels.
    // - Supports column fields and multiple value fields by mapping rendered value headers such as:
    //   - `<value name>`
    //   - `<col item> - <value name>`
    //   - `Grand Total - <value name>`
    //   Note: the current pivot output flattens column keys into header text and does not include
    //   the source field names; column criteria are therefore matched by item string.
    // - Scan limits cap work for pathological ranges.
    const MAX_SCAN_ROWS: u32 = 10_000;
    const MAX_SCAN_COLS: u32 = 64;

    let sheet_id = &pivot_ref.sheet_id;
    let anchor = pivot_ref.start;

    let max_up = anchor.row.min(MAX_SCAN_ROWS);
    let col_start = anchor.col.saturating_sub(MAX_SCAN_COLS);
    let col_end = anchor.col.saturating_add(MAX_SCAN_COLS);

    let (header_row, header_match_col) = with_casefolded_key(data_field, |data_field_fold| {
        // Compare against folded headers without allocating composite strings.
        //
        // Pivot output uses `"Grand Total - <value name>"` and `"<col item> - <value name>"`
        // variants, so we match:
        // - exact header
        // - a leading `"GRAND TOTAL - "` prefix (folded domain)
        // - a `" - "` suffix boundary immediately before the value name
        const GRAND_TOTAL_PREFIX_FOLDED: &str = "GRAND TOTAL - ";

        let mut header_row: Option<u32> = None;
        let mut header_match_col: Option<u32> = None;

        for delta in 0..=max_up {
            let row = anchor.row - delta;
            for col in col_start..=col_end {
                let v = ctx.get_cell_value(sheet_id, crate::eval::CellAddr { row, col });
                let Value::Text(t) = v else {
                    continue;
                };
                if t.is_empty() {
                    continue;
                }

                let matches = with_casefolded_key(&t, |t_fold| {
                    if t_fold == data_field_fold {
                        return true;
                    }
                    if t_fold
                        .strip_prefix(GRAND_TOTAL_PREFIX_FOLDED)
                        .is_some_and(|rest| rest == data_field_fold)
                    {
                        return true;
                    }
                    t_fold
                        .strip_suffix(data_field_fold)
                        .is_some_and(|prefix| prefix.ends_with(" - "))
                });
                if !matches {
                    continue;
                }

                let below = ctx.get_cell_value(
                    sheet_id,
                    crate::eval::CellAddr {
                        row: row.saturating_add(1),
                        col,
                    },
                );
                // Pivot-engine value cells are numbers/blanks; if we see a text value directly below,
                // this is unlikely to be the pivot header row.
                if matches!(below, Value::Text(_)) {
                    continue;
                }

                header_row = Some(row);
                header_match_col = Some(col);
                break;
            }
            if header_row.is_some() {
                break;
            }
        }

        (header_row, header_match_col)
    });

    let header_row = header_row.ok_or(ErrorKind::Ref)?;
    let header_match_col = header_match_col.ok_or(ErrorKind::Ref)?;

    // Find top-left header cell by scanning left across the header row.
    let mut top_left_col = header_match_col;
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
    for col in top_left_col..=top_left_col.saturating_add(MAX_SCAN_COLS) {
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
        // Pivot with no row fields isn't supported in the scan-based MVP.
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
        // Duplicate field names would make criteria ambiguous.
        if row_fields.insert(casefold_owned(name), col).is_some() {
            return Err(ErrorKind::Ref);
        }
    }

    // Ensure the header we matched is inside the detected value area.
    if header_match_col < first_value_col {
        return Err(ErrorKind::Ref);
    }

    // Collect value headers (including column-field and grand-total headers).
    let mut value_cols: Vec<PivotValueCol> = Vec::new();
    let mut value_col_by_header: HashMap<String, usize> = HashMap::new();
    let mut value_cols_by_value_name: HashMap<String, Vec<usize>> = HashMap::new();
    let mut has_column_fields = false;

    for col in first_value_col..=first_value_col.saturating_add(MAX_SCAN_COLS) {
        let v = ctx.get_cell_value(
            sheet_id,
            crate::eval::CellAddr {
                row: header_row,
                col,
            },
        );
        let header = match v {
            Value::Text(s) if !s.is_empty() => s,
            _ => break,
        };

        let ParsedPivotValueHeader {
            kind,
            value_name,
            column_items,
        } = parse_pivot_value_header(&header);
        has_column_fields |= !column_items.is_empty();

        let idx = value_cols.len();
        value_cols.push(PivotValueCol {
            col,
            column_items,
            kind,
        });

        if value_col_by_header.insert(casefold_owned(header), idx).is_some() {
            // Duplicate rendered header -> ambiguous, not a supported pivot shape.
            return Err(ErrorKind::Ref);
        }

        value_cols_by_value_name
            .entry(casefold_owned(value_name))
            .or_default()
            .push(idx);
    }

    if value_cols.is_empty() {
        return Err(ErrorKind::Ref);
    }

    Ok(PivotLayout {
        header_row,
        top_left_col,
        row_fields,
        value_cols,
        value_col_by_header,
        value_cols_by_value_name,
        has_column_fields,
    })
}

struct ParsedPivotValueHeader {
    kind: PivotValueColKind,
    value_name: String,
    column_items: Vec<String>,
}

fn parse_pivot_value_header(header: &str) -> ParsedPivotValueHeader {
    // Pivot engine emits grand total columns as `"Grand Total - <value name>"`.
    const GT_PREFIX: &str = "Grand Total - ";
    if header.len() >= GT_PREFIX.len() && header[..GT_PREFIX.len()].eq_ignore_ascii_case(GT_PREFIX)
    {
        return ParsedPivotValueHeader {
            kind: PivotValueColKind::GrandTotal,
            value_name: header[GT_PREFIX.len()..].to_string(),
            column_items: Vec::new(),
        };
    }

    // Regular value columns may include a flattened column key prefix:
    // - no column fields: `<value name>`
    // - with column fields: `<col item> - <value name>` or `<a> / <b> - <value name>`
    if let Some((prefix, value_name)) = header.split_once(" - ") {
        let count = prefix.split(" / ").filter(|s| !s.is_empty()).count();
        let mut column_items: Vec<String> = Vec::new();
        if column_items.try_reserve_exact(count).is_err() {
            debug_assert!(
                false,
                "allocation failed (pivot header column items, len={count})"
            );
            return ParsedPivotValueHeader {
                kind: PivotValueColKind::Regular,
                value_name: value_name.to_string(),
                column_items: Vec::new(),
            };
        }
        for item in prefix.split(" / ").filter(|s| !s.is_empty()) {
            column_items.push(item.to_string());
        }
        return ParsedPivotValueHeader {
            kind: PivotValueColKind::Regular,
            value_name: value_name.to_string(),
            column_items,
        };
    }

    ParsedPivotValueHeader {
        kind: PivotValueColKind::Regular,
        value_name: header.to_string(),
        column_items: Vec::new(),
    }
}

fn select_pivot_value_col(
    ctx: &dyn FunctionContext,
    layout: &PivotLayout,
    data_field: &str,
    col_criteria: &[Value],
) -> Result<u32, ErrorKind> {
    let (exact_header_idx, value_name_candidates) = with_casefolded_key(data_field, |key| {
        (
            layout.value_col_by_header.get(key).copied(),
            layout.value_cols_by_value_name.get(key),
        )
    });

    if !col_criteria.is_empty() {
        let mut col_items: Vec<String> = Vec::new();
        if col_items.try_reserve_exact(col_criteria.len()).is_err() {
            debug_assert!(false, "GETPIVOTDATA column-criteria allocation failed");
            return Err(ErrorKind::Num);
        }
        for v in col_criteria {
            col_items.push(v.coerce_to_string_with_ctx(ctx)?);
        }
        let matches_column_items = |idx: usize| {
            let col = &layout.value_cols[idx];
            col_items.iter().all(|needle| {
                col.column_items
                    .iter()
                    .any(|item| crate::value::eq_case_insensitive(item, needle))
            })
        };

        if let Some(idx) = exact_header_idx {
            if matches_column_items(idx) {
                return Ok(layout.value_cols[idx].col);
            }
            return Err(ErrorKind::NA);
        }

        let Some(candidates) = value_name_candidates else {
            return Err(ErrorKind::Ref);
        };
        let mut found: Option<u32> = None;
        for &idx in candidates {
            if !matches_column_items(idx) {
                continue;
            }
            if found.is_some() {
                return Err(ErrorKind::Ref);
            }
            found = Some(layout.value_cols[idx].col);
        }
        return found.ok_or(ErrorKind::NA);
    }

    if let Some(idx) = exact_header_idx {
        return Ok(layout.value_cols[idx].col);
    }

    let Some(candidates) = value_name_candidates else {
        return Err(ErrorKind::Ref);
    };
    if candidates.is_empty() {
        return Err(ErrorKind::Ref);
    }
    if candidates.len() == 1 {
        return Ok(layout.value_cols[candidates[0]].col);
    }

    // When column fields exist and no column criteria are provided, prefer the row grand total
    // column (`"Grand Total - <value name>"`) if present; otherwise the selection is ambiguous.
    let mut gt: Option<u32> = None;
    for idx in candidates {
        let col = &layout.value_cols[*idx];
        if col.kind == PivotValueColKind::GrandTotal {
            if gt.is_some() {
                return Err(ErrorKind::Ref);
            }
            gt = Some(col.col);
        }
    }
    gt.ok_or(ErrorKind::Ref)
}

fn getpivotdata_grand_total(
    ctx: &dyn FunctionContext,
    sheet_id: &SheetId,
    layout: &PivotLayout,
    value_col: u32,
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
                    col: value_col,
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
                col: value_col,
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
    value_col: u32,
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
                col: value_col,
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

fn pivot_item_matches(
    ctx: &dyn FunctionContext,
    cell: &Value,
    item: &Value,
) -> Result<bool, ErrorKind> {
    match (cell, item) {
        (Value::Error(e), _) => Err(*e),
        (_, Value::Error(e)) => Err(*e),
        (Value::Text(cell_text), _) => {
            let item_text = item.coerce_to_string_with_ctx(ctx)?;
            Ok(crate::value::eq_case_insensitive(cell_text, &item_text))
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

fn wildcard_pattern_for_lookup(
    ctx: &dyn FunctionContext,
    lookup: &Value,
) -> Option<WildcardPattern> {
    // Excel applies wildcard matching to text patterns. Rich values (entities/records) behave
    // like text by using their display string, including `Record.display_field` semantics.
    let pattern = match lookup {
        Value::Text(pattern) => Cow::Borrowed(pattern.as_str()),
        other => match other.coerce_to_string_with_ctx(ctx) {
            Ok(s) => Cow::Owned(s),
            Err(_) => return None,
        },
    };

    if !pattern.contains('*') && !pattern.contains('?') && !pattern.contains('~') {
        return None;
    }
    Some(WildcardPattern::new(pattern.as_ref()))
}

fn exact_match_values(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    values: &[Value],
) -> Option<usize> {
    if let Some(pattern) = wildcard_pattern_for_lookup(ctx, lookup) {
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

fn approximate_match_values(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    values: &[Value],
    ascending: bool,
) -> Option<usize> {
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
            excel_le(ctx, v, lookup)?
        } else {
            excel_ge(ctx, v, lookup)?
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
    let wildcard_pattern = wildcard_pattern_for_lookup(ctx, lookup);
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
    let wildcard_pattern = wildcard_pattern_for_lookup(ctx, lookup);
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
    let wildcard_pattern = wildcard_pattern_for_lookup(ctx, lookup);
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
    let wildcard_pattern = wildcard_pattern_for_lookup(ctx, lookup);
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
        if excel_le(ctx, &key, lookup)? {
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
        if excel_ge(ctx, &key, lookup)? {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1).map(|idx| idx as u32)
}

fn approximate_match_in_first_col_array(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    table: &Array,
) -> Option<u32> {
    if table.rows == 0 {
        return None;
    }

    let mut lo = 0usize;
    let mut hi = table.rows;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let key = table.get(mid, 0).unwrap_or(&Value::Blank);
        if excel_le(ctx, key, lookup)? {
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
        if excel_le(ctx, &key, lookup)? {
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
        if excel_ge(ctx, &key, lookup)? {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1).map(|idx| idx as u32)
}

fn approximate_match_in_first_row_array(
    ctx: &dyn FunctionContext,
    lookup: &Value,
    table: &Array,
) -> Option<u32> {
    if table.cols == 0 {
        return None;
    }

    let mut lo = 0usize;
    let mut hi = table.cols;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let key = table.get(0, mid).unwrap_or(&Value::Blank);
        if excel_le(ctx, key, lookup)? {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }

    lo.checked_sub(1).map(|idx| idx as u32)
}


fn excel_eq(ctx: &dyn FunctionContext, a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Number(x), Value::Number(y)) => x == y,
        (a, b) => {
            fn text_like_str<'a>(
                ctx: &dyn FunctionContext,
                value: &'a Value,
            ) -> Option<std::borrow::Cow<'a, str>> {
                match value {
                    Value::Text(s) => Some(std::borrow::Cow::Borrowed(s.as_str())),
                    Value::Entity(entity) => {
                        Some(std::borrow::Cow::Borrowed(entity.display.as_str()))
                    }
                    Value::Record(record) => match value.coerce_to_string_with_ctx(ctx) {
                        Ok(s) => Some(std::borrow::Cow::Owned(s)),
                        // Be conservative: if the record's `display_field` points at an invalid
                        // value (error/reference/etc), fall back to the stored display string so
                        // lookups still behave deterministically.
                        Err(_) => Some(std::borrow::Cow::Borrowed(record.display.as_str())),
                    },
                    _ => None,
                }
            }

            if let (Some(a), Some(b)) = (text_like_str(ctx, a), text_like_str(ctx, b)) {
                return crate::value::eq_case_insensitive(a.as_ref(), b.as_ref());
            }

            // Excel parses numeric text for comparisons, but does not treat empty strings as zero.
            if let (Value::Number(num), other) | (other, Value::Number(num)) = (a, b) {
                if let Some(other) = text_like_str(ctx, other) {
                    let trimmed = other.trim();
                    if trimmed.is_empty() {
                        return false;
                    }
                    return parse_value_text(
                        trimmed,
                        ctx.value_locale(),
                        ctx.now_utc(),
                        ctx.date_system(),
                    )
                    .is_ok_and(|parsed| parsed == *num);
                }
            }

            match (a, b) {
                (Value::Bool(x), Value::Bool(y)) => x == y,
                (Value::Blank, Value::Blank) => true,
                (Value::Error(x), Value::Error(y)) => x == y,
                (Value::Bool(b), Value::Number(n)) | (Value::Number(n), Value::Bool(b)) => {
                    (*n == 0.0 && !b) || (*n == 1.0 && *b)
                }
                _ => false,
            }
        }
    }
}

fn excel_le(ctx: &dyn FunctionContext, a: &Value, b: &Value) -> Option<bool> {
    excel_cmp(ctx, a, b).map(|o| o <= 0)
}

fn excel_ge(ctx: &dyn FunctionContext, a: &Value, b: &Value) -> Option<bool> {
    excel_cmp(ctx, a, b).map(|o| o >= 0)
}

fn excel_cmp(ctx: &dyn FunctionContext, a: &Value, b: &Value) -> Option<i32> {
    fn ordering_to_i32(ord: std::cmp::Ordering) -> i32 {
        match ord {
            std::cmp::Ordering::Less => -1,
            std::cmp::Ordering::Equal => 0,
            std::cmp::Ordering::Greater => 1,
        }
    }

    fn text_like_str<'a>(
        ctx: &dyn FunctionContext,
        value: &'a Value,
    ) -> Option<std::borrow::Cow<'a, str>> {
        match value {
            Value::Text(s) => Some(std::borrow::Cow::Borrowed(s)),
            Value::Entity(v) => Some(std::borrow::Cow::Borrowed(v.display.as_str())),
            Value::Record(record) => {
                if record.display_field.is_none() {
                    return Some(std::borrow::Cow::Borrowed(record.display.as_str()));
                }
                match value.coerce_to_string_with_ctx(ctx) {
                    Ok(s) => Some(std::borrow::Cow::Owned(s)),
                    Err(_) => Some(std::borrow::Cow::Borrowed(record.display.as_str())),
                }
            }
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
        (Value::Blank, other) if text_like_str(ctx, other).is_some() => {
            let other = text_like_str(ctx, other)?;
            Some(match crate::value::cmp_case_insensitive("", other.as_ref()) {
                std::cmp::Ordering::Less => -1,
                std::cmp::Ordering::Equal => 0,
                std::cmp::Ordering::Greater => 1,
            })
        }
        (other, Value::Blank) if text_like_str(ctx, other).is_some() => {
            let other = text_like_str(ctx, other)?;
            Some(match crate::value::cmp_case_insensitive(other.as_ref(), "") {
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
                (a, b) if text_like_str(ctx, a).is_some() && text_like_str(ctx, b).is_some() => {
                    let a = text_like_str(ctx, a)?;
                    let b = text_like_str(ctx, b)?;
                    Some(ordering_to_i32(crate::value::cmp_case_insensitive(
                        a.as_ref(),
                        b.as_ref(),
                    )))
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
