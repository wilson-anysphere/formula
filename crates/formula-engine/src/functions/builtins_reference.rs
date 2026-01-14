use crate::eval::{CompiledExpr, Expr};
use crate::functions::array_lift;
use crate::functions::{
    eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec, ThreadSafety,
    ValueType, Volatility,
};
use crate::value::{Array, ErrorKind, Value};

inventory::submit! {
    FunctionSpec {
        name: "OFFSET",
        min_args: 3,
        max_args: 5,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: offset_fn,
    }
}

fn offset_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let base = match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(r) => r,
        ArgValue::ReferenceUnion(_) => return Value::Error(ErrorKind::Value),
        ArgValue::Scalar(Value::Error(e)) => return Value::Error(e),
        ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
    };

    let rows = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let cols = match eval_scalar_arg(ctx, &args[2]).coerce_to_i64_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let base_norm = base.normalized();
    let default_height = (base_norm.end.row as i64) - (base_norm.start.row as i64) + 1;
    let default_width = (base_norm.end.col as i64) - (base_norm.start.col as i64) + 1;

    let height = if args.len() >= 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_i64_with_ctx(ctx) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        default_height
    };

    let width = if args.len() >= 5 {
        match eval_scalar_arg(ctx, &args[4]).coerce_to_i64_with_ctx(ctx) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        default_width
    };

    if height < 1 || width < 1 {
        return Value::Error(ErrorKind::Ref);
    }

    let start_row = (base_norm.start.row as i64) + rows;
    let start_col = (base_norm.start.col as i64) + cols;
    let end_row = start_row + height - 1;
    let end_col = start_col + width - 1;

    // Excel treats OFFSET results that point outside the current sheet dimensions as `#REF!`.
    // Sheet dimensions are dynamic in this engine, so we only validate that the coordinates are
    // representable and non-negative here; the evaluator will apply sheet-dimension-based bounds
    // checks uniformly for all reference values (including those produced by OFFSET).
    //
    // `u32::MAX` is reserved as a sheet-end sentinel for whole-row/whole-column references, so
    // disallow it here (OFFSET should return `#REF!` rather than silently snapping to sheet end).
    let within_u32 = |n: i64| n >= 0 && n < (crate::eval::CellAddr::SHEET_END as i64);
    if !within_u32(start_row)
        || !within_u32(start_col)
        || !within_u32(end_row)
        || !within_u32(end_col)
    {
        return Value::Error(ErrorKind::Ref);
    }

    Value::Reference(crate::functions::Reference {
        sheet_id: base_norm.sheet_id,
        start: crate::eval::CellAddr {
            row: start_row as u32,
            col: start_col as u32,
        },
        end: crate::eval::CellAddr {
            row: end_row as u32,
            col: end_col as u32,
        },
    })
}

inventory::submit! {
    FunctionSpec {
        name: "ROW",
        min_args: 0,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: row_fn,
    }
}

fn row_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.is_empty() {
        return Value::Number((u64::from(ctx.current_cell_addr().row) + 1) as f64);
    }

    let reference = match reference_from_arg(ctx.eval_arg(&args[0])) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let reference = reference.normalized();

    if reference.is_single_cell() {
        return Value::Number((u64::from(reference.start.row) + 1) as f64);
    }

    let rows = (reference.end.row - reference.start.row + 1) as usize;
    let cols = (reference.end.col - reference.start.col + 1) as usize;

    let (sheet_rows, sheet_cols) = ctx.sheet_dimensions(&reference.sheet_id);
    let spans_all_cols =
        reference.start.col == 0 && reference.end.col == sheet_cols.saturating_sub(1);
    let spans_all_rows =
        reference.start.row == 0 && reference.end.row == sheet_rows.saturating_sub(1);

    if spans_all_cols || spans_all_rows {
        if rows > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
            return Value::Error(ErrorKind::Spill);
        }
        let mut values = Vec::new();
        if values.try_reserve_exact(rows).is_err() {
            return Value::Error(ErrorKind::Num);
        }
        for row in reference.start.row..=reference.end.row {
            values.push(Value::Number((u64::from(row) + 1) as f64));
        }
        if rows == 1 {
            return values.first().cloned().unwrap_or(Value::Blank);
        }
        return Value::Array(Array::new(rows, 1, values));
    }

    let total = match rows.checked_mul(cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for row in reference.start.row..=reference.end.row {
        let n = Value::Number((u64::from(row) + 1) as f64);
        for _ in reference.start.col..=reference.end.col {
            values.push(n.clone());
        }
    }
    Value::Array(Array::new(rows, cols, values))
}

inventory::submit! {
    FunctionSpec {
        name: "COLUMN",
        min_args: 0,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: column_fn,
    }
}

fn column_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.is_empty() {
        return Value::Number((u64::from(ctx.current_cell_addr().col) + 1) as f64);
    }

    let reference = match reference_from_arg(ctx.eval_arg(&args[0])) {
        Ok(v) => v,
        Err(e) => return e,
    };
    let reference = reference.normalized();

    if reference.is_single_cell() {
        return Value::Number((u64::from(reference.start.col) + 1) as f64);
    }

    let rows = (reference.end.row - reference.start.row + 1) as usize;
    let cols = (reference.end.col - reference.start.col + 1) as usize;

    let (sheet_rows, sheet_cols) = ctx.sheet_dimensions(&reference.sheet_id);
    let spans_all_cols =
        reference.start.col == 0 && reference.end.col == sheet_cols.saturating_sub(1);
    let spans_all_rows =
        reference.start.row == 0 && reference.end.row == sheet_rows.saturating_sub(1);

    if spans_all_cols || spans_all_rows {
        if cols > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
            return Value::Error(ErrorKind::Spill);
        }
        let mut values = Vec::new();
        if values.try_reserve_exact(cols).is_err() {
            return Value::Error(ErrorKind::Num);
        }
        for col in reference.start.col..=reference.end.col {
            values.push(Value::Number((u64::from(col) + 1) as f64));
        }
        if cols == 1 {
            return values.first().cloned().unwrap_or(Value::Blank);
        }
        return Value::Array(Array::new(1, cols, values));
    }

    let total = match rows.checked_mul(cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for _ in 0..rows {
        for col in reference.start.col..=reference.end.col {
            values.push(Value::Number((u64::from(col) + 1) as f64));
        }
    }
    Value::Array(Array::new(rows, cols, values))
}

inventory::submit! {
    FunctionSpec {
        name: "ROWS",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: rows_fn,
    }
}

fn rows_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(r) => {
            let r = r.normalized();
            let rows = u64::from(r.end.row).saturating_sub(u64::from(r.start.row)) + 1;
            Value::Number(rows as f64)
        }
        ArgValue::ReferenceUnion(_) => Value::Error(ErrorKind::Value),
        ArgValue::Scalar(Value::Reference(r)) => {
            let r = r.normalized();
            let rows = u64::from(r.end.row).saturating_sub(u64::from(r.start.row)) + 1;
            Value::Number(rows as f64)
        }
        ArgValue::Scalar(Value::Array(arr)) => Value::Number(arr.rows as f64),
        ArgValue::Scalar(Value::Error(e)) => Value::Error(e),
        ArgValue::Scalar(_) => Value::Error(ErrorKind::Value),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "COLUMNS",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: columns_fn,
    }
}

fn columns_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(r) => {
            let r = r.normalized();
            let cols = u64::from(r.end.col).saturating_sub(u64::from(r.start.col)) + 1;
            Value::Number(cols as f64)
        }
        ArgValue::ReferenceUnion(_) => Value::Error(ErrorKind::Value),
        ArgValue::Scalar(Value::Reference(r)) => {
            let r = r.normalized();
            let cols = u64::from(r.end.col).saturating_sub(u64::from(r.start.col)) + 1;
            Value::Number(cols as f64)
        }
        ArgValue::Scalar(Value::Array(arr)) => Value::Number(arr.cols as f64),
        ArgValue::Scalar(Value::Error(e)) => Value::Error(e),
        ArgValue::Scalar(_) => Value::Error(ErrorKind::Value),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "AREAS",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: areas_fn,
    }
}

fn areas_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(_) => Value::Number(1.0),
        ArgValue::ReferenceUnion(ranges) => Value::Number(ranges.len() as f64),
        ArgValue::Scalar(Value::Reference(_)) => Value::Number(1.0),
        ArgValue::Scalar(Value::ReferenceUnion(ranges)) => Value::Number(ranges.len() as f64),
        ArgValue::Scalar(Value::Error(e)) => Value::Error(e),
        ArgValue::Scalar(_) => Value::Error(ErrorKind::Value),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "ADDRESS",
        min_args: 2,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Bool, ValueType::Text],
        implementation: address_fn,
    }
}

fn address_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let row_num = array_lift::eval_arg(ctx, &args[0]);
    let col_num = array_lift::eval_arg(ctx, &args[1]);
    let abs_num = if args.len() >= 3 && !matches!(args[2], Expr::Blank) {
        array_lift::eval_arg(ctx, &args[2])
    } else {
        Value::Number(1.0)
    };
    let a1 = if args.len() >= 4 && !matches!(args[3], Expr::Blank) {
        array_lift::eval_arg(ctx, &args[3])
    } else {
        Value::Bool(true)
    };

    let sheet_prefix = if args.len() >= 5 && !matches!(args[4], Expr::Blank) {
        match eval_scalar_arg(ctx, &args[4]) {
            Value::Error(e) => return Value::Error(e),
            Value::Array(_)
            | Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            other => {
                let raw = match other.coerce_to_string_with_ctx(ctx) {
                    Ok(s) => s,
                    Err(e) => return Value::Error(e),
                };
                if raw.is_empty() {
                    None
                } else {
                    Some(format!("{}!", quote_sheet_name(&raw)))
                }
            }
        }
    } else {
        None
    };

    let current_sheet = crate::functions::SheetId::Local(ctx.current_sheet_id());
    let (sheet_rows, sheet_cols) = ctx.sheet_dimensions(&current_sheet);

    let base = array_lift::lift4(row_num, col_num, abs_num, a1, |row, col, abs, a1| {
        let row_num = row.coerce_to_i64_with_ctx(ctx)?;
        let col_num = col.coerce_to_i64_with_ctx(ctx)?;
        if row_num < 1 || row_num > sheet_rows as i64 {
            return Err(ErrorKind::Value);
        }
        if col_num < 1 || col_num > sheet_cols as i64 {
            return Err(ErrorKind::Value);
        }

        let abs_num = abs.coerce_to_i64_with_ctx(ctx)?;
        let (col_abs, row_abs) = match abs_num {
            1 => (true, true),
            2 => (false, true),
            3 => (true, false),
            4 => (false, false),
            _ => return Err(ErrorKind::Value),
        };

        let a1 = a1.coerce_to_bool_with_ctx(ctx)?;
        let address = if a1 {
            format_a1_address(row_num as u32, col_num as u32, row_abs, col_abs)
        } else {
            format_r1c1_address(row_num as i64, col_num as i64, row_abs, col_abs)
        };
        Ok(Value::Text(address))
    });

    let Some(prefix) = sheet_prefix else {
        return base;
    };

    array_lift::lift1(base, |v| match v {
        Value::Error(e) => Ok(Value::Error(*e)),
        Value::Text(s) => Ok(Value::Text(format!("{prefix}{s}"))),
        _ => Err(ErrorKind::Value),
    })
}

fn reference_from_arg(arg: ArgValue) -> Result<crate::functions::Reference, Value> {
    match arg {
        ArgValue::Reference(r) => Ok(r),
        ArgValue::ReferenceUnion(_) => Err(Value::Error(ErrorKind::Value)),
        ArgValue::Scalar(Value::Reference(r)) => Ok(r),
        ArgValue::Scalar(Value::ReferenceUnion(_)) => Err(Value::Error(ErrorKind::Value)),
        ArgValue::Scalar(Value::Error(e)) => Err(Value::Error(e)),
        ArgValue::Scalar(_) => Err(Value::Error(ErrorKind::Value)),
    }
}

fn is_ident_cont_char(c: char) -> bool {
    matches!(c, '$' | '_' | '\\' | '.' | 'A'..='Z' | 'a'..='z' | '0'..='9')
}

fn starts_like_a1_cell_ref(s: &str) -> bool {
    // The lexer tokenizes A1-style cell references even when followed by additional identifier
    // characters (e.g. `A1B`), so treat any sheet name *starting* with a valid A1 ref as requiring
    // quotes. This matches `ast.rs` sheet-name formatting rules.
    let bytes = s.as_bytes();
    let mut i = 0;
    if bytes.get(i) == Some(&b'$') {
        i += 1;
    }

    let start_letters = i;
    while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
        i += 1;
    }
    if i == start_letters {
        return false;
    }

    if bytes.get(i) == Some(&b'$') {
        i += 1;
    }

    let start_digits = i;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i == start_digits {
        return false;
    }

    crate::eval::parse_a1(&s[..i]).is_ok()
}

fn quote_sheet_name(name: &str) -> String {
    if name.is_empty() {
        return String::new();
    }

    let starts_like_number = matches!(name.chars().next(), Some('0'..='9' | '.'));
    let starts_like_r1c1 = matches!(name.chars().next(), Some('R' | 'r' | 'C' | 'c'))
        && matches!(name.chars().nth(1), Some('0'..='9' | '['));
    let starts_like_a1 = starts_like_a1_cell_ref(name);
    // The formula lexer treats TRUE/FALSE as booleans rather than identifiers; quoting is required
    // to disambiguate sheet names that match those keywords.
    let is_reserved = name.eq_ignore_ascii_case("TRUE") || name.eq_ignore_ascii_case("FALSE");
    let needs_quote = starts_like_number
        || is_reserved
        || starts_like_r1c1
        || starts_like_a1
        || name.chars().any(|c| !is_ident_cont_char(c));

    if !needs_quote {
        return name.to_string();
    }

    let escaped = name.replace('\'', "''");
    format!("'{escaped}'")
}

fn col_to_name(col: u32) -> String {
    let mut n = col;
    let mut out = Vec::<u8>::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        out.push(b'A' + rem as u8);
        n = (n - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).expect("column letters are always valid UTF-8")
}

fn format_a1_address(row_num: u32, col_num: u32, row_abs: bool, col_abs: bool) -> String {
    let mut out = String::new();
    if col_abs {
        out.push('$');
    }
    out.push_str(&col_to_name(col_num));
    if row_abs {
        out.push('$');
    }
    out.push_str(&row_num.to_string());
    out
}

fn format_r1c1_address(row_num: i64, col_num: i64, row_abs: bool, col_abs: bool) -> String {
    let mut out = String::new();
    if row_abs {
        out.push('R');
        out.push_str(&row_num.to_string());
    } else {
        out.push_str("R[");
        out.push_str(&row_num.to_string());
        out.push(']');
    }
    if col_abs {
        out.push('C');
        out.push_str(&col_num.to_string());
    } else {
        out.push_str("C[");
        out.push_str(&col_num.to_string());
        out.push(']');
    }
    out
}

inventory::submit! {
    FunctionSpec {
        name: "INDIRECT",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Text, ValueType::Bool],
        implementation: indirect_fn,
    }
}

fn indirect_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let a1 = if args.len() >= 2 {
        match eval_scalar_arg(ctx, &args[1]).coerce_to_bool_with_ctx(ctx) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        true
    };

    let ref_text = text.trim();
    if ref_text.is_empty() {
        return Value::Error(ErrorKind::Ref);
    }

    // Parse the text as a standalone reference expression using the canonical parser so we
    // support A1/R1C1, quoting, and range operators.
    let parsed = match crate::parse_formula(
        ref_text,
        crate::ParseOptions {
            locale: crate::LocaleConfig::en_us(),
            reference_style: if a1 {
                crate::ReferenceStyle::A1
            } else {
                crate::ReferenceStyle::R1C1
            },
            normalize_relative_to: None,
        },
    ) {
        Ok(ast) => ast,
        Err(_) => return Value::Error(ErrorKind::Ref),
    };

    let origin = ctx.current_cell_addr();
    let origin_ast = crate::CellAddr::new(origin.row, origin.col);
    let lowered = crate::eval::lower_ast(&parsed, if a1 { None } else { Some(origin_ast) });

    match lowered {
        // Validate that the parsed expression is a "simple" static reference (cell or rectangular
        // range) before compiling it. This preserves the historical behavior of rejecting unions,
        // intersections, defined names, structured refs, etc.
        crate::eval::Expr::CellRef(_) | crate::eval::Expr::RangeRef(_) => {
            let mut resolve_sheet = |name: &str| {
                // Avoid interpreting bracketed external sheet keys like `"[Book.xlsx]Sheet1"` as a
                // local sheet name (a workbook could contain such a sheet name). The canonical
                // parser represents external workbook references separately, so they do not go
                // through this resolver.
                if name.starts_with('[') {
                    return None;
                }
                ctx.resolve_sheet_name(name)
            };
            let mut sheet_dimensions =
                |sheet_id: usize| ctx.sheet_dimensions(&crate::functions::SheetId::Local(sheet_id));
            let compiled = crate::eval::compile_canonical_expr(
                &parsed.expr,
                ctx.current_sheet_id(),
                ctx.current_cell_addr(),
                &mut resolve_sheet,
                &mut sheet_dimensions,
            );

            match ctx.eval_arg(&compiled) {
                ArgValue::Reference(r) => {
                    // INDIRECT supports single-sheet external workbook references like
                    // `"[Book.xlsx]Sheet1"`, but rejects external 3D spans like
                    // `"[Book.xlsx]Sheet1:Sheet3"`.
                    match &r.sheet_id {
                        crate::functions::SheetId::External(key)
                            if !crate::eval::is_valid_external_sheet_key(key) =>
                        {
                            Value::Error(ErrorKind::Ref)
                        }
                        _ => Value::Reference(r),
                    }
                }
                ArgValue::ReferenceUnion(_) => Value::Error(ErrorKind::Ref),
                ArgValue::Scalar(Value::Error(e)) => Value::Error(e),
                _ => Value::Error(ErrorKind::Ref),
            }
        }
        crate::eval::Expr::NameRef(_) => Value::Error(ErrorKind::Ref),
        crate::eval::Expr::Error(e) => Value::Error(e),
        _ => Value::Error(ErrorKind::Ref),
    }
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
