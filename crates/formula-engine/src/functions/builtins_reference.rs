use crate::eval::{CompiledExpr, SheetReference};
use crate::functions::{
    eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec, ThreadSafety,
    ValueType, Volatility,
};
use crate::value::{ErrorKind, Value};

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

    let rows = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let cols = match eval_scalar_arg(ctx, &args[2]).coerce_to_i64() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let base_norm = base.normalized();
    let default_height = (base_norm.end.row as i64) - (base_norm.start.row as i64) + 1;
    let default_width = (base_norm.end.col as i64) - (base_norm.start.col as i64) + 1;

    let height = if args.len() >= 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_i64() {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        default_height
    };

    let width = if args.len() >= 5 {
        match eval_scalar_arg(ctx, &args[4]).coerce_to_i64() {
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

    let within_bounds = |row: i64, col: i64| {
        row >= 0
            && col >= 0
            && row < formula_model::EXCEL_MAX_ROWS as i64
            && col < formula_model::EXCEL_MAX_COLS as i64
    };
    if !within_bounds(start_row, start_col) || !within_bounds(end_row, end_col) {
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
    let text = match eval_scalar_arg(ctx, &args[0]).coerce_to_string() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let a1 = if args.len() >= 2 {
        match eval_scalar_arg(ctx, &args[1]).coerce_to_bool() {
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
    // support A1/R1C1, quoting, and range operators. External workbook references remain
    // unsupported and surface as #REF!.
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

    fn resolve_sheet(ctx: &dyn FunctionContext, sheet: &SheetReference<String>) -> Option<usize> {
        match sheet {
            SheetReference::Current => Some(ctx.current_sheet_id()),
            SheetReference::Sheet(name) => ctx.resolve_sheet_name(name),
            SheetReference::External(_) => None,
        }
    }

    match lowered {
        crate::eval::Expr::CellRef(r) => {
            let Some(sheet_id) = resolve_sheet(ctx, &r.sheet) else {
                return Value::Error(ErrorKind::Ref);
            };
            Value::Reference(crate::functions::Reference {
                sheet_id,
                start: r.addr,
                end: r.addr,
            })
        }
        crate::eval::Expr::RangeRef(r) => {
            let Some(sheet_id) = resolve_sheet(ctx, &r.sheet) else {
                return Value::Error(ErrorKind::Ref);
            };
            Value::Reference(crate::functions::Reference {
                sheet_id,
                start: r.start,
                end: r.end,
            })
        }
        crate::eval::Expr::NameRef(_) => Value::Error(ErrorKind::Ref),
        crate::eval::Expr::Error(e) => Value::Error(e),
        _ => Value::Error(ErrorKind::Ref),
    }
}
