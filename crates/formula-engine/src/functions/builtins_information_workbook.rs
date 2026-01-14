use crate::eval::split_external_sheet_key;
use crate::eval::CompiledExpr;
use crate::functions::information::workbook as workbook_info;
use crate::functions::{
    ArgValue, ArraySupport, FunctionContext, FunctionSpec, Reference, SheetId, ThreadSafety,
    ValueType, Volatility,
};
use crate::value::{ErrorKind, Value};
use formula_model::sheet_name_eq_case_insensitive;
use std::sync::Arc;

inventory::submit! {
    FunctionSpec {
        name: "SHEET",
        min_args: 0,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: sheet_fn,
    }
}

fn external_sheet_index(ctx: &dyn FunctionContext, sheet_key: &str) -> Option<usize> {
    let (workbook, sheet) = split_external_sheet_key(sheet_key)?;
    let order = ctx.workbook_sheet_names(workbook)?;
    order
        .iter()
        .position(|s| sheet_name_eq_case_insensitive(s, sheet))
}

fn sheet_number_value(ctx: &dyn FunctionContext, sheet_id: &SheetId) -> Value {
    match sheet_id {
        SheetId::Local(id) => match ctx.sheet_order_index(*id) {
            Some(idx) => Value::Number((idx + 1) as f64),
            None => Value::Error(ErrorKind::NA),
        },
        SheetId::External(key) => match external_sheet_index(ctx, key) {
            Some(idx) => Value::Number((idx + 1) as f64),
            None => Value::Error(ErrorKind::NA),
        },
    }
}

fn sheet_number_value_for_references(ctx: &dyn FunctionContext, references: &[Reference]) -> Value {
    match workbook_info::sheet_number_for_references(ctx, references) {
        Ok(n) => Value::Number(n),
        Err(_e) => {
            // If there are no local sheets in the reference union, attempt to resolve the sheet
            // order for an external workbook.
            let mut workbook: Option<String> = None;
            let mut order: Option<Arc<[String]>> = None;
            let mut min_idx: Option<usize> = None;

            for r in references {
                let SheetId::External(key) = &r.sheet_id else {
                    return Value::Error(ErrorKind::NA);
                };
                let Some((wb, sheet)) = split_external_sheet_key(key) else {
                    return Value::Error(ErrorKind::NA);
                };

                match &workbook {
                    Some(existing) if existing != wb => return Value::Error(ErrorKind::NA),
                    Some(_) => {}
                    None => {
                        workbook = Some(wb.to_string());
                        order = ctx.workbook_sheet_names(wb);
                        if order.is_none() {
                            return Value::Error(ErrorKind::NA);
                        }
                    }
                }

                let order = order.as_ref().expect("checked is_none above");
                let idx = match order
                    .iter()
                    .position(|s| sheet_name_eq_case_insensitive(s, sheet))
                {
                    Some(idx) => idx,
                    None => return Value::Error(ErrorKind::NA),
                };
                min_idx = Some(match min_idx {
                    Some(existing) => existing.min(idx),
                    None => idx,
                });
            }

            match min_idx {
                Some(idx) => Value::Number((idx + 1) as f64),
                None => Value::Error(ErrorKind::NA),
            }
        }
    }
}

fn sheet_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.is_empty() {
        return match ctx.sheet_order_index(ctx.current_sheet_id()) {
            Some(idx) => Value::Number((idx + 1) as f64),
            None => Value::Error(ErrorKind::NA),
        };
    }

    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(r) => sheet_number_value(ctx, &r.sheet_id),
        ArgValue::ReferenceUnion(ranges) => sheet_number_value_for_references(ctx, &ranges),
        ArgValue::Scalar(Value::Reference(r)) => sheet_number_value(ctx, &r.sheet_id),
        ArgValue::Scalar(Value::ReferenceUnion(ranges)) => {
            sheet_number_value_for_references(ctx, &ranges)
        }
        ArgValue::Scalar(Value::Text(name)) => {
            let name = name.trim();
            if name.is_empty() {
                return Value::Error(ErrorKind::NA);
            }
            match ctx.resolve_sheet_name(name) {
                Some(id) => match ctx.sheet_order_index(id) {
                    Some(idx) => Value::Number((idx + 1) as f64),
                    None => Value::Error(ErrorKind::NA),
                },
                None => Value::Error(ErrorKind::NA),
            }
        }
        ArgValue::Scalar(Value::Error(e)) => Value::Error(e),
        ArgValue::Scalar(_) => Value::Error(ErrorKind::Value),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "SHEETS",
        min_args: 0,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: sheets_fn,
    }
}

fn sheets_count_value_for_references(references: &[Reference]) -> Value {
    Value::Number(workbook_info::count_distinct_sheets(references) as f64)
}

fn sheets_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.is_empty() {
        return Value::Number(ctx.sheet_count() as f64);
    }

    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(_r) => Value::Number(1.0),
        ArgValue::ReferenceUnion(ranges) => sheets_count_value_for_references(&ranges),
        ArgValue::Scalar(Value::Reference(_r)) => Value::Number(1.0),
        ArgValue::Scalar(Value::ReferenceUnion(ranges)) => {
            sheets_count_value_for_references(&ranges)
        }
        ArgValue::Scalar(Value::Error(e)) => Value::Error(e),
        ArgValue::Scalar(_) => Value::Error(ErrorKind::Value),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "FORMULATEXT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Any],
        implementation: formulatext_fn,
    }
}

fn single_cell_reference_from_arg(arg: ArgValue, err: ErrorKind) -> Result<Reference, Value> {
    let mut reference = match arg {
        ArgValue::Reference(r) => r,
        ArgValue::ReferenceUnion(mut ranges) => {
            if ranges.len() == 1 {
                ranges.pop().expect("checked len")
            } else {
                return Err(Value::Error(err));
            }
        }
        ArgValue::Scalar(Value::Reference(r)) => r,
        ArgValue::Scalar(Value::ReferenceUnion(mut ranges)) => {
            if ranges.len() == 1 {
                ranges.pop().expect("checked len")
            } else {
                return Err(Value::Error(err));
            }
        }
        ArgValue::Scalar(Value::Error(e)) => return Err(Value::Error(e)),
        ArgValue::Scalar(_) => return Err(Value::Error(err)),
    };

    reference = reference.normalized();
    if !reference.is_single_cell() {
        return Err(Value::Error(err));
    }
    Ok(reference)
}

fn formulatext_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let reference = match single_cell_reference_from_arg(ctx.eval_arg(&args[0]), ErrorKind::NA) {
        Ok(r) => r,
        Err(v) => return v,
    };

    // FORMULATEXT reads formula metadata (not the computed cell value), so a direct self-reference
    // (e.g. `=FORMULATEXT(A1)` in `A1`) is not a true dependency and should not be recorded for
    // dynamic dependency tracing (avoids spurious self-edges).
    let is_self_reference = matches!(&reference.sheet_id, SheetId::Local(id) if *id == ctx.current_sheet_id())
        && reference.start == ctx.current_cell_addr();
    if !is_self_reference {
        ctx.record_reference(&reference);
    }

    match ctx.get_cell_formula(&reference.sheet_id, reference.start) {
        Some(formula) => Value::Text(workbook_info::normalize_formula_text(formula)),
        None => Value::Error(ErrorKind::NA),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "ISFORMULA",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: isformula_fn,
    }
}

fn isformula_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let reference = match single_cell_reference_from_arg(ctx.eval_arg(&args[0]), ErrorKind::Value) {
        Ok(r) => r,
        Err(v) => return v,
    };

    // ISFORMULA reads formula presence metadata (not the computed value), so a direct self-reference
    // should not be recorded as a dynamic dependency.
    let is_self_reference = matches!(&reference.sheet_id, SheetId::Local(id) if *id == ctx.current_sheet_id())
        && reference.start == ctx.current_cell_addr();
    if !is_self_reference {
        ctx.record_reference(&reference);
    }

    let has_formula = ctx
        .get_cell_formula(&reference.sheet_id, reference.start)
        .is_some();
    Value::Bool(has_formula)
}
