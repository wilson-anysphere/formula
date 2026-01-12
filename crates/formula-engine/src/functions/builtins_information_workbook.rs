use crate::eval::CompiledExpr;
use crate::functions::information::workbook as workbook_info;
use crate::functions::{
    ArgValue, ArraySupport, FunctionContext, FunctionSpec, Reference, SheetId, ThreadSafety,
    ValueType, Volatility,
};
use crate::value::{ErrorKind, Value};

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

fn sheet_number_value(sheet_id: &SheetId) -> Value {
    match workbook_info::sheet_number(sheet_id) {
        Ok(n) => Value::Number(n),
        Err(e) => Value::Error(e),
    }
}

fn sheet_number_value_for_references(references: &[Reference]) -> Value {
    match workbook_info::sheet_number_for_references(references) {
        Ok(n) => Value::Number(n),
        Err(e) => Value::Error(e),
    }
}

fn sheet_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    if args.is_empty() {
        return Value::Number((ctx.current_sheet_id() + 1) as f64);
    }

    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(r) => sheet_number_value(&r.sheet_id),
        ArgValue::ReferenceUnion(ranges) => sheet_number_value_for_references(&ranges),
        ArgValue::Scalar(Value::Reference(r)) => sheet_number_value(&r.sheet_id),
        ArgValue::Scalar(Value::ReferenceUnion(ranges)) => sheet_number_value_for_references(&ranges),
        ArgValue::Scalar(Value::Text(name)) => {
            let name = name.trim();
            if name.is_empty() {
                return Value::Error(ErrorKind::NA);
            }
            match ctx.resolve_sheet_name(name) {
                Some(id) => Value::Number((id + 1) as f64),
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
        ArgValue::Scalar(Value::ReferenceUnion(ranges)) => sheets_count_value_for_references(&ranges),
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

    ctx.record_reference(&reference);

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

    ctx.record_reference(&reference);

    let has_formula = ctx
        .get_cell_formula(&reference.sheet_id, reference.start)
        .is_some();
    Value::Bool(has_formula)
}
