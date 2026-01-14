use crate::eval::{CompiledExpr, SheetReference};
use crate::functions::array_lift;
use crate::functions::{ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

const VAR_ARGS: usize = 255;

fn is_text_like(value: &Value) -> bool {
    matches!(value, Value::Text(_) | Value::Entity(_) | Value::Record(_))
}

fn is_scalar_cell_ref(expr: &CompiledExpr) -> bool {
    match expr {
        // Excel treats single-cell references passed to logical aggregations (AND/OR) like scalar
        // values, which differs from range semantics (e.g. `A1:A1`).
        //
        // Avoid changing semantics for 3D references (`Sheet1:Sheet3!A1`), which behave like a
        // multi-cell reference.
        CompiledExpr::CellRef(r) => !matches!(r.sheet, SheetReference::SheetRange(_, _)),
        _ => false,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IF",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: if_fn,
    }
}

fn if_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let cond_val = array_lift::eval_arg(ctx, &args[0]);
    match cond_val {
        Value::Array(ref arr) => {
            let mut needs_true = false;
            let mut needs_false = false;
            for el in arr.iter() {
                match el.coerce_to_bool_with_ctx(ctx) {
                    Ok(true) => needs_true = true,
                    Ok(false) => needs_false = true,
                    Err(_) => {}
                }
                if needs_true && needs_false {
                    break;
                }
            }

            if needs_true && needs_false {
                let true_val = array_lift::eval_arg(ctx, &args[1]);
                let false_val = if args.len() >= 3 {
                    array_lift::eval_arg(ctx, &args[2])
                } else {
                    Value::Bool(false)
                };
                return array_lift::lift3(cond_val, true_val, false_val, |cond, t, f| {
                    if cond.coerce_to_bool_with_ctx(ctx)? {
                        Ok(t.clone())
                    } else {
                        Ok(f.clone())
                    }
                });
            }

            if needs_true {
                let true_val = array_lift::eval_arg(ctx, &args[1]);
                return array_lift::lift2(cond_val, true_val, |cond, t| {
                    if cond.coerce_to_bool_with_ctx(ctx)? {
                        Ok(t.clone())
                    } else {
                        // `needs_false` is false, so this branch is unreachable unless the
                        // condition array contains values that coerce differently on a second pass.
                        Ok(Value::Bool(false))
                    }
                });
            }

            if needs_false {
                let false_val = if args.len() >= 3 {
                    array_lift::eval_arg(ctx, &args[2])
                } else {
                    Value::Bool(false)
                };
                return array_lift::lift2(cond_val, false_val, |cond, f| {
                    if cond.coerce_to_bool_with_ctx(ctx)? {
                        // `needs_true` is false, so this branch is unreachable unless the
                        // condition array contains values that coerce differently on a second pass.
                        Ok(Value::Bool(false))
                    } else {
                        Ok(f.clone())
                    }
                });
            }

            // The condition array contains only errors / invalid values, so map the coercion
            // error for each element without forcing evaluation of either branch.
            array_lift::lift1(cond_val, |cond| {
                let _ = cond.coerce_to_bool_with_ctx(ctx)?;
                Ok(Value::Bool(false))
            })
        }
        other => {
            if let Value::Error(e) = other {
                return Value::Error(e);
            }

            let cond = match other.coerce_to_bool_with_ctx(ctx) {
                Ok(b) => b,
                Err(e) => return Value::Error(e),
            };

            if cond {
                array_lift::eval_arg(ctx, &args[1])
            } else if args.len() >= 3 {
                array_lift::eval_arg(ctx, &args[2])
            } else {
                Value::Bool(false)
            }
        }
    }
}

inventory::submit! {
    FunctionSpec {
        name: "AND",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: and_fn,
    }
}

fn and_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut all_true = true;
    let mut any = false;

    for arg in args {
        let arg_value = if is_scalar_cell_ref(arg) {
            ArgValue::Scalar(ctx.eval_scalar(arg))
        } else {
            ctx.eval_arg(arg)
        };

        match arg_value {
            ArgValue::Scalar(v) => match v {
                Value::Error(e) => return Value::Error(e),
                Value::Number(n) => {
                    any = true;
                    if n == 0.0 {
                        all_true = false;
                    }
                }
                Value::Bool(b) => {
                    any = true;
                    if !b {
                        all_true = false;
                    }
                }
                Value::Blank => {}
                Value::Text(_) | Value::Entity(_) | Value::Record(_) => {
                    return Value::Error(ErrorKind::Value)
                }
                Value::Reference(_) | Value::ReferenceUnion(_) => {
                    return Value::Error(ErrorKind::Value)
                }
                Value::Array(arr) => {
                    for v in arr.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => {
                                any = true;
                                if *n == 0.0 {
                                    all_true = false;
                                }
                            }
                            Value::Bool(b) => {
                                any = true;
                                if !*b {
                                    all_true = false;
                                }
                            }
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                            // Text and blanks in arrays are ignored (same as references).
                            other => {
                                if is_text_like(other)
                                    || matches!(
                                        other,
                                        Value::Blank
                                            | Value::Array(_)
                                            | Value::Spill { .. }
                                            | Value::Reference(_)
                                            | Value::ReferenceUnion(_)
                                    )
                                {
                                    continue;
                                }
                            }
                        }
                    }
                }
                Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            },
            ArgValue::Reference(r) => {
                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => {
                            any = true;
                            if n == 0.0 {
                                all_true = false;
                            }
                        }
                        Value::Bool(b) => {
                            any = true;
                            if !b {
                                all_true = false;
                            }
                        }
                        Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                        // Text and blanks in references are ignored.
                        other => {
                            if is_text_like(&other)
                                || matches!(
                                    other,
                                    Value::Blank
                                        | Value::Array(_)
                                        | Value::Spill { .. }
                                        | Value::Reference(_)
                                        | Value::ReferenceUnion(_)
                                )
                            {
                                continue;
                            }
                        }
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                let mut seen = std::collections::HashSet::new();
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => {
                                any = true;
                                if n == 0.0 {
                                    all_true = false;
                                }
                            }
                            Value::Bool(b) => {
                                any = true;
                                if !b {
                                    all_true = false;
                                }
                            }
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                            // Text and blanks in references are ignored.
                            other => {
                                if is_text_like(&other)
                                    || matches!(
                                        other,
                                        Value::Blank
                                            | Value::Array(_)
                                            | Value::Spill { .. }
                                            | Value::Reference(_)
                                            | Value::ReferenceUnion(_)
                                    )
                                {
                                    continue;
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    if !any {
        return Value::Bool(true);
    }
    Value::Bool(all_true)
}

inventory::submit! {
    FunctionSpec {
        name: "OR",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: or_fn,
    }
}

fn or_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut any_true = false;
    let mut any = false;

    for arg in args {
        let arg_value = if is_scalar_cell_ref(arg) {
            ArgValue::Scalar(ctx.eval_scalar(arg))
        } else {
            ctx.eval_arg(arg)
        };

        match arg_value {
            ArgValue::Scalar(v) => match v {
                Value::Error(e) => return Value::Error(e),
                Value::Number(n) => {
                    any = true;
                    if n != 0.0 {
                        any_true = true;
                    }
                }
                Value::Bool(b) => {
                    any = true;
                    if b {
                        any_true = true;
                    }
                }
                Value::Blank => {}
                Value::Text(_) | Value::Entity(_) | Value::Record(_) => {
                    return Value::Error(ErrorKind::Value)
                }
                Value::Reference(_) | Value::ReferenceUnion(_) => {
                    return Value::Error(ErrorKind::Value)
                }
                Value::Array(arr) => {
                    for v in arr.iter() {
                        match v {
                            Value::Error(e) => return Value::Error(*e),
                            Value::Number(n) => {
                                any = true;
                                if *n != 0.0 {
                                    any_true = true;
                                }
                            }
                            Value::Bool(b) => {
                                any = true;
                                if *b {
                                    any_true = true;
                                }
                            }
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                            other => {
                                if is_text_like(other)
                                    || matches!(
                                        other,
                                        Value::Blank
                                            | Value::Array(_)
                                            | Value::Spill { .. }
                                            | Value::Reference(_)
                                            | Value::ReferenceUnion(_)
                                    )
                                {
                                    continue;
                                }
                            }
                        }
                    }
                }
                Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            },
            ArgValue::Reference(r) => {
                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => {
                            any = true;
                            if n != 0.0 {
                                any_true = true;
                            }
                        }
                        Value::Bool(b) => {
                            any = true;
                            if b {
                                any_true = true;
                            }
                        }
                        Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                        other => {
                            if is_text_like(&other)
                                || matches!(
                                    other,
                                    Value::Blank
                                        | Value::Array(_)
                                        | Value::Spill { .. }
                                        | Value::Reference(_)
                                        | Value::ReferenceUnion(_)
                                )
                            {
                                continue;
                            }
                        }
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                let mut seen = std::collections::HashSet::new();
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => {
                                any = true;
                                if n != 0.0 {
                                    any_true = true;
                                }
                            }
                            Value::Bool(b) => {
                                any = true;
                                if b {
                                    any_true = true;
                                }
                            }
                            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
                            other => {
                                if is_text_like(&other)
                                    || matches!(
                                        other,
                                        Value::Blank
                                            | Value::Array(_)
                                            | Value::Spill { .. }
                                            | Value::Reference(_)
                                            | Value::ReferenceUnion(_)
                                    )
                                {
                                    continue;
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    if !any {
        return Value::Bool(false);
    }
    Value::Bool(any_true)
}

inventory::submit! {
    FunctionSpec {
        name: "NOT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: not_fn,
    }
}

fn not_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(value, |v| Ok(Value::Bool(!v.coerce_to_bool_with_ctx(ctx)?)))
}

inventory::submit! {
    FunctionSpec {
        name: "IFERROR",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: iferror_fn,
    }
}

fn iferror_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let first = array_lift::eval_arg(ctx, &args[0]);
    let needs_fallback = match &first {
        Value::Error(_) => true,
        Value::Array(arr) => arr.iter().any(|v| matches!(v, Value::Error(_))),
        _ => false,
    };
    if !needs_fallback {
        return first;
    }

    let fallback = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(first, fallback, |first, fallback| {
        if matches!(first, Value::Error(_)) {
            Ok(fallback.clone())
        } else {
            Ok(first.clone())
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "IFNA",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: ifna_fn,
    }
}

fn ifna_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let first = array_lift::eval_arg(ctx, &args[0]);
    let needs_fallback = match &first {
        Value::Error(ErrorKind::NA) => true,
        Value::Array(arr) => arr.iter().any(|v| matches!(v, Value::Error(ErrorKind::NA))),
        _ => false,
    };
    if !needs_fallback {
        return first;
    }

    let fallback = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(first, fallback, |first, fallback| {
        if matches!(first, Value::Error(ErrorKind::NA)) {
            Ok(fallback.clone())
        } else {
            Ok(first.clone())
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "ISERROR",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: iserror_fn,
    }
}

fn iserror_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let v = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(v, |v| Ok(Value::Bool(matches!(v, Value::Error(_)))))
}

inventory::submit! {
    FunctionSpec {
        name: "NA",
        min_args: 0,
        max_args: 0,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[],
        implementation: na_fn,
    }
}

fn na_fn(_ctx: &dyn FunctionContext, _args: &[CompiledExpr]) -> Value {
    Value::Error(ErrorKind::NA)
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
