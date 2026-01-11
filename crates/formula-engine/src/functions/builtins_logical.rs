use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

const VAR_ARGS: usize = 255;

inventory::submit! {
    FunctionSpec {
        name: "IF",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: if_fn,
    }
}

fn if_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let cond_val = eval_scalar_arg(ctx, &args[0]);
    if let Value::Error(e) = cond_val {
        return Value::Error(e);
    }

    let cond = match cond_val.coerce_to_bool() {
        Ok(b) => b,
        Err(e) => return Value::Error(e),
    };

    if cond {
        ctx.eval_scalar(&args[1])
    } else if args.len() >= 3 {
        ctx.eval_scalar(&args[2])
    } else {
        Value::Bool(false)
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
        match ctx.eval_arg(arg) {
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
                Value::Text(_) => return Value::Error(ErrorKind::Value),
                Value::Array(_) | Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            },
            ArgValue::Reference(r) => {
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(r.sheet_id, addr);
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
                        // Text and blanks in references are ignored.
                        Value::Text(_) | Value::Blank | Value::Array(_) | Value::Spill { .. } => {}
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
        match ctx.eval_arg(arg) {
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
                Value::Text(_) => return Value::Error(ErrorKind::Value),
                Value::Array(_) | Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            },
            ArgValue::Reference(r) => {
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(r.sheet_id, addr);
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
                        Value::Text(_) | Value::Blank | Value::Array(_) | Value::Spill { .. } => {}
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
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: not_fn,
    }
}

fn not_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let v = eval_scalar_arg(ctx, &args[0]);
    let b = match v.coerce_to_bool() {
        Ok(b) => b,
        Err(e) => return Value::Error(e),
    };
    Value::Bool(!b)
}

inventory::submit! {
    FunctionSpec {
        name: "IFERROR",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: iferror_fn,
    }
}

fn iferror_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let first = ctx.eval_scalar(&args[0]);
    match first {
        Value::Error(_) => ctx.eval_scalar(&args[1]),
        other => other,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "IFNA",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: ifna_fn,
    }
}

fn ifna_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let first = ctx.eval_scalar(&args[0]);
    match first {
        Value::Error(ErrorKind::NA) => ctx.eval_scalar(&args[1]),
        other => other,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "ISERROR",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: iserror_fn,
    }
}

fn iserror_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let v = ctx.eval_scalar(&args[0]);
    Value::Bool(matches!(v, Value::Error(_)))
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
