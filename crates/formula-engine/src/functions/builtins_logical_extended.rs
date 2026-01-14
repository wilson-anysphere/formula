use std::collections::HashSet;

use crate::eval::CompiledExpr;
use crate::functions::{ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

const VAR_ARGS: usize = 255;

fn is_text_like(value: &Value) -> bool {
    matches!(value, Value::Text(_) | Value::Entity(_) | Value::Record(_))
}

inventory::submit! {
    FunctionSpec {
        name: "XOR",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Any],
        implementation: xor_fn,
    }
}

fn xor_update(ctx: &dyn FunctionContext, acc: &mut bool, value: &Value) -> Result<(), ErrorKind> {
    match value {
        Value::Error(e) => Err(*e),
        Value::Number(n) => {
            *acc ^= *n != 0.0;
            Ok(())
        }
        Value::Bool(b) => {
            *acc ^= *b;
            Ok(())
        }
        Value::Blank => Ok(()),
        // Scalar text arguments accept TRUE/FALSE (and numeric text) coercions like NOT().
        Value::Text(_) => {
            *acc ^= value.coerce_to_bool_with_ctx(ctx)?;
            Ok(())
        }
        Value::Entity(_) | Value::Record(_) => Err(ErrorKind::Value),
        Value::Array(arr) => {
            for v in arr.iter() {
                match v {
                    Value::Error(e) => return Err(*e),
                    Value::Number(n) => *acc ^= *n != 0.0,
                    Value::Bool(b) => *acc ^= *b,
                    Value::Lambda(_) => return Err(ErrorKind::Value),
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
            Ok(())
        }
        Value::Reference(_) | Value::ReferenceUnion(_) | Value::Lambda(_) | Value::Spill { .. } => {
            Err(ErrorKind::Value)
        }
    }
}

fn xor_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut acc = false;

    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => {
                if let Err(e) = xor_update(ctx, &mut acc, &v) {
                    return Value::Error(e);
                }
            }
            ArgValue::Reference(r) => {
                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Value::Error(e),
                        Value::Number(n) => acc ^= n != 0.0,
                        Value::Bool(b) => acc ^= b,
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
                let mut seen = HashSet::new();
                for r in ranges {
                    for addr in ctx.iter_reference_cells(&r) {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        match v {
                            Value::Error(e) => return Value::Error(e),
                            Value::Number(n) => acc ^= n != 0.0,
                            Value::Bool(b) => acc ^= b,
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

    Value::Bool(acc)
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
