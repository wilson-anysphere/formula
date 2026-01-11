use std::collections::HashSet;

use crate::eval::CompiledExpr;
use crate::functions::{ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

const VAR_ARGS: usize = 255;

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

fn xor_update(acc: &mut bool, value: &Value) -> Result<(), ErrorKind> {
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
            *acc ^= value.coerce_to_bool()?;
            Ok(())
        }
        Value::Array(arr) => {
            for v in arr.iter() {
                match v {
                    Value::Error(e) => return Err(*e),
                    Value::Number(n) => *acc ^= *n != 0.0,
                    Value::Bool(b) => *acc ^= *b,
                    // Text and blanks in arrays are ignored (same as references).
                    Value::Text(_)
                    | Value::Blank
                    | Value::Array(_)
                    | Value::Lambda(_)
                    | Value::Spill { .. }
                    | Value::Reference(_)
                    | Value::ReferenceUnion(_) => {}
                }
            }
            Ok(())
        }
        Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => Err(ErrorKind::Value),
    }
}

fn xor_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut acc = false;

    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => {
                if let Err(e) = xor_update(&mut acc, &v) {
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
                        // Text and blanks in references are ignored.
                        Value::Text(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Lambda(_)
                        | Value::Spill { .. }
                        | Value::Reference(_)
                        | Value::ReferenceUnion(_) => {}
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
                            Value::Text(_)
                            | Value::Blank
                            | Value::Array(_)
                            | Value::Lambda(_)
                            | Value::Spill { .. }
                            | Value::Reference(_)
                            | Value::ReferenceUnion(_) => {}
                        }
                    }
                }
            }
        }
    }

    Value::Bool(acc)
}
