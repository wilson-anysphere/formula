use crate::eval::CompiledExpr;
use crate::functions::information::worksheet;
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

inventory::submit! {
    FunctionSpec {
        name: "INFO",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Text],
        implementation: info_fn,
    }
}

fn info_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let type_text = match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    worksheet::info(ctx, &type_text)
}

inventory::submit! {
    FunctionSpec {
        name: "CELL",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Text, ValueType::Any],
        implementation: cell_fn,
    }
}

fn cell_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let info_type = match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };

    let reference = if args.len() >= 2 {
        match ctx.eval_arg(&args[1]) {
            ArgValue::Reference(r) => Some(r),
            ArgValue::ReferenceUnion(_) => return Value::Error(ErrorKind::Value),
            ArgValue::Scalar(Value::Reference(r)) => Some(r),
            ArgValue::Scalar(Value::ReferenceUnion(_)) => return Value::Error(ErrorKind::Value),
            ArgValue::Scalar(Value::Error(e)) => return Value::Error(e),
            ArgValue::Scalar(_) => return Value::Error(ErrorKind::Value),
        }
    } else {
        None
    };

    worksheet::cell(ctx, &info_type, reference)
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
