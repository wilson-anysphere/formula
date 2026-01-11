use crate::eval::CompiledExpr;
use crate::functions::{ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

inventory::submit! {
    FunctionSpec {
        name: "TRANSPOSE",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: transpose_fn,
    }
}

fn transpose_fn(_ctx: &dyn FunctionContext, _args: &[CompiledExpr]) -> Value {
    // Dynamic array spilling is not implemented yet. For now, treat array-returning
    // functions as a spill attempt rather than an unknown function (#NAME?).
    Value::Error(ErrorKind::Spill)
}

inventory::submit! {
    FunctionSpec {
        name: "SEQUENCE",
        min_args: 1,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: sequence_fn,
    }
}

fn sequence_fn(_ctx: &dyn FunctionContext, _args: &[CompiledExpr]) -> Value {
    // SEQUENCE is a dynamic array function. Until we support arrays/spilling,
    // return a spill error.
    Value::Error(ErrorKind::Spill)
}

