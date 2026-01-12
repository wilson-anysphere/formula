use crate::functions::{ArraySupport, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};

inventory::submit! {
    FunctionSpec {
        name: "MDETERM",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: crate::functions::math::matrix::mdeterm,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "MINVERSE",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any],
        implementation: crate::functions::math::matrix::minverse,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "MMULT",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: crate::functions::math::matrix::mmult,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "MUNIT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Number],
        implementation: crate::functions::math::matrix::munit,
    }
}

