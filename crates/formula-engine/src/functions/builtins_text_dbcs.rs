use crate::functions::text::dbcs;
use crate::functions::{ArraySupport, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};

inventory::submit! {
    FunctionSpec {
        name: "FINDB",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Number],
        implementation: dbcs::findb_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "SEARCHB",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Number],
        implementation: dbcs::searchb_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "REPLACEB",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number, ValueType::Number, ValueType::Text],
        implementation: dbcs::replaceb_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "LEFTB",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: dbcs::leftb_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "RIGHTB",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: dbcs::rightb_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "MIDB",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number, ValueType::Number],
        implementation: dbcs::midb_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "LENB",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: dbcs::lenb_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "ASC",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: dbcs::asc_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DBCS",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: dbcs::dbcs_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "PHONETIC",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: dbcs::phonetic_fn,
    }
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
