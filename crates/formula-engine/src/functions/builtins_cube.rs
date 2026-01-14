use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

const VAR_ARGS: usize = 255;

inventory::submit! {
    FunctionSpec {
        name: "RTD",
        min_args: 3,
        max_args: VAR_ARGS,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::NotThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Text],
        implementation: rtd_fn,
    }
}

fn rtd_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let Some(provider) = ctx.external_data_provider() else {
        return Value::Error(ErrorKind::NA);
    };

    let prog_id = match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let server = match eval_scalar_arg(ctx, &args[1]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let mut topics = Vec::with_capacity(args.len().saturating_sub(2));
    for expr in &args[2..] {
        match eval_scalar_arg(ctx, expr).coerce_to_string_with_ctx(ctx) {
            Ok(v) => topics.push(v),
            Err(e) => return Value::Error(e),
        }
    }

    provider.rtd(&prog_id, &server, &topics)
}

inventory::submit! {
    FunctionSpec {
        name: "CUBEVALUE",
        min_args: 2,
        max_args: VAR_ARGS,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::NotThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Text, ValueType::Text],
        implementation: cubevalue_fn,
    }
}

fn cubevalue_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let Some(provider) = ctx.external_data_provider() else {
        return Value::Error(ErrorKind::NA);
    };

    let connection = match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let mut tuples = Vec::with_capacity(args.len().saturating_sub(1));
    for expr in &args[1..] {
        match eval_scalar_arg(ctx, expr).coerce_to_string_with_ctx(ctx) {
            Ok(v) => tuples.push(v),
            Err(e) => return Value::Error(e),
        }
    }

    provider.cube_value(&connection, &tuples)
}

inventory::submit! {
    FunctionSpec {
        name: "CUBEMEMBER",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::NotThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Text],
        implementation: cubemember_fn,
    }
}

fn cubemember_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let Some(provider) = ctx.external_data_provider() else {
        return Value::Error(ErrorKind::NA);
    };

    let connection = match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let member_expression = match eval_scalar_arg(ctx, &args[1]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let caption = if args.len() >= 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_string_with_ctx(ctx) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    provider.cube_member(&connection, &member_expression, caption.as_deref())
}

inventory::submit! {
    FunctionSpec {
        name: "CUBEMEMBERPROPERTY",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::NotThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Text],
        implementation: cubememberproperty_fn,
    }
}

fn cubememberproperty_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let Some(provider) = ctx.external_data_provider() else {
        return Value::Error(ErrorKind::NA);
    };

    let connection = match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let member_expression_or_handle =
        match eval_scalar_arg(ctx, &args[1]).coerce_to_string_with_ctx(ctx) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
    let property = match eval_scalar_arg(ctx, &args[2]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    provider.cube_member_property(&connection, &member_expression_or_handle, &property)
}

inventory::submit! {
    FunctionSpec {
        name: "CUBERANKEDMEMBER",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::NotThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Number, ValueType::Text],
        implementation: cuberankedmember_fn,
    }
}

fn cuberankedmember_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let Some(provider) = ctx.external_data_provider() else {
        return Value::Error(ErrorKind::NA);
    };

    let connection = match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let set_expression_or_handle =
        match eval_scalar_arg(ctx, &args[1]).coerce_to_string_with_ctx(ctx) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
    let rank = match eval_scalar_arg(ctx, &args[2]).coerce_to_i64_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let caption = if args.len() >= 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_string_with_ctx(ctx) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    provider.cube_ranked_member(
        &connection,
        &set_expression_or_handle,
        rank,
        caption.as_deref(),
    )
}

inventory::submit! {
    FunctionSpec {
        name: "CUBESET",
        min_args: 2,
        max_args: 5,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::NotThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Text, ValueType::Number, ValueType::Text],
        implementation: cubeset_fn,
    }
}

fn cubeset_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let Some(provider) = ctx.external_data_provider() else {
        return Value::Error(ErrorKind::NA);
    };

    let connection = match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let set_expression = match eval_scalar_arg(ctx, &args[1]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let caption = if args.len() >= 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_string_with_ctx(ctx) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    let sort_order = if args.len() >= 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_i64_with_ctx(ctx) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    let sort_by = if args.len() >= 5 {
        match eval_scalar_arg(ctx, &args[4]).coerce_to_string_with_ctx(ctx) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    provider.cube_set(
        &connection,
        &set_expression,
        caption.as_deref(),
        sort_order,
        sort_by.as_deref(),
    )
}

inventory::submit! {
    FunctionSpec {
        name: "CUBESETCOUNT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::NotThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: cubesetcount_fn,
    }
}

fn cubesetcount_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let Some(provider) = ctx.external_data_provider() else {
        return Value::Error(ErrorKind::NA);
    };

    let set_expression_or_handle =
        match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };

    provider.cube_set_count(&set_expression_or_handle)
}

inventory::submit! {
    FunctionSpec {
        name: "CUBEKPIMEMBER",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::NotThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Text, ValueType::Text],
        implementation: cubekpimember_fn,
    }
}

fn cubekpimember_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let Some(provider) = ctx.external_data_provider() else {
        return Value::Error(ErrorKind::NA);
    };

    let connection = match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let kpi_name = match eval_scalar_arg(ctx, &args[1]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let kpi_property = match eval_scalar_arg(ctx, &args[2]).coerce_to_string_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let caption = if args.len() >= 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_string_with_ctx(ctx) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    provider.cube_kpi_member(&connection, &kpi_name, &kpi_property, caption.as_deref())
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
