use crate::eval::CompiledExpr;
use crate::functions::database::DatabaseQuery;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

const DB_ARGS: usize = 3;

inventory::submit! {
    FunctionSpec {
        name: "DAVERAGE",
        min_args: DB_ARGS,
        max_args: DB_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: daverage_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DCOUNT",
        min_args: DB_ARGS,
        max_args: DB_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: dcount_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DCOUNTA",
        min_args: DB_ARGS,
        max_args: DB_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: dcounta_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DGET",
        min_args: DB_ARGS,
        max_args: DB_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: dget_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DMAX",
        min_args: DB_ARGS,
        max_args: DB_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: dmax_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DMIN",
        min_args: DB_ARGS,
        max_args: DB_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: dmin_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DPRODUCT",
        min_args: DB_ARGS,
        max_args: DB_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: dproduct_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DSUM",
        min_args: DB_ARGS,
        max_args: DB_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: dsum_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DSTDEV",
        min_args: DB_ARGS,
        max_args: DB_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: dstdev_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DSTDEVP",
        min_args: DB_ARGS,
        max_args: DB_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: dstdevp_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DVAR",
        min_args: DB_ARGS,
        max_args: DB_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: dvar_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DVARP",
        min_args: DB_ARGS,
        max_args: DB_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any],
        implementation: dvarp_fn,
    }
}

fn parse_query(
    ctx: &dyn FunctionContext,
    args: &[CompiledExpr],
) -> Result<DatabaseQuery, ErrorKind> {
    let database = ctx.eval_arg(&args[0]);
    let field = eval_scalar_arg(ctx, &args[1]);
    let criteria = ctx.eval_arg(&args[2]);
    crate::functions::database::parse_query(ctx, database, field, criteria)
}

fn dsum_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let query = match parse_query(ctx, args) {
        Ok(q) => q,
        Err(e) => return Value::Error(e),
    };

    let mut sum = 0.0;
    for row in query.iter_matching_rows(ctx) {
        let row = match row {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        match query.field_value(row) {
            Value::Number(n) => sum += *n,
            Value::Error(e) => return Value::Error(*e),
            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            _ => {}
        }
    }
    Value::Number(sum)
}

fn daverage_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let query = match parse_query(ctx, args) {
        Ok(q) => q,
        Err(e) => return Value::Error(e),
    };

    let mut sum = 0.0;
    let mut count: u64 = 0;
    for row in query.iter_matching_rows(ctx) {
        let row = match row {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        match query.field_value(row) {
            Value::Number(n) => {
                sum += *n;
                count += 1;
            }
            Value::Error(e) => return Value::Error(*e),
            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            _ => {}
        }
    }

    if count == 0 {
        return Value::Error(ErrorKind::Div0);
    }
    Value::Number(sum / count as f64)
}

fn dmin_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let query = match parse_query(ctx, args) {
        Ok(q) => q,
        Err(e) => return Value::Error(e),
    };

    let mut best: Option<f64> = None;
    for row in query.iter_matching_rows(ctx) {
        let row = match row {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        match query.field_value(row) {
            Value::Number(n) => best = Some(best.map_or(*n, |b| b.min(*n))),
            Value::Error(e) => return Value::Error(*e),
            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            _ => {}
        }
    }
    Value::Number(best.unwrap_or(0.0))
}

fn dmax_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let query = match parse_query(ctx, args) {
        Ok(q) => q,
        Err(e) => return Value::Error(e),
    };

    let mut best: Option<f64> = None;
    for row in query.iter_matching_rows(ctx) {
        let row = match row {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        match query.field_value(row) {
            Value::Number(n) => best = Some(best.map_or(*n, |b| b.max(*n))),
            Value::Error(e) => return Value::Error(*e),
            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            _ => {}
        }
    }
    Value::Number(best.unwrap_or(0.0))
}

fn dproduct_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let query = match parse_query(ctx, args) {
        Ok(q) => q,
        Err(e) => return Value::Error(e),
    };

    let mut out = 1.0;
    let mut saw_number = false;
    for row in query.iter_matching_rows(ctx) {
        let row = match row {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        match query.field_value(row) {
            Value::Number(n) => {
                saw_number = true;
                out *= *n;
            }
            Value::Error(e) => return Value::Error(*e),
            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            _ => {}
        }
    }

    if !saw_number {
        return Value::Number(1.0);
    }
    if out.is_finite() {
        Value::Number(out)
    } else {
        Value::Error(ErrorKind::Num)
    }
}

fn dcount_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let query = match parse_query(ctx, args) {
        Ok(q) => q,
        Err(e) => return Value::Error(e),
    };

    let mut count: u64 = 0;
    for row in query.iter_matching_rows(ctx) {
        let row = match row {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        if matches!(query.field_value(row), Value::Number(_)) {
            count += 1;
        }
    }
    Value::Number(count as f64)
}

fn dcounta_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let query = match parse_query(ctx, args) {
        Ok(q) => q,
        Err(e) => return Value::Error(e),
    };

    let mut count: u64 = 0;
    for row in query.iter_matching_rows(ctx) {
        let row = match row {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        if !matches!(query.field_value(row), Value::Blank) {
            count += 1;
        }
    }
    Value::Number(count as f64)
}

fn dget_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let query = match parse_query(ctx, args) {
        Ok(q) => q,
        Err(e) => return Value::Error(e),
    };

    let mut match_row: Option<usize> = None;
    for row in query.iter_matching_rows(ctx) {
        let row = match row {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        if match_row.is_some() {
            return Value::Error(ErrorKind::Num);
        }
        match_row = Some(row);
    }

    let Some(row) = match_row else {
        return Value::Error(ErrorKind::Value);
    };

    match query.field_value(row) {
        Value::Lambda(_)
        | Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Array(_)
        | Value::Spill { .. } => Value::Error(ErrorKind::Value),
        other => other.clone(),
    }
}

fn dstdev_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    dstdev_impl(ctx, args, StdevVariant::Sample)
}

fn dstdevp_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    dstdev_impl(ctx, args, StdevVariant::Population)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StdevVariant {
    Sample,
    Population,
}

fn dstdev_impl(ctx: &dyn FunctionContext, args: &[CompiledExpr], variant: StdevVariant) -> Value {
    let query = match parse_query(ctx, args) {
        Ok(q) => q,
        Err(e) => return Value::Error(e),
    };

    let mut values = Vec::new();
    for row in query.iter_matching_rows(ctx) {
        let row = match row {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        match query.field_value(row) {
            Value::Number(n) => values.push(*n),
            Value::Error(e) => return Value::Error(*e),
            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            _ => {}
        }
    }

    let out = match variant {
        StdevVariant::Sample => crate::functions::statistical::stdev_s(&values),
        StdevVariant::Population => crate::functions::statistical::stdev_p(&values),
    };
    match out {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}

fn dvar_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    dvar_impl(ctx, args, VarVariant::Sample)
}

fn dvarp_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    dvar_impl(ctx, args, VarVariant::Population)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum VarVariant {
    Sample,
    Population,
}

fn dvar_impl(ctx: &dyn FunctionContext, args: &[CompiledExpr], variant: VarVariant) -> Value {
    let query = match parse_query(ctx, args) {
        Ok(q) => q,
        Err(e) => return Value::Error(e),
    };

    let mut values = Vec::new();
    for row in query.iter_matching_rows(ctx) {
        let row = match row {
            Ok(r) => r,
            Err(e) => return Value::Error(e),
        };
        match query.field_value(row) {
            Value::Number(n) => values.push(*n),
            Value::Error(e) => return Value::Error(*e),
            Value::Lambda(_) => return Value::Error(ErrorKind::Value),
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            _ => {}
        }
    }

    let out = match variant {
        VarVariant::Sample => crate::functions::statistical::var_s(&values),
        VarVariant::Population => crate::functions::statistical::var_p(&values),
    };
    match out {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}
