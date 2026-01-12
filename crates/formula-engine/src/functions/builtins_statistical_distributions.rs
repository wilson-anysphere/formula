use crate::eval::CompiledExpr;
use crate::functions::array_lift;
use crate::functions::{ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::Value;

inventory::submit! {
    FunctionSpec {
        name: "NORM.DIST",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: norm_dist_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "NORMDIST",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: norm_dist_fn,
    }
}

fn norm_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let mean = array_lift::eval_arg(ctx, &args[1]);
    let std_dev = array_lift::eval_arg(ctx, &args[2]);
    let cumulative = array_lift::eval_arg(ctx, &args[3]);
    array_lift::lift4(x, mean, std_dev, cumulative, |x, mean, std_dev, cumulative| {
        let x = x.coerce_to_number()?;
        let mean = mean.coerce_to_number()?;
        let std_dev = std_dev.coerce_to_number()?;
        let cumulative = cumulative.coerce_to_bool()?;
        Ok(Value::Number(crate::functions::statistical::norm_dist(
            x, mean, std_dev, cumulative,
        )?))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "NORM.S.DIST",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Bool],
        implementation: norm_s_dist_fn,
    }
}

fn norm_s_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let z = array_lift::eval_arg(ctx, &args[0]);
    let cumulative = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(z, cumulative, |z, cumulative| {
        let z = z.coerce_to_number()?;
        let cumulative = cumulative.coerce_to_bool()?;
        Ok(Value::Number(crate::functions::statistical::norm_s_dist(
            z, cumulative,
        )?))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "NORMSDIST",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: norms_dist_fn,
    }
}

fn norms_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let z = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(z, |z| {
        let z = z.coerce_to_number()?;
        Ok(Value::Number(crate::functions::statistical::norm_s_dist(z, true)?))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "NORM.INV",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: norm_inv_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "NORMINV",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: norm_inv_fn,
    }
}

fn norm_inv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let probability = array_lift::eval_arg(ctx, &args[0]);
    let mean = array_lift::eval_arg(ctx, &args[1]);
    let std_dev = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(probability, mean, std_dev, |probability, mean, std_dev| {
        let probability = probability.coerce_to_number()?;
        let mean = mean.coerce_to_number()?;
        let std_dev = std_dev.coerce_to_number()?;
        Ok(Value::Number(crate::functions::statistical::norm_inv(
            probability, mean, std_dev,
        )?))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "NORM.S.INV",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: norm_s_inv_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "NORMSINV",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: norm_s_inv_fn,
    }
}

fn norm_s_inv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let probability = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(probability, |probability| {
        let probability = probability.coerce_to_number()?;
        Ok(Value::Number(crate::functions::statistical::norm_s_inv(
            probability,
        )?))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "PHI",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: phi_fn,
    }
}

fn phi_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(x, |x| {
        let x = x.coerce_to_number()?;
        Ok(Value::Number(crate::functions::statistical::phi(x)?))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "GAUSS",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: gauss_fn,
    }
}

fn gauss_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let z = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(z, |z| {
        let z = z.coerce_to_number()?;
        Ok(Value::Number(crate::functions::statistical::gauss(z)?))
    })
}
