use crate::eval::CompiledExpr;
use crate::functions::array_lift;
use crate::functions::{ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

// ----------------------------------------------------------------------
// Normal distribution
// ----------------------------------------------------------------------

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
    array_lift::lift4(
        x,
        mean,
        std_dev,
        cumulative,
        |x, mean, std_dev, cumulative| {
            let x = x.coerce_to_number_with_ctx(ctx)?;
            let mean = mean.coerce_to_number_with_ctx(ctx)?;
            let std_dev = std_dev.coerce_to_number_with_ctx(ctx)?;
            let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
            Ok(Value::Number(crate::functions::statistical::norm_dist(
                x, mean, std_dev, cumulative,
            )?))
        },
    )
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
        let z = z.coerce_to_number_with_ctx(ctx)?;
        let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
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
        let z = z.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(crate::functions::statistical::norm_s_dist(
            z, true,
        )?))
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
        let probability = probability.coerce_to_number_with_ctx(ctx)?;
        let mean = mean.coerce_to_number_with_ctx(ctx)?;
        let std_dev = std_dev.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(crate::functions::statistical::norm_inv(
            probability,
            mean,
            std_dev,
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
        let probability = probability.coerce_to_number_with_ctx(ctx)?;
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
        let x = x.coerce_to_number_with_ctx(ctx)?;
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
        let z = z.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(crate::functions::statistical::gauss(z)?))
    })
}

// ----------------------------------------------------------------------
// T distribution
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "T.DIST",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: t_dist_fn,
    }
}

fn t_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let df = array_lift::eval_arg(ctx, &args[1]);
    let cumulative = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(x, df, cumulative, |x, df, cumulative| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let df = df.coerce_to_number_with_ctx(ctx)?;
        let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::t_dist(x, df, cumulative)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "T.DIST.RT",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: t_dist_rt_fn,
    }
}

fn t_dist_rt_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let df = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(x, df, |x, df| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let df = df.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::t_dist_rt(x, df)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "T.DIST.2T",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: t_dist_2t_fn,
    }
}

fn t_dist_2t_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let df = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(x, df, |x, df| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let df = df.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::t_dist_2t(x, df)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "T.INV",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: t_inv_fn,
    }
}

fn t_inv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let p = array_lift::eval_arg(ctx, &args[0]);
    let df = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(p, df, |p, df| {
        let p = p.coerce_to_number_with_ctx(ctx)?;
        let df = df.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::t_inv(p, df)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "T.INV.2T",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: t_inv_2t_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "TINV",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: t_inv_2t_fn,
    }
}

fn t_inv_2t_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let p = array_lift::eval_arg(ctx, &args[0]);
    let df = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(p, df, |p, df| {
        let p = p.coerce_to_number_with_ctx(ctx)?;
        let df = df.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::t_inv_2t(p, df)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "TDIST",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: tdist_fn,
    }
}

fn tdist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let df = array_lift::eval_arg(ctx, &args[1]);
    let tails = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(x, df, tails, |x, df, tails| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let df = df.coerce_to_number_with_ctx(ctx)?;
        let tails = tails.coerce_to_i64_with_ctx(ctx)?;
        let out = match tails {
            1 => crate::functions::statistical::distributions::t_dist_rt(x, df)?,
            2 => crate::functions::statistical::distributions::t_dist_2t(x, df)?,
            _ => return Err(ErrorKind::Num),
        };
        Ok(Value::Number(out))
    })
}

// ----------------------------------------------------------------------
// Chi-square distribution
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "CHISQ.DIST",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: chisq_dist_fn,
    }
}

fn chisq_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let df = array_lift::eval_arg(ctx, &args[1]);
    let cumulative = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(x, df, cumulative, |x, df, cumulative| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let df = df.coerce_to_number_with_ctx(ctx)?;
        let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::chisq_dist(x, df, cumulative)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "CHISQ.DIST.RT",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: chisq_dist_rt_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "CHIDIST",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: chisq_dist_rt_fn,
    }
}

fn chisq_dist_rt_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let df = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(x, df, |x, df| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let df = df.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::chisq_dist_rt(x, df)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "CHISQ.INV",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: chisq_inv_fn,
    }
}

fn chisq_inv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let p = array_lift::eval_arg(ctx, &args[0]);
    let df = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(p, df, |p, df| {
        let p = p.coerce_to_number_with_ctx(ctx)?;
        let df = df.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::chisq_inv(p, df)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "CHISQ.INV.RT",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: chisq_inv_rt_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "CHIINV",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: chisq_inv_rt_fn,
    }
}

fn chisq_inv_rt_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let p = array_lift::eval_arg(ctx, &args[0]);
    let df = array_lift::eval_arg(ctx, &args[1]);
    array_lift::lift2(p, df, |p, df| {
        let p = p.coerce_to_number_with_ctx(ctx)?;
        let df = df.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::chisq_inv_rt(p, df)?,
        ))
    })
}

// ----------------------------------------------------------------------
// F distribution
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "F.DIST",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: f_dist_fn,
    }
}

fn f_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let df1 = array_lift::eval_arg(ctx, &args[1]);
    let df2 = array_lift::eval_arg(ctx, &args[2]);
    let cumulative = array_lift::eval_arg(ctx, &args[3]);
    array_lift::lift4(x, df1, df2, cumulative, |x, df1, df2, cumulative| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let df1 = df1.coerce_to_number_with_ctx(ctx)?;
        let df2 = df2.coerce_to_number_with_ctx(ctx)?;
        let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::f_dist(x, df1, df2, cumulative)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "F.DIST.RT",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: f_dist_rt_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "FDIST",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: f_dist_rt_fn,
    }
}

fn f_dist_rt_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let df1 = array_lift::eval_arg(ctx, &args[1]);
    let df2 = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(x, df1, df2, |x, df1, df2| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let df1 = df1.coerce_to_number_with_ctx(ctx)?;
        let df2 = df2.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::f_dist_rt(x, df1, df2)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "F.INV",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: f_inv_fn,
    }
}

fn f_inv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let p = array_lift::eval_arg(ctx, &args[0]);
    let df1 = array_lift::eval_arg(ctx, &args[1]);
    let df2 = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(p, df1, df2, |p, df1, df2| {
        let p = p.coerce_to_number_with_ctx(ctx)?;
        let df1 = df1.coerce_to_number_with_ctx(ctx)?;
        let df2 = df2.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::f_inv(p, df1, df2)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "F.INV.RT",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: f_inv_rt_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "FINV",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: f_inv_rt_fn,
    }
}

fn f_inv_rt_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let p = array_lift::eval_arg(ctx, &args[0]);
    let df1 = array_lift::eval_arg(ctx, &args[1]);
    let df2 = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(p, df1, df2, |p, df1, df2| {
        let p = p.coerce_to_number_with_ctx(ctx)?;
        let df1 = df1.coerce_to_number_with_ctx(ctx)?;
        let df2 = df2.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::f_inv_rt(p, df1, df2)?,
        ))
    })
}

// ----------------------------------------------------------------------
// Beta distribution
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "BETA.DIST",
        min_args: 4,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Bool,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: beta_dist_fn,
    }
}

fn beta_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let alpha = array_lift::eval_arg(ctx, &args[1]);
    let beta = array_lift::eval_arg(ctx, &args[2]);
    let cumulative = array_lift::eval_arg(ctx, &args[3]);
    let a = if args.len() >= 5 {
        array_lift::eval_arg(ctx, &args[4])
    } else {
        Value::Number(0.0)
    };
    let b = if args.len() >= 6 {
        array_lift::eval_arg(ctx, &args[5])
    } else {
        Value::Number(1.0)
    };
    array_lift::lift6(
        x,
        alpha,
        beta,
        cumulative,
        a,
        b,
        |x, alpha, beta, cumulative, a, b| {
            let x = x.coerce_to_number_with_ctx(ctx)?;
            let alpha = alpha.coerce_to_number_with_ctx(ctx)?;
            let beta = beta.coerce_to_number_with_ctx(ctx)?;
            let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
            let a = a.coerce_to_number_with_ctx(ctx)?;
            let b = b.coerce_to_number_with_ctx(ctx)?;
            Ok(Value::Number(
                crate::functions::statistical::distributions::beta_dist(
                    x, alpha, beta, cumulative, a, b,
                )?,
            ))
        },
    )
}

inventory::submit! {
    FunctionSpec {
        name: "BETA.INV",
        min_args: 3,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: beta_inv_fn,
    }
}

fn beta_inv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let p = array_lift::eval_arg(ctx, &args[0]);
    let alpha = array_lift::eval_arg(ctx, &args[1]);
    let beta = array_lift::eval_arg(ctx, &args[2]);
    let a = if args.len() >= 4 {
        array_lift::eval_arg(ctx, &args[3])
    } else {
        Value::Number(0.0)
    };
    let b = if args.len() >= 5 {
        array_lift::eval_arg(ctx, &args[4])
    } else {
        Value::Number(1.0)
    };
    array_lift::lift5(p, alpha, beta, a, b, |p, alpha, beta, a, b| {
        let p = p.coerce_to_number_with_ctx(ctx)?;
        let alpha = alpha.coerce_to_number_with_ctx(ctx)?;
        let beta = beta.coerce_to_number_with_ctx(ctx)?;
        let a = a.coerce_to_number_with_ctx(ctx)?;
        let b = b.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::beta_inv(p, alpha, beta, a, b)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "BETADIST",
        min_args: 3,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: betadist_fn,
    }
}

fn betadist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let alpha = array_lift::eval_arg(ctx, &args[1]);
    let beta = array_lift::eval_arg(ctx, &args[2]);
    let a = if args.len() >= 4 {
        array_lift::eval_arg(ctx, &args[3])
    } else {
        Value::Number(0.0)
    };
    let b = if args.len() >= 5 {
        array_lift::eval_arg(ctx, &args[4])
    } else {
        Value::Number(1.0)
    };
    array_lift::lift5(x, alpha, beta, a, b, |x, alpha, beta, a, b| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let alpha = alpha.coerce_to_number_with_ctx(ctx)?;
        let beta = beta.coerce_to_number_with_ctx(ctx)?;
        let a = a.coerce_to_number_with_ctx(ctx)?;
        let b = b.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::beta_dist(x, alpha, beta, true, a, b)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "BETAINV",
        min_args: 3,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: beta_inv_fn,
    }
}

// ----------------------------------------------------------------------
// Gamma distribution + special functions
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "GAMMA.DIST",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: gamma_dist_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "GAMMADIST",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: gamma_dist_fn,
    }
}

fn gamma_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let alpha = array_lift::eval_arg(ctx, &args[1]);
    let beta = array_lift::eval_arg(ctx, &args[2]);
    let cumulative = array_lift::eval_arg(ctx, &args[3]);
    array_lift::lift4(x, alpha, beta, cumulative, |x, alpha, beta, cumulative| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let alpha = alpha.coerce_to_number_with_ctx(ctx)?;
        let beta = beta.coerce_to_number_with_ctx(ctx)?;
        let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::gamma_dist(x, alpha, beta, cumulative)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "GAMMA.INV",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: gamma_inv_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "GAMMAINV",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: gamma_inv_fn,
    }
}

fn gamma_inv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let p = array_lift::eval_arg(ctx, &args[0]);
    let alpha = array_lift::eval_arg(ctx, &args[1]);
    let beta = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(p, alpha, beta, |p, alpha, beta| {
        let p = p.coerce_to_number_with_ctx(ctx)?;
        let alpha = alpha.coerce_to_number_with_ctx(ctx)?;
        let beta = beta.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::gamma_inv(p, alpha, beta)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "GAMMA",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: gamma_fn,
    }
}

fn gamma_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(number, |number| {
        let number = number.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::gamma_fn(number)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "GAMMALN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: gammaln_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "GAMMALN.PRECISE",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: gammaln_fn,
    }
}

fn gammaln_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(number, |number| {
        let number = number.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::gammaln(number)?,
        ))
    })
}

// ----------------------------------------------------------------------
// Lognormal distribution
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "LOGNORM.DIST",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: lognorm_dist_fn,
    }
}

fn lognorm_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let mean = array_lift::eval_arg(ctx, &args[1]);
    let std_dev = array_lift::eval_arg(ctx, &args[2]);
    let cumulative = array_lift::eval_arg(ctx, &args[3]);
    array_lift::lift4(
        x,
        mean,
        std_dev,
        cumulative,
        |x, mean, std_dev, cumulative| {
            let x = x.coerce_to_number_with_ctx(ctx)?;
            let mean = mean.coerce_to_number_with_ctx(ctx)?;
            let std_dev = std_dev.coerce_to_number_with_ctx(ctx)?;
            let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
            Ok(Value::Number(
                crate::functions::statistical::distributions::lognorm_dist(
                    x, mean, std_dev, cumulative,
                )?,
            ))
        },
    )
}

inventory::submit! {
    FunctionSpec {
        name: "LOGNORMDIST",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: lognormdist_fn,
    }
}

fn lognormdist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let mean = array_lift::eval_arg(ctx, &args[1]);
    let std_dev = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(x, mean, std_dev, |x, mean, std_dev| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let mean = mean.coerce_to_number_with_ctx(ctx)?;
        let std_dev = std_dev.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::lognorm_dist(x, mean, std_dev, true)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "LOGNORM.INV",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: lognorm_inv_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "LOGINV",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: lognorm_inv_fn,
    }
}

fn lognorm_inv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let p = array_lift::eval_arg(ctx, &args[0]);
    let mean = array_lift::eval_arg(ctx, &args[1]);
    let std_dev = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(p, mean, std_dev, |p, mean, std_dev| {
        let p = p.coerce_to_number_with_ctx(ctx)?;
        let mean = mean.coerce_to_number_with_ctx(ctx)?;
        let std_dev = std_dev.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::lognorm_inv(p, mean, std_dev)?,
        ))
    })
}

// ----------------------------------------------------------------------
// Exponential distribution
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "EXPON.DIST",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: expon_dist_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "EXPONDIST",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: expon_dist_fn,
    }
}

fn expon_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let lambda = array_lift::eval_arg(ctx, &args[1]);
    let cumulative = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(x, lambda, cumulative, |x, lambda, cumulative| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let lambda = lambda.coerce_to_number_with_ctx(ctx)?;
        let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::expon_dist(x, lambda, cumulative)?,
        ))
    })
}

// ----------------------------------------------------------------------
// Weibull distribution
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "WEIBULL.DIST",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: weibull_dist_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "WEIBULL",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Bool],
        implementation: weibull_dist_fn,
    }
}

fn weibull_dist_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    let alpha = array_lift::eval_arg(ctx, &args[1]);
    let beta = array_lift::eval_arg(ctx, &args[2]);
    let cumulative = array_lift::eval_arg(ctx, &args[3]);
    array_lift::lift4(x, alpha, beta, cumulative, |x, alpha, beta, cumulative| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        let alpha = alpha.coerce_to_number_with_ctx(ctx)?;
        let beta = beta.coerce_to_number_with_ctx(ctx)?;
        let cumulative = cumulative.coerce_to_bool_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::weibull_dist(x, alpha, beta, cumulative)?,
        ))
    })
}

// ----------------------------------------------------------------------
// Fisher transformation
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "FISHER",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: fisher_fn,
    }
}

fn fisher_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let x = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(x, |x| {
        let x = x.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::fisher(x)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "FISHERINV",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: fisherinv_fn,
    }
}

fn fisherinv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let y = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(y, |y| {
        let y = y.coerce_to_number_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::fisherinv(y)?,
        ))
    })
}

// ----------------------------------------------------------------------
// Confidence intervals
// ----------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "CONFIDENCE.NORM",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: confidence_norm_fn,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "CONFIDENCE",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: confidence_norm_fn,
    }
}

fn confidence_norm_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let alpha = array_lift::eval_arg(ctx, &args[0]);
    let std_dev = array_lift::eval_arg(ctx, &args[1]);
    let size = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(alpha, std_dev, size, |alpha, std_dev, size| {
        let alpha = alpha.coerce_to_number_with_ctx(ctx)?;
        let std_dev = std_dev.coerce_to_number_with_ctx(ctx)?;
        let size = size.coerce_to_i64_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::confidence_norm(alpha, std_dev, size)?,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "CONFIDENCE.T",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: confidence_t_fn,
    }
}

fn confidence_t_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let alpha = array_lift::eval_arg(ctx, &args[0]);
    let std_dev = array_lift::eval_arg(ctx, &args[1]);
    let size = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(alpha, std_dev, size, |alpha, std_dev, size| {
        let alpha = alpha.coerce_to_number_with_ctx(ctx)?;
        let std_dev = std_dev.coerce_to_number_with_ctx(ctx)?;
        let size = size.coerce_to_i64_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::statistical::distributions::confidence_t(alpha, std_dev, size)?,
        ))
    })
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
