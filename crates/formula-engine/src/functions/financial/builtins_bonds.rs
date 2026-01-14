use crate::eval::CompiledExpr;
use crate::functions::{ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::Value;

use super::builtins_helpers::{
    basis_from_optional_arg, eval_date_arg, eval_finite_number_arg, eval_i32_trunc_arg,
    excel_result_number, excel_result_serial,
};
// ---------------------------------------------------------------------
// COUP* schedule functions
// ---------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "COUPDAYBS",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: coupdaybs_fn,
    }
}

fn coupdaybs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::coupdaybs(
        settlement,
        maturity,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "COUPDAYS",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: coupdays_fn,
    }
}

fn coupdays_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::coupdays(
        settlement,
        maturity,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "COUPDAYSNC",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: coupdaysnc_fn,
    }
}

fn coupdaysnc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::coupdaysnc(
        settlement,
        maturity,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "COUPNCD",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: coupncd_fn,
    }
}

fn coupncd_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_serial(super::coupncd(
        settlement,
        maturity,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "COUPNUM",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: coupnum_fn,
    }
}

fn coupnum_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::coupnum(
        settlement,
        maturity,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "COUPPCD",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: couppcd_fn,
    }
}

fn couppcd_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_serial(super::couppcd(
        settlement,
        maturity,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

// ---------------------------------------------------------------------
// PRICE
// ---------------------------------------------------------------------
inventory::submit! {
    FunctionSpec {
        name: "PRICE",
        min_args: 6,
        max_args: 7,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any,
            ValueType::Any,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: price_fn,
    }
}

fn price_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match eval_finite_number_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match eval_finite_number_arg(ctx, &args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(ctx, args.get(6)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::price(
        settlement,
        maturity,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

// ---------------------------------------------------------------------
// YIELD
// ---------------------------------------------------------------------
inventory::submit! {
    FunctionSpec {
        name: "YIELD",
        min_args: 6,
        max_args: 7,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any,
            ValueType::Any,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: yield_fn,
    }
}

fn yield_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match eval_finite_number_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match eval_finite_number_arg(ctx, &args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[5]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(ctx, args.get(6)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::yield_rate(
        settlement,
        maturity,
        rate,
        pr,
        redemption,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

// ---------------------------------------------------------------------
// DURATION
// ---------------------------------------------------------------------
inventory::submit! {
    FunctionSpec {
        name: "DURATION",
        min_args: 5,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any,
            ValueType::Any,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: duration_fn,
    }
}

fn duration_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let coupon = match eval_finite_number_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(ctx, args.get(5)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::duration(
        settlement,
        maturity,
        coupon,
        yld,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

// ---------------------------------------------------------------------
// MDURATION
// ---------------------------------------------------------------------
inventory::submit! {
    FunctionSpec {
        name: "MDURATION",
        min_args: 5,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any,
            ValueType::Any,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: mduration_fn,
    }
}

fn mduration_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = match eval_date_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match eval_date_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let coupon = match eval_finite_number_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match eval_i32_trunc_arg(ctx, &args[4]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let basis = match basis_from_optional_arg(ctx, args.get(5)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::mduration(
        settlement,
        maturity,
        coupon,
        yld,
        frequency,
        basis,
        ctx.date_system(),
    ))
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function (indirectly via
// `financial::__force_link`) ensures the module (and its `inventory::submit!` entries) are
// retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
