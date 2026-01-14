use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::Value;

use super::builtins_helpers::{
    basis_from_optional_arg, coerce_to_finite_number, datevalue_from_value, excel_result_number,
};

inventory::submit! {
    FunctionSpec {
        name: "DISC",
        min_args: 4,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: disc_fn,
    }
}

fn disc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let pr = eval_scalar_arg(ctx, &args[2]);
    let redemption = eval_scalar_arg(ctx, &args[3]);
    let basis = match basis_from_optional_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match coerce_to_finite_number(ctx, &pr) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(ctx, &redemption) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::disc(
        settlement, maturity, pr, redemption, basis, system,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "PRICEDISC",
        min_args: 4,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: pricedisc_fn,
    }
}

fn pricedisc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let discount = eval_scalar_arg(ctx, &args[2]);
    let redemption = eval_scalar_arg(ctx, &args[3]);
    let basis = match basis_from_optional_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let discount = match coerce_to_finite_number(ctx, &discount) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(ctx, &redemption) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::pricedisc(
        settlement, maturity, discount, redemption, basis, system,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "YIELDDISC",
        min_args: 4,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: yielddisc_fn,
    }
}

fn yielddisc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let pr = eval_scalar_arg(ctx, &args[2]);
    let redemption = eval_scalar_arg(ctx, &args[3]);
    let basis = match basis_from_optional_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match coerce_to_finite_number(ctx, &pr) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(ctx, &redemption) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::yielddisc(
        settlement, maturity, pr, redemption, basis, system,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "INTRATE",
        min_args: 4,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: intrate_fn,
    }
}

fn intrate_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let investment = eval_scalar_arg(ctx, &args[2]);
    let redemption = eval_scalar_arg(ctx, &args[3]);
    let basis = match basis_from_optional_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let investment = match coerce_to_finite_number(ctx, &investment) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let redemption = match coerce_to_finite_number(ctx, &redemption) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::intrate(
        settlement, maturity, investment, redemption, basis, system,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "RECEIVED",
        min_args: 4,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: received_fn,
    }
}

fn received_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let investment = eval_scalar_arg(ctx, &args[2]);
    let discount = eval_scalar_arg(ctx, &args[3]);
    let basis = match basis_from_optional_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let investment = match coerce_to_finite_number(ctx, &investment) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let discount = match coerce_to_finite_number(ctx, &discount) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::received(
        settlement, maturity, investment, discount, basis, system,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "PRICEMAT",
        min_args: 5,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: pricemat_fn,
    }
}

fn pricemat_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let issue = eval_scalar_arg(ctx, &args[2]);
    let rate = eval_scalar_arg(ctx, &args[3]);
    let yld = eval_scalar_arg(ctx, &args[4]);
    let basis = match basis_from_optional_arg(ctx, args.get(5)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let issue = match datevalue_from_value(ctx, &issue, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(ctx, &rate) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let yld = match coerce_to_finite_number(ctx, &yld) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::pricemat(
        settlement, maturity, issue, rate, yld, basis, system,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "YIELDMAT",
        min_args: 5,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: yieldmat_fn,
    }
}

fn yieldmat_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let issue = eval_scalar_arg(ctx, &args[2]);
    let rate = eval_scalar_arg(ctx, &args[3]);
    let pr = eval_scalar_arg(ctx, &args[4]);
    let basis = match basis_from_optional_arg(ctx, args.get(5)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let issue = match datevalue_from_value(ctx, &issue, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let rate = match coerce_to_finite_number(ctx, &rate) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match coerce_to_finite_number(ctx, &pr) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::yieldmat(
        settlement, maturity, issue, rate, pr, basis, system,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "TBILLPRICE",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number],
        implementation: tbillprice_fn,
    }
}

fn tbillprice_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let discount = eval_scalar_arg(ctx, &args[2]);
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let discount = match coerce_to_finite_number(ctx, &discount) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::tbillprice(settlement, maturity, discount))
}

inventory::submit! {
    FunctionSpec {
        name: "TBILLYIELD",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number],
        implementation: tbillyield_fn,
    }
}

fn tbillyield_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let pr = eval_scalar_arg(ctx, &args[2]);
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let pr = match coerce_to_finite_number(ctx, &pr) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::tbillyield(settlement, maturity, pr))
}

inventory::submit! {
    FunctionSpec {
        name: "TBILLEQ",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number],
        implementation: tbilleq_fn,
    }
}

fn tbilleq_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let discount = eval_scalar_arg(ctx, &args[2]);
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let discount = match coerce_to_finite_number(ctx, &discount) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::tbilleq(settlement, maturity, discount))
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function (indirectly via
// `financial::__force_link`) ensures the module (and its `inventory::submit!` entries) are
// retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
