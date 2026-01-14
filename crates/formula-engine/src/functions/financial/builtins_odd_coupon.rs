use chrono::{DateTime, Utc};

use crate::date::{serial_to_ymd, ExcelDateSystem};
use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

use super::builtins_helpers::{
    basis_from_optional_arg, coerce_to_finite_number, coerce_to_i32_trunc, datevalue_from_value,
    excel_error_kind, excel_result_number,
};

fn datevalue_checked_from_value(
    ctx: &dyn FunctionContext,
    value: &Value,
    system: ExcelDateSystem,
    now_utc: DateTime<Utc>,
) -> Result<i32, ErrorKind> {
    let serial = datevalue_from_value(ctx, value, system, now_utc)?;

    // Ensure the resulting serial is representable in the workbook date system.
    serial_to_ymd(serial, system).map_err(excel_error_kind)?;

    Ok(serial)
}

fn frequency_from_value(ctx: &dyn FunctionContext, v: &Value) -> Result<i32, ErrorKind> {
    let frequency = coerce_to_i32_trunc(ctx, v)?;
    super::coupon_schedule::validate_frequency(frequency).map_err(excel_error_kind)
}

inventory::submit! {
    FunctionSpec {
        name: "ODDFPRICE",
        min_args: 8,
        max_args: 9,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any, // settlement
            ValueType::Any, // maturity
            ValueType::Any, // issue
            ValueType::Any, // first_coupon
            ValueType::Number, // rate
            ValueType::Number, // yld
            ValueType::Number, // redemption
            ValueType::Number, // frequency
            ValueType::Number, // basis
        ],
        implementation: oddfprice_fn,
    }
}

fn oddfprice_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let issue = eval_scalar_arg(ctx, &args[2]);
    let first_coupon = eval_scalar_arg(ctx, &args[3]);
    let rate = eval_scalar_arg(ctx, &args[4]);
    let yld = eval_scalar_arg(ctx, &args[5]);
    let redemption = eval_scalar_arg(ctx, &args[6]);
    let frequency = eval_scalar_arg(ctx, &args[7]);
    let basis = match basis_from_optional_arg(ctx, args.get(8)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_checked_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_checked_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let issue = match datevalue_checked_from_value(ctx, &issue, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let first_coupon = match datevalue_checked_from_value(ctx, &first_coupon, system, now_utc) {
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
    let redemption = match coerce_to_finite_number(ctx, &redemption) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match frequency_from_value(ctx, &frequency) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "ODDFYIELD",
        min_args: 8,
        max_args: 9,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any, // settlement
            ValueType::Any, // maturity
            ValueType::Any, // issue
            ValueType::Any, // first_coupon
            ValueType::Number, // rate
            ValueType::Number, // pr
            ValueType::Number, // redemption
            ValueType::Number, // frequency
            ValueType::Number, // basis
        ],
        implementation: oddfyield_fn,
    }
}

fn oddfyield_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let issue = eval_scalar_arg(ctx, &args[2]);
    let first_coupon = eval_scalar_arg(ctx, &args[3]);
    let rate = eval_scalar_arg(ctx, &args[4]);
    let pr = eval_scalar_arg(ctx, &args[5]);
    let redemption = eval_scalar_arg(ctx, &args[6]);
    let frequency = eval_scalar_arg(ctx, &args[7]);
    let basis = match basis_from_optional_arg(ctx, args.get(8)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_checked_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_checked_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let issue = match datevalue_checked_from_value(ctx, &issue, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let first_coupon = match datevalue_checked_from_value(ctx, &first_coupon, system, now_utc) {
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
    let redemption = match coerce_to_finite_number(ctx, &redemption) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match frequency_from_value(ctx, &frequency) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        pr,
        redemption,
        frequency,
        basis,
        system,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "ODDLPRICE",
        min_args: 7,
        max_args: 8,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any, // settlement
            ValueType::Any, // maturity
            ValueType::Any, // last_interest
            ValueType::Number, // rate
            ValueType::Number, // yld
            ValueType::Number, // redemption
            ValueType::Number, // frequency
            ValueType::Number, // basis
        ],
        implementation: oddlprice_fn,
    }
}

fn oddlprice_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let last_interest = eval_scalar_arg(ctx, &args[2]);
    let rate = eval_scalar_arg(ctx, &args[3]);
    let yld = eval_scalar_arg(ctx, &args[4]);
    let redemption = eval_scalar_arg(ctx, &args[5]);
    let frequency = eval_scalar_arg(ctx, &args[6]);
    let basis = match basis_from_optional_arg(ctx, args.get(7)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_checked_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_checked_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let last_interest = match datevalue_checked_from_value(ctx, &last_interest, system, now_utc) {
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
    let redemption = match coerce_to_finite_number(ctx, &redemption) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match frequency_from_value(ctx, &frequency) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    ))
}

inventory::submit! {
    FunctionSpec {
        name: "ODDLYIELD",
        min_args: 7,
        max_args: 8,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[
            ValueType::Any, // settlement
            ValueType::Any, // maturity
            ValueType::Any, // last_interest
            ValueType::Number, // rate
            ValueType::Number, // pr
            ValueType::Number, // redemption
            ValueType::Number, // frequency
            ValueType::Number, // basis
        ],
        implementation: oddlyield_fn,
    }
}

fn oddlyield_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let settlement = eval_scalar_arg(ctx, &args[0]);
    let maturity = eval_scalar_arg(ctx, &args[1]);
    let last_interest = eval_scalar_arg(ctx, &args[2]);
    let rate = eval_scalar_arg(ctx, &args[3]);
    let pr = eval_scalar_arg(ctx, &args[4]);
    let redemption = eval_scalar_arg(ctx, &args[5]);
    let frequency = eval_scalar_arg(ctx, &args[6]);
    let basis = match basis_from_optional_arg(ctx, args.get(7)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();

    let settlement = match datevalue_checked_from_value(ctx, &settlement, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let maturity = match datevalue_checked_from_value(ctx, &maturity, system, now_utc) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let last_interest = match datevalue_checked_from_value(ctx, &last_interest, system, now_utc) {
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
    let redemption = match coerce_to_finite_number(ctx, &redemption) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let frequency = match frequency_from_value(ctx, &frequency) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::oddlyield(
        settlement,
        maturity,
        last_interest,
        rate,
        pr,
        redemption,
        frequency,
        basis,
        system,
    ))
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function (indirectly via
// `financial::__force_link`) ensures the module (and its `inventory::submit!` entries) are
// retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
