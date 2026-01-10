use chrono::{Datelike, Timelike};

use crate::date::{serial_to_ymd, ymd_to_serial, ExcelDate, ExcelDateSystem};
use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

inventory::submit! {
    FunctionSpec {
        name: "DATE",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: date_fn,
    }
}

fn date_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let year = match eval_scalar_arg(ctx, &args[0]).coerce_to_i64() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let month = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let day = match eval_scalar_arg(ctx, &args[2]).coerce_to_i64() {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let (year, month) = normalize_year_month(year, month);
    let year_i32 = match i32::try_from(year) {
        Ok(y) => y,
        Err(_) => return Value::Error(ErrorKind::Num),
    };
    let month_u8 = match u8::try_from(month) {
        Ok(m) if (1..=12).contains(&m) => m,
        _ => return Value::Error(ErrorKind::Num),
    };

    let system = ExcelDateSystem::EXCEL_1900;
    let first_serial = match ymd_to_serial(ExcelDate::new(year_i32, month_u8, 1), system) {
        Ok(s) => s as i64,
        Err(_) => return Value::Error(ErrorKind::Num),
    };
    let serial = first_serial + (day - 1);
    if serial < i64::from(i32::MIN) || serial > i64::from(i32::MAX) {
        return Value::Error(ErrorKind::Num);
    }
    Value::Number(serial as f64)
}

inventory::submit! {
    FunctionSpec {
        name: "TODAY",
        min_args: 0,
        max_args: 0,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[],
        implementation: today_fn,
    }
}

fn today_fn(ctx: &dyn FunctionContext, _args: &[CompiledExpr]) -> Value {
    let now = ctx.now_utc();
    let date = now.date_naive();
    let system = ExcelDateSystem::EXCEL_1900;
    match ymd_to_serial(
        ExcelDate::new(date.year(), date.month() as u8, date.day() as u8),
        system,
    ) {
        Ok(s) => Value::Number(s as f64),
        Err(_) => Value::Error(ErrorKind::Num),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "NOW",
        min_args: 0,
        max_args: 0,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[],
        implementation: now_fn,
    }
}

fn now_fn(ctx: &dyn FunctionContext, _args: &[CompiledExpr]) -> Value {
    let now = ctx.now_utc();
    let date = now.date_naive();
    let system = ExcelDateSystem::EXCEL_1900;
    let base = match ymd_to_serial(
        ExcelDate::new(date.year(), date.month() as u8, date.day() as u8),
        system,
    ) {
        Ok(s) => s as f64,
        Err(_) => return Value::Error(ErrorKind::Num),
    };
    let seconds = now.time().num_seconds_from_midnight() as f64
        + (now.time().nanosecond() as f64 / 1_000_000_000.0);
    Value::Number(base + seconds / 86_400.0)
}

inventory::submit! {
    FunctionSpec {
        name: "YEAR",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: year_fn,
    }
}

fn year_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    match serial_to_components(eval_scalar_arg(ctx, &args[0])) {
        Ok((y, _, _)) => Value::Number(y as f64),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "MONTH",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: month_fn,
    }
}

fn month_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    match serial_to_components(eval_scalar_arg(ctx, &args[0])) {
        Ok((_, m, _)) => Value::Number(m as f64),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DAY",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: day_fn,
    }
}

fn day_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    match serial_to_components(eval_scalar_arg(ctx, &args[0])) {
        Ok((_, _, d)) => Value::Number(d as f64),
        Err(e) => Value::Error(e),
    }
}

fn serial_to_components(v: Value) -> Result<(i32, i32, i32), ErrorKind> {
    let n = v.coerce_to_number()?;
    if !n.is_finite() {
        return Err(ErrorKind::Num);
    }
    let serial = n.floor();
    let serial_i32 = i32::try_from(serial as i64).map_err(|_| ErrorKind::Num)?;
    let system = ExcelDateSystem::EXCEL_1900;
    let date = serial_to_ymd(serial_i32, system).map_err(|_| ErrorKind::Num)?;
    Ok((date.year, date.month as i32, date.day as i32))
}

fn normalize_year_month(year: i64, month: i64) -> (i64, i64) {
    let total_months = year * 12 + (month - 1);
    let new_year = total_months.div_euclid(12);
    let new_month = total_months.rem_euclid(12) + 1;
    (new_year, new_month)
}
