use chrono::{DateTime, Datelike, Timelike, Utc};

use crate::coercion::ValueLocaleConfig;
use crate::date::{serial_to_ymd, ymd_to_serial, ExcelDate, ExcelDateSystem};
use crate::error::ExcelError;
use crate::eval::CompiledExpr;
use crate::functions::array_lift;
use crate::functions::date_time;
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{Array, ErrorKind, Value};

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
    let year = eval_scalar_arg(ctx, &args[0]);
    let month = eval_scalar_arg(ctx, &args[1]);
    let day = eval_scalar_arg(ctx, &args[2]);
    let system = ctx.date_system();
    broadcast_map3(year, month, day, |y, m, d| {
        date_from_parts(ctx, &y, &m, &d, system)
    })
}

fn date_from_parts(
    ctx: &dyn FunctionContext,
    year: &Value,
    month: &Value,
    day: &Value,
    system: ExcelDateSystem,
) -> Value {
    let year = match coerce_to_i64_trunc(ctx, year) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let month = match coerce_to_i64_trunc(ctx, month) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let day = match coerce_to_i64_trunc(ctx, day) {
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

    let first_serial = match ymd_to_serial(ExcelDate::new(year_i32, month_u8, 1), system) {
        Ok(s) => i64::from(s),
        Err(_) => return Value::Error(ErrorKind::Num),
    };
    let day_offset = match day.checked_sub(1) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Num),
    };
    let serial = match first_serial.checked_add(day_offset) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Num),
    };
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
    let system = ctx.date_system();
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
    let system = ctx.date_system();
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
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: year_fn,
    }
}

fn year_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    map_serial_arg(
        ctx,
        array_lift::eval_arg(ctx, &args[0]),
        ctx.date_system(),
        |(y, _, _)| y as f64,
    )
}

inventory::submit! {
    FunctionSpec {
        name: "MONTH",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: month_fn,
    }
}

fn month_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    map_serial_arg(
        ctx,
        array_lift::eval_arg(ctx, &args[0]),
        ctx.date_system(),
        |(_, m, _)| m as f64,
    )
}

inventory::submit! {
    FunctionSpec {
        name: "DAY",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: day_fn,
    }
}

fn day_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    map_serial_arg(
        ctx,
        array_lift::eval_arg(ctx, &args[0]),
        ctx.date_system(),
        |(_, _, d)| d as f64,
    )
}

inventory::submit! {
    FunctionSpec {
        name: "TIME",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: time_fn,
    }
}

fn time_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let hour = eval_scalar_arg(ctx, &args[0]);
    let minute = eval_scalar_arg(ctx, &args[1]);
    let second = eval_scalar_arg(ctx, &args[2]);
    broadcast_map3(hour, minute, second, |h, m, s| {
        match time_from_parts(ctx, &h, &m, &s) {
            Ok(v) => Value::Number(v),
            Err(e) => Value::Error(e),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "HOUR",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: hour_fn,
    }
}

fn hour_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let cfg = ctx.value_locale();
    map_unary(
        array_lift::eval_arg(ctx, &args[0]),
        |v| match time_components_from_value(ctx, &v, cfg) {
            Ok((h, _, _)) => Value::Number(h as f64),
            Err(e) => Value::Error(e),
        },
    )
}

inventory::submit! {
    FunctionSpec {
        name: "MINUTE",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: minute_fn,
    }
}

fn minute_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let cfg = ctx.value_locale();
    map_unary(
        array_lift::eval_arg(ctx, &args[0]),
        |v| match time_components_from_value(ctx, &v, cfg) {
            Ok((_, m, _)) => Value::Number(m as f64),
            Err(e) => Value::Error(e),
        },
    )
}

inventory::submit! {
    FunctionSpec {
        name: "SECOND",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: second_fn,
    }
}

fn second_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let cfg = ctx.value_locale();
    map_unary(
        array_lift::eval_arg(ctx, &args[0]),
        |v| match time_components_from_value(ctx, &v, cfg) {
            Ok((_, _, s)) => Value::Number(s as f64),
            Err(e) => Value::Error(e),
        },
    )
}

inventory::submit! {
    FunctionSpec {
        name: "TIMEVALUE",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: timevalue_fn,
    }
}

fn timevalue_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let cfg = ctx.value_locale();
    map_unary(
        eval_scalar_arg(ctx, &args[0]),
        |v| match timevalue_from_value(ctx, &v, cfg) {
            Ok(n) => Value::Number(n),
            Err(e) => Value::Error(e),
        },
    )
}

inventory::submit! {
    FunctionSpec {
        name: "DATEVALUE",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any],
        implementation: datevalue_fn,
    }
}

fn datevalue_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();
    let cfg = ctx.value_locale();
    map_unary(
        eval_scalar_arg(ctx, &args[0]),
        |v| match datevalue_from_value(ctx, &v, system, cfg, now_utc) {
            Ok(n) => Value::Number(n as f64),
            Err(e) => Value::Error(e),
        },
    )
}

inventory::submit! {
    FunctionSpec {
        name: "DAYS",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: days_fn,
    }
}

fn days_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let end_date = eval_scalar_arg(ctx, &args[0]);
    let start_date = eval_scalar_arg(ctx, &args[1]);
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();
    let cfg = ctx.value_locale();
    broadcast_map2(end_date, start_date, |end_date, start_date| {
        let end_serial = match datevalue_from_value(ctx, &end_date, system, cfg, now_utc) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let start_serial = match datevalue_from_value(ctx, &start_date, system, cfg, now_utc) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        Value::Number((i64::from(end_serial) - i64::from(start_serial)) as f64)
    })
}

inventory::submit! {
    FunctionSpec {
        name: "DAYS360",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Bool],
        implementation: days360_fn,
    }
}

fn days360_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let start_date = eval_scalar_arg(ctx, &args[0]);
    let end_date = eval_scalar_arg(ctx, &args[1]);
    let method = if args.len() == 3 {
        eval_scalar_arg(ctx, &args[2])
    } else {
        Value::Blank
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();
    let cfg = ctx.value_locale();

    broadcast_map3(
        start_date,
        end_date,
        method,
        |start_date, end_date, method| {
            let start_serial = match datevalue_from_value(ctx, &start_date, system, cfg, now_utc) {
                Ok(v) => v,
                Err(e) => return Value::Error(e),
            };
            let end_serial = match datevalue_from_value(ctx, &end_date, system, cfg, now_utc) {
                Ok(v) => v,
                Err(e) => return Value::Error(e),
            };
            let method = match coerce_to_bool_finite(ctx, &method) {
                Ok(v) => v,
                Err(e) => return Value::Error(e),
            };
            match date_time::days360(start_serial, end_serial, method, system) {
                Ok(v) => Value::Number(v as f64),
                Err(e) => Value::Error(excel_error_kind(e)),
            }
        },
    )
}

inventory::submit! {
    FunctionSpec {
        name: "YEARFRAC",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number],
        implementation: yearfrac_fn,
    }
}

fn yearfrac_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let start_date = eval_scalar_arg(ctx, &args[0]);
    let end_date = eval_scalar_arg(ctx, &args[1]);
    let basis = if args.len() == 3 {
        eval_scalar_arg(ctx, &args[2])
    } else {
        Value::Blank
    };
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();
    let cfg = ctx.value_locale();

    broadcast_map3(
        start_date,
        end_date,
        basis,
        |start_date, end_date, basis| {
            let start_serial = match datevalue_from_value(ctx, &start_date, system, cfg, now_utc) {
                Ok(v) => v,
                Err(e) => return Value::Error(e),
            };
            let end_serial = match datevalue_from_value(ctx, &end_date, system, cfg, now_utc) {
                Ok(v) => v,
                Err(e) => return Value::Error(e),
            };

            let basis = if matches!(basis, Value::Blank) {
                0
            } else {
                match coerce_to_i32_trunc(ctx, &basis) {
                    Ok(v) => v,
                    Err(e) => return Value::Error(e),
                }
            };

            match date_time::yearfrac(start_serial, end_serial, basis, system) {
                Ok(v) => Value::Number(v),
                Err(e) => Value::Error(excel_error_kind(e)),
            }
        },
    )
}

inventory::submit! {
    FunctionSpec {
        name: "DATEDIF",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Text],
        implementation: datedif_fn,
    }
}

fn datedif_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let start_date = eval_scalar_arg(ctx, &args[0]);
    let end_date = eval_scalar_arg(ctx, &args[1]);
    let unit = eval_scalar_arg(ctx, &args[2]);
    let system = ctx.date_system();
    let now_utc = ctx.now_utc();
    let cfg = ctx.value_locale();

    broadcast_map3(start_date, end_date, unit, |start_date, end_date, unit| {
        let start_serial = match datevalue_from_value(ctx, &start_date, system, cfg, now_utc) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let end_serial = match datevalue_from_value(ctx, &end_date, system, cfg, now_utc) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };

        if let Value::Number(n) = &unit {
            if !n.is_finite() {
                return Value::Error(ErrorKind::Num);
            }
        }
        let unit = match unit.coerce_to_string_with_ctx(ctx) {
            Ok(s) => s,
            Err(e) => return Value::Error(e),
        };

        match date_time::datedif(start_serial, end_serial, &unit, system) {
            Ok(v) => Value::Number(v as f64),
            Err(e) => Value::Error(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "EOMONTH",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: eomonth_fn,
    }
}

fn eomonth_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let start_date = eval_scalar_arg(ctx, &args[0]);
    let months = eval_scalar_arg(ctx, &args[1]);
    let system = ctx.date_system();
    broadcast_map2(start_date, months, |start, months| {
        let start_serial = match coerce_to_serial_floor(ctx, &start) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let months = match coerce_to_i32_trunc(ctx, &months) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        match date_time::eomonth(start_serial, months, system) {
            Ok(serial) => Value::Number(serial as f64),
            Err(e) => Value::Error(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "EDATE",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: edate_fn,
    }
}

fn edate_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let start_date = eval_scalar_arg(ctx, &args[0]);
    let months = eval_scalar_arg(ctx, &args[1]);
    let system = ctx.date_system();
    broadcast_map2(start_date, months, |start, months| {
        let start_serial = match coerce_to_serial_floor(ctx, &start) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let months = match coerce_to_i32_trunc(ctx, &months) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        match date_time::edate(start_serial, months, system) {
            Ok(serial) => Value::Number(serial as f64),
            Err(e) => Value::Error(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "WEEKDAY",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: weekday_fn,
    }
}

fn weekday_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let serial = eval_scalar_arg(ctx, &args[0]);
    let system = ctx.date_system();
    let return_type = if args.len() == 2 {
        Some(eval_scalar_arg(ctx, &args[1]))
    } else {
        None
    };

    match return_type {
        None => map_unary(serial, |serial| weekday_scalar(ctx, &serial, None, system)),
        Some(rt) => broadcast_map2(serial, rt, |serial, rt| {
            let rt = match coerce_to_i32_trunc(ctx, &rt) {
                Ok(v) => Some(v),
                Err(e) => return Value::Error(e),
            };
            weekday_scalar(ctx, &serial, rt, system)
        }),
    }
}

fn weekday_scalar(
    ctx: &dyn FunctionContext,
    serial: &Value,
    return_type: Option<i32>,
    system: ExcelDateSystem,
) -> Value {
    let serial = match coerce_to_serial_floor(ctx, serial) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match date_time::weekday(serial, return_type, system) {
        Ok(v) => Value::Number(v as f64),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "WEEKNUM",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: weeknum_fn,
    }
}

fn weeknum_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let serial = eval_scalar_arg(ctx, &args[0]);
    let system = ctx.date_system();
    let return_type = if args.len() == 2 {
        Some(eval_scalar_arg(ctx, &args[1]))
    } else {
        None
    };

    match return_type {
        None => map_unary(serial, |serial| weeknum_scalar(ctx, &serial, None, system)),
        Some(rt) => broadcast_map2(serial, rt, |serial, rt| {
            let rt = match coerce_to_i32_trunc(ctx, &rt) {
                Ok(v) => Some(v),
                Err(e) => return Value::Error(e),
            };
            weeknum_scalar(ctx, &serial, rt, system)
        }),
    }
}

fn weeknum_scalar(
    ctx: &dyn FunctionContext,
    serial: &Value,
    return_type: Option<i32>,
    system: ExcelDateSystem,
) -> Value {
    let serial = match coerce_to_serial_floor(ctx, serial) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match date_time::weeknum(serial, return_type, system) {
        Ok(v) => Value::Number(v as f64),
        Err(e) => Value::Error(excel_error_kind(e)),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "ISOWEEKNUM",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: isoweeknum_fn,
    }
}

// Excel stores ISOWEEKNUM with an `_xlfn.` prefix in older file formats.
// Some spreadsheets also surface it as `ISO.WEEKNUM`; register an alias for compatibility.
inventory::submit! {
    FunctionSpec {
        name: "ISO.WEEKNUM",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: isoweeknum_fn,
    }
}

fn isoweeknum_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let serial = eval_scalar_arg(ctx, &args[0]);
    let system = ctx.date_system();
    map_unary(serial, |serial| {
        let serial = match coerce_to_serial_floor(ctx, &serial) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        match date_time::weeknum(serial, Some(21), system) {
            Ok(v) => Value::Number(v as f64),
            Err(e) => Value::Error(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "WORKDAY",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Any],
        implementation: workday_fn,
    }
}

fn workday_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let start_date = eval_scalar_arg(ctx, &args[0]);
    let days = eval_scalar_arg(ctx, &args[1]);
    let system = ctx.date_system();
    let holidays = if args.len() == 3 {
        match collect_holidays(ctx, &args[2], system) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };
    let holidays_ref = holidays.as_deref();

    broadcast_map2(start_date, days, |start, days| {
        let start = match coerce_to_serial_floor(ctx, &start) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let days = match coerce_to_i32_trunc(ctx, &days) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        match date_time::workday(start, days, holidays_ref, system) {
            Ok(v) => Value::Number(v as f64),
            Err(e) => Value::Error(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "NETWORKDAYS",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Any],
        implementation: networkdays_fn,
    }
}

fn networkdays_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let start_date = eval_scalar_arg(ctx, &args[0]);
    let end_date = eval_scalar_arg(ctx, &args[1]);
    let system = ctx.date_system();
    let holidays = if args.len() == 3 {
        match collect_holidays(ctx, &args[2], system) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };
    let holidays_ref = holidays.as_deref();

    broadcast_map2(start_date, end_date, |start, end| {
        let start = match coerce_to_serial_floor(ctx, &start) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let end = match coerce_to_serial_floor(ctx, &end) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        match date_time::networkdays(start, end, holidays_ref, system) {
            Ok(v) => Value::Number(v as f64),
            Err(e) => Value::Error(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "WORKDAY.INTL",
        min_args: 2,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Any, ValueType::Any],
        implementation: workday_intl_fn,
    }
}

fn workday_intl_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let start_date = eval_scalar_arg(ctx, &args[0]);
    let days = eval_scalar_arg(ctx, &args[1]);
    let system = ctx.date_system();
    let weekend = if args.len() >= 3 {
        eval_scalar_arg(ctx, &args[2])
    } else {
        Value::Blank
    };
    let weekend_mask = match parse_weekend_mask(ctx, &weekend) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let holidays = if args.len() == 4 {
        match collect_holidays(ctx, &args[3], system) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };
    let holidays_ref = holidays.as_deref();

    broadcast_map2(start_date, days, |start, days| {
        let start = match coerce_to_serial_floor(ctx, &start) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let days = match coerce_to_i32_trunc(ctx, &days) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        match date_time::workday_intl(start, days, weekend_mask, holidays_ref, system) {
            Ok(v) => Value::Number(v as f64),
            Err(e) => Value::Error(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "NETWORKDAYS.INTL",
        min_args: 2,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Any, ValueType::Any],
        implementation: networkdays_intl_fn,
    }
}

fn networkdays_intl_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let start_date = eval_scalar_arg(ctx, &args[0]);
    let end_date = eval_scalar_arg(ctx, &args[1]);
    let system = ctx.date_system();
    let weekend = if args.len() >= 3 {
        eval_scalar_arg(ctx, &args[2])
    } else {
        Value::Blank
    };
    let weekend_mask = match parse_weekend_mask(ctx, &weekend) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let holidays = if args.len() == 4 {
        match collect_holidays(ctx, &args[3], system) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };
    let holidays_ref = holidays.as_deref();

    broadcast_map2(start_date, end_date, |start, end| {
        let start = match coerce_to_serial_floor(ctx, &start) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        let end = match coerce_to_serial_floor(ctx, &end) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
        match date_time::networkdays_intl(start, end, weekend_mask, holidays_ref, system) {
            Ok(v) => Value::Number(v as f64),
            Err(e) => Value::Error(excel_error_kind(e)),
        }
    })
}

fn serial_to_components(
    ctx: &dyn FunctionContext,
    v: &Value,
    system: ExcelDateSystem,
) -> Result<(i32, i32, i32), ErrorKind> {
    let serial_i32 = coerce_to_serial_floor(ctx, v)?;
    let date = serial_to_ymd(serial_i32, system).map_err(|_| ErrorKind::Num)?;
    Ok((date.year, date.month as i32, date.day as i32))
}

fn map_serial_arg(
    ctx: &dyn FunctionContext,
    v: Value,
    system: ExcelDateSystem,
    f: impl Fn((i32, i32, i32)) -> f64 + Copy,
) -> Value {
    map_unary(v, |v| match serial_to_components(ctx, &v, system) {
        Ok(parts) => Value::Number(f(parts)),
        Err(e) => Value::Error(e),
    })
}

fn excel_error_kind(err: ExcelError) -> ErrorKind {
    match err {
        ExcelError::Div0 => ErrorKind::Div0,
        ExcelError::Num => ErrorKind::Num,
        ExcelError::Value => ErrorKind::Value,
    }
}

fn map_unary(v: Value, mut f: impl FnMut(Value) -> Value) -> Value {
    match v {
        Value::Array(arr) => {
            if arr.rows == 1 && arr.cols == 1 {
                let el = arr.values.into_iter().next().unwrap_or(Value::Blank);
                return f(el);
            }
            let mut out = Vec::with_capacity(arr.values.len());
            for el in arr.values {
                out.push(f(el));
            }
            Value::Array(Array::new(arr.rows, arr.cols, out))
        }
        other => f(other),
    }
}

fn broadcast_map2(a: Value, b: Value, mut f: impl FnMut(Value, Value) -> Value) -> Value {
    let (rows, cols) = match broadcast_shape(&[&a, &b]) {
        Ok(shape) => shape,
        Err(e) => return Value::Error(e),
    };
    if !is_broadcastable(&a, rows, cols) || !is_broadcastable(&b, rows, cols) {
        return Value::Error(ErrorKind::Value);
    }

    if rows == 1 && cols == 1 {
        return f(broadcast_get(&a, 0, 0), broadcast_get(&b, 0, 0));
    }

    let mut out = Vec::with_capacity(rows.saturating_mul(cols));
    for r in 0..rows {
        for c in 0..cols {
            let va = broadcast_get(&a, r, c);
            let vb = broadcast_get(&b, r, c);
            out.push(f(va, vb));
        }
    }
    Value::Array(Array::new(rows, cols, out))
}

fn broadcast_map3(
    a: Value,
    b: Value,
    c: Value,
    mut f: impl FnMut(Value, Value, Value) -> Value,
) -> Value {
    let (rows, cols) = match broadcast_shape(&[&a, &b, &c]) {
        Ok(shape) => shape,
        Err(e) => return Value::Error(e),
    };
    if !is_broadcastable(&a, rows, cols)
        || !is_broadcastable(&b, rows, cols)
        || !is_broadcastable(&c, rows, cols)
    {
        return Value::Error(ErrorKind::Value);
    }

    if rows == 1 && cols == 1 {
        return f(
            broadcast_get(&a, 0, 0),
            broadcast_get(&b, 0, 0),
            broadcast_get(&c, 0, 0),
        );
    }

    let mut out = Vec::with_capacity(rows.saturating_mul(cols));
    for r in 0..rows {
        for c_idx in 0..cols {
            let va = broadcast_get(&a, r, c_idx);
            let vb = broadcast_get(&b, r, c_idx);
            let vc = broadcast_get(&c, r, c_idx);
            out.push(f(va, vb, vc));
        }
    }
    Value::Array(Array::new(rows, cols, out))
}

fn broadcast_shape(values: &[&Value]) -> Result<(usize, usize), ErrorKind> {
    let mut shape: Option<(usize, usize)> = None;
    for v in values {
        if let Value::Array(arr) = v {
            if arr.rows == 1 && arr.cols == 1 {
                continue;
            }
            match shape {
                None => shape = Some((arr.rows, arr.cols)),
                Some((rows, cols)) => {
                    if arr.rows != rows || arr.cols != cols {
                        return Err(ErrorKind::Value);
                    }
                }
            }
        }
    }
    Ok(shape.unwrap_or((1, 1)))
}

fn is_broadcastable(v: &Value, rows: usize, cols: usize) -> bool {
    match v {
        Value::Array(arr) => {
            (arr.rows == 1 && arr.cols == 1) || (arr.rows == rows && arr.cols == cols)
        }
        _ => true,
    }
}

fn broadcast_get(v: &Value, row: usize, col: usize) -> Value {
    match v {
        Value::Array(arr) => {
            let r = if arr.rows == 1 { 0 } else { row };
            let c = if arr.cols == 1 { 0 } else { col };
            arr.get(r, c).cloned().unwrap_or(Value::Blank)
        }
        other => other.clone(),
    }
}

fn time_from_parts(
    ctx: &dyn FunctionContext,
    hour: &Value,
    minute: &Value,
    second: &Value,
) -> Result<f64, ErrorKind> {
    let hour_num = coerce_to_finite_number(ctx, hour)?;
    let minute_num = coerce_to_finite_number(ctx, minute)?;
    let second_num = coerce_to_finite_number(ctx, second)?;
    if hour_num < 0.0 || minute_num < 0.0 || second_num < 0.0 {
        return Err(ErrorKind::Num);
    }

    let hour_i32 = coerce_number_to_i32_trunc(hour_num)?;
    let minute_i32 = coerce_number_to_i32_trunc(minute_num)?;
    let second_i32 = coerce_number_to_i32_trunc(second_num)?;

    date_time::time(hour_i32, minute_i32, second_i32).map_err(excel_error_kind)
}

fn timevalue_from_value(
    ctx: &dyn FunctionContext,
    value: &Value,
    cfg: ValueLocaleConfig,
) -> Result<f64, ErrorKind> {
    match value {
        Value::Text(s) => {
            let n = date_time::timevalue(s, cfg).map_err(excel_error_kind)?;
            Ok(n.rem_euclid(1.0))
        }
        _ => {
            let n = value.coerce_to_number_with_ctx(ctx)?;
            if !n.is_finite() {
                return Err(ErrorKind::Num);
            }
            Ok(n.rem_euclid(1.0))
        }
    }
}

fn datevalue_from_value(
    ctx: &dyn FunctionContext,
    value: &Value,
    system: ExcelDateSystem,
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
) -> Result<i32, ErrorKind> {
    match value {
        Value::Text(s) => date_time::datevalue(s, cfg, now_utc, system).map_err(excel_error_kind),
        _ => {
            let n = value.coerce_to_number_with_ctx(ctx)?;
            if !n.is_finite() {
                return Err(ErrorKind::Num);
            }
            let serial = n.floor();
            if serial < (i32::MIN as f64) || serial > (i32::MAX as f64) {
                return Err(ErrorKind::Num);
            }
            Ok(serial as i32)
        }
    }
}

fn time_components_from_value(
    ctx: &dyn FunctionContext,
    value: &Value,
    cfg: ValueLocaleConfig,
) -> Result<(i32, i32, i32), ErrorKind> {
    let mut fraction = timevalue_from_value(ctx, value, cfg)?;
    fraction = fraction.rem_euclid(1.0);
    let mut total = (fraction * 86_400.0).floor();
    if total >= 86_400.0 {
        total = 86_399.0;
    }
    let total = total as i32;
    let hour = total / 3600;
    let minute = (total % 3600) / 60;
    let second = total % 60;
    Ok((hour, minute, second))
}

fn coerce_to_finite_number(ctx: &dyn FunctionContext, v: &Value) -> Result<f64, ErrorKind> {
    let n = v.coerce_to_number_with_ctx(ctx)?;
    if !n.is_finite() {
        return Err(ErrorKind::Num);
    }
    Ok(n)
}

fn coerce_to_bool_finite(ctx: &dyn FunctionContext, v: &Value) -> Result<bool, ErrorKind> {
    match v {
        Value::Number(n) => {
            if !n.is_finite() {
                return Err(ErrorKind::Num);
            }
            Ok(*n != 0.0)
        }
        _ => v.coerce_to_bool_with_ctx(ctx),
    }
}

fn coerce_number_to_i32_trunc(n: f64) -> Result<i32, ErrorKind> {
    let t = n.trunc();
    if t < (i32::MIN as f64) || t > (i32::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(t as i32)
}

fn coerce_to_i32_trunc(ctx: &dyn FunctionContext, v: &Value) -> Result<i32, ErrorKind> {
    let n = coerce_to_finite_number(ctx, v)?;
    coerce_number_to_i32_trunc(n)
}

fn coerce_to_i64_trunc(ctx: &dyn FunctionContext, v: &Value) -> Result<i64, ErrorKind> {
    let n = coerce_to_finite_number(ctx, v)?;
    let t = n.trunc();
    if t < (i64::MIN as f64) || t > (i64::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(t as i64)
}

fn coerce_to_serial_floor(ctx: &dyn FunctionContext, v: &Value) -> Result<i32, ErrorKind> {
    let n = coerce_to_finite_number(ctx, v)?;
    let serial = n.floor();
    if serial < (i32::MIN as f64) || serial > (i32::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(serial as i32)
}

fn collect_holidays(
    ctx: &dyn FunctionContext,
    expr: &CompiledExpr,
    _system: ExcelDateSystem,
) -> Result<Vec<i32>, ErrorKind> {
    fn push_value(
        ctx: &dyn FunctionContext,
        out: &mut Vec<i32>,
        v: Value,
    ) -> Result<(), ErrorKind> {
        match v {
            Value::Blank => Ok(()),
            Value::Error(e) => Err(e),
            Value::Array(arr) => {
                for el in arr.values {
                    push_value(ctx, out, el)?;
                }
                Ok(())
            }
            other => {
                out.push(coerce_to_serial_floor(ctx, &other)?);
                Ok(())
            }
        }
    }

    let mut out = Vec::new();
    match ctx.eval_arg(expr) {
        ArgValue::Scalar(v) => push_value(ctx, &mut out, v)?,
        ArgValue::Reference(r) => {
            for addr in ctx.iter_reference_cells(&r) {
                let v = ctx.get_cell_value(&r.sheet_id, addr);
                push_value(ctx, &mut out, v)?;
            }
        }
        ArgValue::ReferenceUnion(ranges) => {
            for r in ranges {
                for addr in ctx.iter_reference_cells(&r) {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    push_value(ctx, &mut out, v)?;
                }
            }
        }
    }
    Ok(out)
}

fn parse_weekend_mask(ctx: &dyn FunctionContext, v: &Value) -> Result<u8, ErrorKind> {
    // Excel defaults to Saturday/Sunday when the weekend argument is omitted or blank.
    if matches!(v, Value::Blank) {
        return Ok((1 << 5) | (1 << 6));
    }

    if let Value::Text(s) = v {
        let s = s.trim();
        if s.is_empty() {
            return Ok((1 << 5) | (1 << 6));
        }
        // Weekend strings are a 7-char mask (Mon..Sun) using '0' and '1'. If the text isn't
        // a valid mask, fall back to numeric coercion (e.g. "1" -> weekend code 1).
        let chars: Vec<char> = s.chars().collect();
        if chars.len() == 7 {
            let mut mask: u8 = 0;
            for (idx, ch) in chars.into_iter().enumerate() {
                match ch {
                    '0' => {}
                    '1' => mask |= 1 << idx,
                    _ => return Err(ErrorKind::Value),
                }
            }
            if mask == 0b111_1111 {
                return Err(ErrorKind::Num);
            }
            return Ok(mask);
        }
    }

    let code = coerce_to_i32_trunc(ctx, v)?;
    let mask = match code {
        1 => (1 << 5) | (1 << 6),
        2 => (1 << 6) | (1 << 0),
        3 => (1 << 0) | (1 << 1),
        4 => (1 << 1) | (1 << 2),
        5 => (1 << 2) | (1 << 3),
        6 => (1 << 3) | (1 << 4),
        7 => (1 << 4) | (1 << 5),
        11 => 1 << 6,
        12 => 1 << 0,
        13 => 1 << 1,
        14 => 1 << 2,
        15 => 1 << 3,
        16 => 1 << 4,
        17 => 1 << 5,
        _ => return Err(ErrorKind::Num),
    };
    if mask == 0b111_1111 {
        return Err(ErrorKind::Num);
    }
    Ok(mask)
}

fn normalize_year_month(year: i64, month: i64) -> (i64, i64) {
    let total_months = year * 12 + (month - 1);
    let new_year = total_months.div_euclid(12);
    let new_month = total_months.rem_euclid(12) + 1;
    (new_year, new_month)
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
