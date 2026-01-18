use std::collections::HashSet;

use crate::eval::{CompiledExpr, MAX_MATERIALIZED_ARRAY_CELLS};
use crate::functions::statistical::ets::{self, AggregationMethod};
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

fn collect_optional_numbers_from_arg(
    ctx: &dyn FunctionContext,
    expr: &CompiledExpr,
) -> Result<Vec<Option<f64>>, ErrorKind> {
    fn coerce_cell(ctx: &dyn FunctionContext, v: &Value) -> Result<Option<f64>, ErrorKind> {
        match v {
            Value::Error(e) => Err(*e),
            Value::Blank => Ok(None),
            Value::Number(n) => Ok(Some(*n)),
            Value::Bool(b) => Ok(Some(if *b { 1.0 } else { 0.0 })),
            Value::Text(_) | Value::Entity(_) | Value::Record(_) => {
                Ok(Some(v.coerce_to_number_with_ctx(ctx)?))
            }
            Value::Array(_)
            | Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => Err(ErrorKind::Value),
        }
    }

    match ctx.eval_arg(expr) {
        ArgValue::Scalar(v) => match &v {
            Value::Array(arr) => {
                let total = arr.values.len();
                if total > MAX_MATERIALIZED_ARRAY_CELLS {
                    return Err(ErrorKind::Spill);
                }
                let mut out: Vec<Option<f64>> = Vec::new();
                if out.try_reserve_exact(total).is_err() {
                    debug_assert!(false, "ETS allocation failed (cells={total})");
                    return Err(ErrorKind::Num);
                }
                for cell in arr.iter() {
                    out.push(coerce_cell(ctx, cell)?);
                }
                Ok(out)
            }
            _ => Ok(vec![coerce_cell(ctx, &v)?]),
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            ctx.record_reference(&r);
            let total = r.size() as usize;
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Err(ErrorKind::Spill);
            }
            let mut out: Vec<Option<f64>> = Vec::new();
            if out.try_reserve_exact(total).is_err() {
                debug_assert!(false, "ETS allocation failed (cells={total})");
                return Err(ErrorKind::Num);
            }
            for addr in r.iter_cells() {
                let v = ctx.get_cell_value(&r.sheet_id, addr);
                out.push(coerce_cell(ctx, &v)?);
            }
            Ok(out)
        }
        ArgValue::ReferenceUnion(ranges) => {
            let mut seen = HashSet::new();
            let mut out = Vec::new();
            for r in ranges {
                let r = r.normalized();
                ctx.record_reference(&r);
                let reserve = r.size() as usize;
                if out.len().saturating_add(reserve) > MAX_MATERIALIZED_ARRAY_CELLS {
                    return Err(ErrorKind::Spill);
                }
                if out.try_reserve(reserve).is_err() {
                    debug_assert!(false, "ETS allocation failed (reserve={reserve})");
                    return Err(ErrorKind::Num);
                }
                for addr in r.iter_cells() {
                    if !seen.insert((r.sheet_id.clone(), addr)) {
                        continue;
                    }
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    out.push(coerce_cell(ctx, &v)?);
                }
            }
            Ok(out)
        }
    }
}

fn collect_paired_series(
    ctx: &dyn FunctionContext,
    values_expr: &CompiledExpr,
    timeline_expr: &CompiledExpr,
) -> Result<(Vec<f64>, Vec<f64>), ErrorKind> {
    let values = collect_optional_numbers_from_arg(ctx, values_expr)?;
    let timeline = collect_optional_numbers_from_arg(ctx, timeline_expr)?;
    if values.len() != timeline.len() {
        return Err(ErrorKind::NA);
    }

    let mut out_values = Vec::new();
    let mut out_timeline = Vec::new();
    for (v, t) in values.into_iter().zip(timeline.into_iter()) {
        let (Some(v), Some(t)) = (v, t) else {
            continue;
        };
        out_values.push(v);
        out_timeline.push(t);
    }
    if out_values.len() < 2 {
        return Err(ErrorKind::Num);
    }
    Ok((out_values, out_timeline))
}

fn parse_optional_i64(
    ctx: &dyn FunctionContext,
    expr: Option<&CompiledExpr>,
) -> Result<Option<i64>, ErrorKind> {
    match expr {
        None => Ok(None),
        Some(e) => {
            let v = eval_scalar_arg(ctx, e);
            match v {
                Value::Blank => Ok(None),
                Value::Error(e) => Err(e),
                other => Ok(Some(other.coerce_to_i64_with_ctx(ctx)?)),
            }
        }
    }
}

fn parse_optional_number(
    ctx: &dyn FunctionContext,
    expr: Option<&CompiledExpr>,
) -> Result<Option<f64>, ErrorKind> {
    match expr {
        None => Ok(None),
        Some(e) => {
            let v = eval_scalar_arg(ctx, e);
            match v {
                Value::Blank => Ok(None),
                Value::Error(e) => Err(e),
                other => Ok(Some(other.coerce_to_number_with_ctx(ctx)?)),
            }
        }
    }
}

fn parse_data_completion(code: Option<i64>) -> Result<bool, ErrorKind> {
    match code.unwrap_or(1) {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(ErrorKind::Num),
    }
}

fn parse_aggregation(code: Option<i64>) -> Result<AggregationMethod, ErrorKind> {
    let code = code.unwrap_or(1);
    AggregationMethod::from_code(code)
}

fn parse_seasonality(code: Option<i64>) -> Result<Option<usize>, ErrorKind> {
    match code {
        None => Ok(None),
        Some(v) => {
            if v < 0 || v > 8760 {
                return Err(ErrorKind::Num);
            }
            Ok(Some(v as usize))
        }
    }
}

fn interpolate_observed(values: &[f64], pos: f64) -> Result<f64, ErrorKind> {
    if values.is_empty() || !pos.is_finite() {
        return Err(ErrorKind::Num);
    }
    if pos <= 0.0 {
        return Ok(values[0]);
    }
    let last = (values.len() - 1) as f64;
    if pos >= last {
        return Ok(values[values.len() - 1]);
    }

    let idx0 = pos.floor() as usize;
    let frac = pos - (idx0 as f64);
    if frac == 0.0 {
        return Ok(values[idx0]);
    }
    let a = values[idx0];
    let b = values[idx0 + 1];
    let out = a + frac * (b - a);
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

inventory::submit! {
    FunctionSpec {
        name: "FORECAST.ETS",
        min_args: 3,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: forecast_ets_fn,
    }
}

fn forecast_ets_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let target_date = match eval_scalar_arg(ctx, &args[0]).coerce_to_number_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let (values, timeline) = match collect_paired_series(ctx, &args[1], &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let seasonality_code = match parse_optional_i64(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let seasonality = match parse_seasonality(seasonality_code) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let data_completion_code = match parse_optional_i64(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let data_completion = match parse_data_completion(data_completion_code) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let aggregation_code = match parse_optional_i64(ctx, args.get(5)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let aggregation = match parse_aggregation(aggregation_code) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let series = match ets::prepare_series(
        &values,
        &timeline,
        data_completion,
        aggregation,
        ctx.date_system(),
    ) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };

    let seasonality = match seasonality {
        None | Some(0) => match ets::detect_seasonality(&series.values) {
            Ok(m) => m,
            Err(e) => return Value::Error(e),
        },
        Some(m) => m.max(1),
    };

    let fit = match ets::fit(&series.values, seasonality) {
        Ok(f) => f,
        Err(e) => return Value::Error(e),
    };

    let pos = match series.position(target_date) {
        Ok(p) => p,
        Err(e) => return Value::Error(e),
    };
    let last_pos = series.last_pos();
    let out = if pos <= last_pos {
        match interpolate_observed(&series.values, pos) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        match fit.forecast(pos - last_pos) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    };

    if out.is_finite() {
        Value::Number(out)
    } else {
        Value::Error(ErrorKind::Num)
    }
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}

inventory::submit! {
    FunctionSpec {
        name: "FORECAST.ETS.CONFINT",
        min_args: 3,
        max_args: 7,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: forecast_ets_confint_fn,
    }
}

fn forecast_ets_confint_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let target_date = match eval_scalar_arg(ctx, &args[0]).coerce_to_number_with_ctx(ctx) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let (values, timeline) = match collect_paired_series(ctx, &args[1], &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let confidence_level = match parse_optional_number(ctx, args.get(3)) {
        Ok(v) => v.unwrap_or(0.95),
        Err(e) => return Value::Error(e),
    };
    if !(0.0 < confidence_level && confidence_level < 1.0) {
        return Value::Error(ErrorKind::Num);
    }

    let seasonality_code = match parse_optional_i64(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let seasonality = match parse_seasonality(seasonality_code) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let data_completion_code = match parse_optional_i64(ctx, args.get(5)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let data_completion = match parse_data_completion(data_completion_code) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let aggregation_code = match parse_optional_i64(ctx, args.get(6)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let aggregation = match parse_aggregation(aggregation_code) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let series = match ets::prepare_series(
        &values,
        &timeline,
        data_completion,
        aggregation,
        ctx.date_system(),
    ) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };

    let seasonality = match seasonality {
        None | Some(0) => match ets::detect_seasonality(&series.values) {
            Ok(m) => m,
            Err(e) => return Value::Error(e),
        },
        Some(m) => m.max(1),
    };

    let fit = match ets::fit(&series.values, seasonality) {
        Ok(f) => f,
        Err(e) => return Value::Error(e),
    };

    let pos = match series.position(target_date) {
        Ok(p) => p,
        Err(e) => return Value::Error(e),
    };
    let last_pos = series.last_pos();
    let h = (pos - last_pos).max(0.0);
    if h == 0.0 || fit.rmse == 0.0 {
        return Value::Number(0.0);
    }

    let z = match ets::norm_s_inv((1.0 + confidence_level) / 2.0) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let out = z.abs() * fit.rmse * h.sqrt();
    if out.is_finite() {
        Value::Number(out)
    } else {
        Value::Error(ErrorKind::Num)
    }
}

inventory::submit! {
    FunctionSpec {
        name: "FORECAST.ETS.SEASONALITY",
        min_args: 2,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: forecast_ets_seasonality_fn,
    }
}

fn forecast_ets_seasonality_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let (values, timeline) = match collect_paired_series(ctx, &args[0], &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let data_completion_code = match parse_optional_i64(ctx, args.get(2)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let data_completion = match parse_data_completion(data_completion_code) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let aggregation_code = match parse_optional_i64(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let aggregation = match parse_aggregation(aggregation_code) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let series = match ets::prepare_series(
        &values,
        &timeline,
        data_completion,
        aggregation,
        ctx.date_system(),
    ) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };

    match ets::detect_seasonality(&series.values) {
        Ok(v) => Value::Number(v as f64),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "FORECAST.ETS.STAT",
        min_args: 2,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number, ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: forecast_ets_stat_fn,
    }
}

fn forecast_ets_stat_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let (values, timeline) = match collect_paired_series(ctx, &args[0], &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let seasonality_code = match parse_optional_i64(ctx, args.get(2)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let seasonality = match parse_seasonality(seasonality_code) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let data_completion_code = match parse_optional_i64(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let data_completion = match parse_data_completion(data_completion_code) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let aggregation_code = match parse_optional_i64(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let aggregation = match parse_aggregation(aggregation_code) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let statistic_type = match parse_optional_i64(ctx, args.get(5)) {
        Ok(v) => v.unwrap_or(1),
        Err(e) => return Value::Error(e),
    };

    if !(1..=8).contains(&statistic_type) {
        return Value::Error(ErrorKind::Num);
    }

    let series = match ets::prepare_series(
        &values,
        &timeline,
        data_completion,
        aggregation,
        ctx.date_system(),
    ) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };

    let seasonality = match seasonality {
        None | Some(0) => match ets::detect_seasonality(&series.values) {
            Ok(m) => m,
            Err(e) => return Value::Error(e),
        },
        Some(m) => m.max(1),
    };

    let fit = match ets::fit(&series.values, seasonality) {
        Ok(f) => f,
        Err(e) => return Value::Error(e),
    };

    let out = match statistic_type {
        1 => fit.alpha,
        2 => fit.beta,
        3 => fit.gamma,
        4 => fit.phi,
        5 => fit.mase,
        6 => fit.smape,
        7 => fit.mae,
        8 => fit.rmse,
        _ => return Value::Error(ErrorKind::Num),
    };

    if out.is_finite() {
        Value::Number(out)
    } else {
        Value::Error(ErrorKind::Num)
    }
}
