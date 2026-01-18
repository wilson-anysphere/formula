use super::builtins_helpers::excel_result_number;
use crate::eval::MAX_MATERIALIZED_ARRAY_CELLS;
use crate::eval::CompiledExpr;
use crate::functions::{ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

const VAR_ARGS: usize = 255;

fn eval_number_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<f64, ErrorKind> {
    let v = ctx.eval_scalar(expr);
    match v {
        Value::Error(e) => Err(e),
        other => other.coerce_to_number_with_ctx(ctx),
    }
}

fn eval_optional_number_arg(
    ctx: &dyn FunctionContext,
    expr: Option<&CompiledExpr>,
) -> Result<Option<f64>, ErrorKind> {
    match expr {
        Some(e) => Ok(Some(eval_number_arg(ctx, e)?)),
        None => Ok(None),
    }
}

fn collect_npv_values_from_arg(
    ctx: &dyn FunctionContext,
    arg: &CompiledExpr,
) -> Result<Vec<f64>, ErrorKind> {
    // NPV's discounting depends on the position of each cashflow. When values are
    // supplied via references, Excel ignores non-numeric cells. For NPV we treat
    // them as 0 so the period index is preserved (i.e. a blank cell represents a
    // period with no cashflow rather than "removing" a period).
    match ctx.eval_arg(arg) {
        ArgValue::Scalar(v) => match v {
            Value::Error(e) => Err(e),
            Value::Number(n) => Ok(vec![n]),
            Value::Bool(b) => Ok(vec![if b { 1.0 } else { 0.0 }]),
            Value::Blank => Ok(vec![0.0]),
            Value::Text(s) => {
                let trimmed = s.trim();
                if trimmed.is_empty() {
                    return Ok(vec![0.0]);
                }
                match crate::coercion::datetime::parse_value_text(
                    trimmed,
                    ctx.value_locale(),
                    ctx.now_utc(),
                    ctx.date_system(),
                ) {
                    Ok(n) => Ok(vec![n]),
                    Err(crate::error::ExcelError::Value) => Ok(vec![0.0]),
                    Err(crate::error::ExcelError::Div0) => Err(ErrorKind::Div0),
                    Err(crate::error::ExcelError::Num) => Err(ErrorKind::Num),
                }
            }
            Value::Entity(_) | Value::Record(_) => Ok(vec![0.0]),
            Value::Reference(_) | Value::ReferenceUnion(_) => Err(ErrorKind::Value),
            Value::Array(arr) => {
                let total = arr.values.len();
                if total > MAX_MATERIALIZED_ARRAY_CELLS {
                    return Err(ErrorKind::Spill);
                }
                let mut out: Vec<f64> = Vec::new();
                if out.try_reserve_exact(total).is_err() {
                    debug_assert!(false, "NPV allocation failed (cells={total})");
                    return Err(ErrorKind::Num);
                }
                for v in arr.iter() {
                    match v {
                        Value::Error(e) => return Err(*e),
                        Value::Number(n) => out.push(*n),
                        Value::Lambda(_) => return Err(ErrorKind::Value),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Entity(_)
                        | Value::Record(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Spill { .. }
                        | Value::Reference(_)
                        | Value::ReferenceUnion(_) => out.push(0.0),
                    }
                }
                Ok(out)
            }
            Value::Lambda(_) => Err(ErrorKind::Value),
            Value::Spill { .. } => Ok(vec![0.0]),
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            ctx.record_reference(&r);
            let total = r.size() as usize;
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Err(ErrorKind::Spill);
            }
            let mut out: Vec<f64> = Vec::new();
            if out.try_reserve_exact(total).is_err() {
                debug_assert!(false, "NPV allocation failed (cells={total})");
                return Err(ErrorKind::Num);
            }
            for addr in r.iter_cells() {
                let v = ctx.get_cell_value(&r.sheet_id, addr);
                match v {
                    Value::Error(e) => return Err(e),
                    Value::Number(n) => out.push(n),
                    Value::Lambda(_) => return Err(ErrorKind::Value),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Blank
                    | Value::Array(_)
                    | Value::Spill { .. }
                    | Value::Reference(_)
                    | Value::ReferenceUnion(_) => out.push(0.0),
                }
            }
            Ok(out)
        }
        ArgValue::ReferenceUnion(ranges) => {
            let mut out = Vec::new();
            for r in ranges {
                let r = r.normalized();
                ctx.record_reference(&r);
                let reserve = r.size() as usize;
                if out.len().saturating_add(reserve) > MAX_MATERIALIZED_ARRAY_CELLS {
                    return Err(ErrorKind::Spill);
                }
                if out.try_reserve(reserve).is_err() {
                    debug_assert!(false, "NPV allocation failed (reserve={reserve})");
                    return Err(ErrorKind::Num);
                }
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Err(e),
                        Value::Number(n) => out.push(n),
                        Value::Lambda(_) => return Err(ErrorKind::Value),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Entity(_)
                        | Value::Record(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Spill { .. }
                        | Value::Reference(_)
                        | Value::ReferenceUnion(_) => out.push(0.0),
                    }
                }
            }
            Ok(out)
        }
    }
}

fn collect_irr_values_from_arg(
    ctx: &dyn FunctionContext,
    arg: &CompiledExpr,
) -> Result<Vec<f64>, ErrorKind> {
    // Excel IRR accepts references that contain non-numeric cells; text/logical/blank
    // entries are ignored (i.e. contribute 0) but their position still counts as a period.
    // Errors propagate.
    match ctx.eval_arg(arg) {
        ArgValue::Scalar(v) => match v {
            Value::Error(e) => Err(e),
            Value::Number(n) => Ok(vec![n]),
            Value::Bool(_)
            | Value::Text(_)
            | Value::Entity(_)
            | Value::Record(_)
            | Value::Blank => Ok(vec![0.0]),
            Value::Lambda(_) => Err(ErrorKind::Value),
            Value::Reference(_) | Value::ReferenceUnion(_) => Err(ErrorKind::Value),
            Value::Array(arr) => {
                let total = arr.values.len();
                if total > MAX_MATERIALIZED_ARRAY_CELLS {
                    return Err(ErrorKind::Spill);
                }
                let mut out: Vec<f64> = Vec::new();
                if out.try_reserve_exact(total).is_err() {
                    debug_assert!(false, "IRR allocation failed (cells={total})");
                    return Err(ErrorKind::Num);
                }
                for v in arr.iter() {
                    match v {
                        Value::Error(e) => return Err(*e),
                        Value::Number(n) => out.push(*n),
                        Value::Lambda(_) => return Err(ErrorKind::Value),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Entity(_)
                        | Value::Record(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Spill { .. }
                        | Value::Reference(_)
                        | Value::ReferenceUnion(_) => out.push(0.0),
                    }
                }
                Ok(out)
            }
            Value::Spill { .. } => Ok(vec![0.0]),
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            ctx.record_reference(&r);
            let total = r.size() as usize;
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Err(ErrorKind::Spill);
            }
            let mut out: Vec<f64> = Vec::new();
            if out.try_reserve_exact(total).is_err() {
                debug_assert!(false, "IRR allocation failed (cells={total})");
                return Err(ErrorKind::Num);
            }
            for addr in r.iter_cells() {
                let v = ctx.get_cell_value(&r.sheet_id, addr);
                match v {
                    Value::Error(e) => return Err(e),
                    Value::Number(n) => out.push(n),
                    Value::Lambda(_) => return Err(ErrorKind::Value),
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Entity(_)
                    | Value::Record(_)
                    | Value::Blank
                    | Value::Array(_)
                    | Value::Spill { .. }
                    | Value::Reference(_)
                    | Value::ReferenceUnion(_) => out.push(0.0),
                }
            }
            Ok(out)
        }
        ArgValue::ReferenceUnion(ranges) => {
            let mut out = Vec::new();
            for r in ranges {
                let r = r.normalized();
                ctx.record_reference(&r);
                let reserve = r.size() as usize;
                if out.len().saturating_add(reserve) > MAX_MATERIALIZED_ARRAY_CELLS {
                    return Err(ErrorKind::Spill);
                }
                if out.try_reserve(reserve).is_err() {
                    debug_assert!(false, "IRR allocation failed (reserve={reserve})");
                    return Err(ErrorKind::Num);
                }
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v {
                        Value::Error(e) => return Err(e),
                        Value::Number(n) => out.push(n),
                        Value::Lambda(_) => return Err(ErrorKind::Value),
                        Value::Bool(_)
                        | Value::Text(_)
                        | Value::Entity(_)
                        | Value::Record(_)
                        | Value::Blank
                        | Value::Array(_)
                        | Value::Spill { .. }
                        | Value::Reference(_)
                        | Value::ReferenceUnion(_) => out.push(0.0),
                    }
                }
            }
            Ok(out)
        }
    }
}

fn collect_numbers_strict_from_arg(
    ctx: &dyn FunctionContext,
    arg: &CompiledExpr,
) -> Result<Vec<f64>, ErrorKind> {
    match ctx.eval_arg(arg) {
        ArgValue::Scalar(v) => match v {
            Value::Array(arr) => {
                let total = arr.values.len();
                if total > MAX_MATERIALIZED_ARRAY_CELLS {
                    return Err(ErrorKind::Spill);
                }
                let mut out: Vec<f64> = Vec::new();
                if out.try_reserve_exact(total).is_err() {
                    debug_assert!(false, "financial allocation failed (cells={total})");
                    return Err(ErrorKind::Num);
                }
                for v in arr.iter() {
                    out.push(v.coerce_to_number_with_ctx(ctx)?);
                }
                Ok(out)
            }
            other => Ok(vec![other.coerce_to_number_with_ctx(ctx)?]),
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            ctx.record_reference(&r);
            let total = r.size() as usize;
            if total > MAX_MATERIALIZED_ARRAY_CELLS {
                return Err(ErrorKind::Spill);
            }
            let mut out: Vec<f64> = Vec::new();
            if out.try_reserve_exact(total).is_err() {
                debug_assert!(false, "financial allocation failed (cells={total})");
                return Err(ErrorKind::Num);
            }
            for addr in r.iter_cells() {
                let v = ctx.get_cell_value(&r.sheet_id, addr);
                out.push(v.coerce_to_number_with_ctx(ctx)?);
            }
            Ok(out)
        }
        ArgValue::ReferenceUnion(ranges) => {
            let mut out = Vec::new();
            for r in ranges {
                let r = r.normalized();
                ctx.record_reference(&r);
                let reserve = r.size() as usize;
                if out.len().saturating_add(reserve) > MAX_MATERIALIZED_ARRAY_CELLS {
                    return Err(ErrorKind::Spill);
                }
                if out.try_reserve(reserve).is_err() {
                    debug_assert!(false, "financial allocation failed (reserve={reserve})");
                    return Err(ErrorKind::Num);
                }
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    out.push(v.coerce_to_number_with_ctx(ctx)?);
                }
            }
            Ok(out)
        }
    }
}

inventory::submit! {
    FunctionSpec {
        name: "PV",
        min_args: 3,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: pv_fn,
    }
}

fn pv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rate = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let nper = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pmt = match eval_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let fv = match eval_optional_number_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let typ = match eval_optional_number_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::pv(rate, nper, pmt, fv, typ))
}

inventory::submit! {
    FunctionSpec {
        name: "FV",
        min_args: 3,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: fv_fn,
    }
}

fn fv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rate = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let nper = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pmt = match eval_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pv = match eval_optional_number_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let typ = match eval_optional_number_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::fv(rate, nper, pmt, pv, typ))
}

inventory::submit! {
    FunctionSpec {
        name: "PMT",
        min_args: 3,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: pmt_fn,
    }
}

fn pmt_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rate = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let nper = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pv = match eval_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let fv = match eval_optional_number_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let typ = match eval_optional_number_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::pmt(rate, nper, pv, fv, typ))
}

inventory::submit! {
    FunctionSpec {
        name: "NPER",
        min_args: 3,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: nper_fn,
    }
}

fn nper_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rate = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pmt = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pv = match eval_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let fv = match eval_optional_number_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let typ = match eval_optional_number_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::nper(rate, pmt, pv, fv, typ))
}

inventory::submit! {
    FunctionSpec {
        name: "RATE",
        min_args: 3,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: rate_fn,
    }
}

fn rate_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let nper = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pmt = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pv = match eval_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let fv = match eval_optional_number_arg(ctx, args.get(3)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let typ = match eval_optional_number_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let guess = match eval_optional_number_arg(ctx, args.get(5)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::rate(nper, pmt, pv, fv, typ, guess))
}

inventory::submit! {
    FunctionSpec {
        name: "EFFECT",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: effect_fn,
    }
}

fn effect_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let nominal_rate = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let npery = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::effect(nominal_rate, npery))
}

inventory::submit! {
    FunctionSpec {
        name: "NOMINAL",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: nominal_fn,
    }
}

fn nominal_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let effect_rate = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let npery = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::nominal(effect_rate, npery))
}

inventory::submit! {
    FunctionSpec {
        name: "RRI",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: rri_fn,
    }
}

fn rri_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let nper = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pv = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let fv = match eval_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::rri(nper, pv, fv))
}

inventory::submit! {
    FunctionSpec {
        name: "IPMT",
        min_args: 4,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: ipmt_fn,
    }
}

fn ipmt_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rate = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let per = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let nper = match eval_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pv = match eval_number_arg(ctx, &args[3]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let fv = match eval_optional_number_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let typ = match eval_optional_number_arg(ctx, args.get(5)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::ipmt(rate, per, nper, pv, fv, typ))
}

inventory::submit! {
    FunctionSpec {
        name: "PPMT",
        min_args: 4,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: ppmt_fn,
    }
}

fn ppmt_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rate = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let per = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let nper = match eval_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let pv = match eval_number_arg(ctx, &args[3]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let fv = match eval_optional_number_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let typ = match eval_optional_number_arg(ctx, args.get(5)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::ppmt(rate, per, nper, pv, fv, typ))
}

inventory::submit! {
    FunctionSpec {
        name: "SLN",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: sln_fn,
    }
}

fn sln_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let cost = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let salvage = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let life = match eval_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::sln(cost, salvage, life))
}

inventory::submit! {
    FunctionSpec {
        name: "SYD",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: syd_fn,
    }
}

fn syd_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let cost = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let salvage = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let life = match eval_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let per = match eval_number_arg(ctx, &args[3]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::syd(cost, salvage, life, per))
}

inventory::submit! {
    FunctionSpec {
        name: "DDB",
        min_args: 4,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: ddb_fn,
    }
}

fn ddb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let cost = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let salvage = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let life = match eval_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let period = match eval_number_arg(ctx, &args[3]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let factor = match eval_optional_number_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::ddb(cost, salvage, life, period, factor))
}

inventory::submit! {
    FunctionSpec {
        name: "NPV",
        min_args: 2,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Any],
        implementation: npv_fn,
    }
}

fn npv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rate = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    let mut values = Vec::new();
    for arg in &args[1..] {
        match collect_npv_values_from_arg(ctx, arg) {
            Ok(mut nums) => values.append(&mut nums),
            Err(e) => return Value::Error(e),
        }
    }

    excel_result_number(super::npv(rate, &values))
}

inventory::submit! {
    FunctionSpec {
        name: "IRR",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number],
        implementation: irr_fn,
    }
}

fn irr_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_irr_values_from_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let guess = match eval_optional_number_arg(ctx, args.get(1)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::irr(&values, guess))
}

inventory::submit! {
    FunctionSpec {
        name: "MIRR",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Number, ValueType::Number],
        implementation: mirr_fn,
    }
}

fn mirr_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_irr_values_from_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let finance_rate = match eval_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let reinvest_rate = match eval_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::mirr(&values, finance_rate, reinvest_rate))
}

inventory::submit! {
    FunctionSpec {
        name: "XNPV",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Any, ValueType::Any],
        implementation: xnpv_fn,
    }
}

fn xnpv_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let rate = match eval_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    let values = match collect_numbers_strict_from_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let dates = match collect_numbers_strict_from_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::xnpv(rate, &values, &dates))
}

inventory::submit! {
    FunctionSpec {
        name: "XIRR",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Any, ValueType::Any, ValueType::Number],
        implementation: xirr_fn,
    }
}

fn xirr_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let values = match collect_numbers_strict_from_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let dates = match collect_numbers_strict_from_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let guess = match eval_optional_number_arg(ctx, args.get(2)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::xirr(&values, &dates, guess))
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function (indirectly via the
// `financial::__force_link` shim) ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
