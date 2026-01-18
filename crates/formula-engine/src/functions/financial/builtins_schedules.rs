use super::builtins_helpers::excel_result_number;
use crate::eval::MAX_MATERIALIZED_ARRAY_CELLS;
use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

fn collect_schedule_values_from_arg(
    ctx: &dyn FunctionContext,
    arg: &CompiledExpr,
) -> Result<Vec<f64>, ErrorKind> {
    match ctx.eval_arg(arg) {
        ArgValue::Scalar(v) => match v {
            Value::Error(e) => Err(e),
            Value::Array(arr) => {
                let total = arr.values.len();
                if total > MAX_MATERIALIZED_ARRAY_CELLS {
                    return Err(ErrorKind::Spill);
                }
                let mut out: Vec<f64> = Vec::new();
                if out.try_reserve_exact(total).is_err() {
                    debug_assert!(false, "FVSCHEDULE allocation failed (cells={total})");
                    return Err(ErrorKind::Num);
                }
                for v in arr.iter() {
                    out.push(v.coerce_to_number_with_ctx(ctx)?);
                }
                Ok(out)
            }
            Value::Lambda(_) => Err(ErrorKind::Value),
            Value::Reference(_) | Value::ReferenceUnion(_) => Err(ErrorKind::Value),
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
                debug_assert!(false, "FVSCHEDULE allocation failed (cells={total})");
                return Err(ErrorKind::Num);
            }
            for addr in r.iter_cells() {
                let v = ctx.get_cell_value(&r.sheet_id, addr);
                match v {
                    Value::Error(e) => return Err(e),
                    Value::Number(n) => out.push(n),
                    Value::Lambda(_) => return Err(ErrorKind::Value),
                    // Excel: within a reference/range, blanks/text/logicals do not participate.
                    // For FVSCHEDULE, treating them as 0% yields the same result as ignoring them.
                    Value::Bool(_)
                    | Value::Text(_)
                    | Value::Blank
                    | Value::Entity(_)
                    | Value::Record(_)
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
                    debug_assert!(false, "FVSCHEDULE allocation failed (reserve={reserve})");
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
                        | Value::Blank
                        | Value::Entity(_)
                        | Value::Record(_)
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

inventory::submit! {
    FunctionSpec {
        name: "FVSCHEDULE",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Any],
        implementation: fvschedule_fn,
    }
}

fn fvschedule_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let principal = match eval_scalar_arg(ctx, &args[0]).coerce_to_number_with_ctx(ctx) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };

    let schedule = match collect_schedule_values_from_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::fvschedule(principal, &schedule))
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function (indirectly via
// `financial::__force_link`) ensures the module (and its `inventory::submit!` entries) are
// retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
