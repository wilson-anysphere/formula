use crate::error::{ExcelError, ExcelResult};
use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

fn excel_result_number(res: ExcelResult<f64>) -> Value {
    match res {
        Ok(n) => Value::Number(n),
        Err(e) => Value::Error(match e {
            ExcelError::Div0 => ErrorKind::Div0,
            ExcelError::Value => ErrorKind::Value,
            ExcelError::Num => ErrorKind::Num,
        }),
    }
}

fn collect_schedule_values_from_arg(
    ctx: &dyn FunctionContext,
    arg: &CompiledExpr,
) -> Result<Vec<f64>, ErrorKind> {
    match ctx.eval_arg(arg) {
        ArgValue::Scalar(v) => match v {
            Value::Array(arr) => {
                let mut out = Vec::with_capacity(arr.rows.saturating_mul(arr.cols));
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
            let mut out = Vec::new();
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
        name: "FVSCHEDULE",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
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

