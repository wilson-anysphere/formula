use crate::error::{ExcelError, ExcelResult};
use crate::eval::CompiledExpr;
use crate::functions::{ArraySupport, FunctionContext, FunctionSpec};
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

fn eval_finite_number_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<f64, ErrorKind> {
    let v = ctx.eval_scalar(expr);
    match v {
        Value::Error(e) => Err(e),
        other => {
            let n = other.coerce_to_number_with_ctx(ctx)?;
            if n.is_finite() {
                Ok(n)
            } else {
                Err(ErrorKind::Num)
            }
        }
    }
}

fn eval_optional_finite_number_arg(
    ctx: &dyn FunctionContext,
    expr: Option<&CompiledExpr>,
) -> Result<Option<f64>, ErrorKind> {
    match expr {
        Some(e) => Ok(Some(eval_finite_number_arg(ctx, e)?)),
        None => Ok(None),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DB",
        min_args: 4,
        max_args: 5,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: db_fn,
    }
}

fn db_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let cost = match eval_finite_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let salvage = match eval_finite_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let life = match eval_finite_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let period = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let month = match eval_optional_finite_number_arg(ctx, args.get(4)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::db(cost, salvage, life, period, month))
}

inventory::submit! {
    FunctionSpec {
        name: "VDB",
        min_args: 5,
        max_args: 7,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: vdb_fn,
    }
}

fn vdb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let cost = match eval_finite_number_arg(ctx, &args[0]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let salvage = match eval_finite_number_arg(ctx, &args[1]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let life = match eval_finite_number_arg(ctx, &args[2]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let start = match eval_finite_number_arg(ctx, &args[3]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let end = match eval_finite_number_arg(ctx, &args[4]) {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let factor = match eval_optional_finite_number_arg(ctx, args.get(5)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let no_switch = match eval_optional_finite_number_arg(ctx, args.get(6)) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    excel_result_number(super::vdb(cost, salvage, life, start, end, factor, no_switch))
}

