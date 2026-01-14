use crate::error::ExcelError;
use crate::eval::CompiledExpr;
use crate::functions::array_lift;
use crate::functions::{ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

inventory::submit! {
    FunctionSpec {
        name: "BAHTTEXT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Number],
        implementation: bahttext_fn,
    }
}

fn bahttext_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(number, |v| {
        let number = coerce_finite_number(ctx, v)?;
        match crate::functions::text::thai::bahttext(number) {
            Ok(s) => Ok(Value::Text(s)),
            Err(e) => Err(excel_error_to_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "THAIDAYOFWEEK",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Number],
        implementation: thaidayofweek_fn,
    }
}

fn thaidayofweek_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let serial = array_lift::eval_arg(ctx, &args[0]);
    let system = ctx.date_system();
    array_lift::lift1(serial, |v| {
        let serial = coerce_serial_floor(ctx, v)?;
        match crate::functions::date_time::thai::thaidayofweek(serial, system) {
            Ok(name) => Ok(Value::Text(name.to_string())),
            Err(e) => Err(excel_error_to_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "THAIDIGIT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: thaidigit_fn,
    }
}

fn thaidigit_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(text, |v| {
        let text = v.coerce_to_string_with_ctx(ctx)?;
        Ok(Value::Text(crate::functions::text::thai::thai_digit(&text)))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "THAIMONTHOFYEAR",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Number],
        implementation: thaimonthofyear_fn,
    }
}

fn thaimonthofyear_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let serial = array_lift::eval_arg(ctx, &args[0]);
    let system = ctx.date_system();
    array_lift::lift1(serial, |v| {
        let serial = coerce_serial_floor(ctx, v)?;
        match crate::functions::date_time::thai::thaimonthofyear(serial, system) {
            Ok(name) => Ok(Value::Text(name.to_string())),
            Err(e) => Err(excel_error_to_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "THAINUMSOUND",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Number],
        implementation: thainumsound_fn,
    }
}

fn thainumsound_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(number, |v| {
        let number = coerce_finite_number(ctx, v)?;
        match crate::functions::text::thai::thainumsound(number) {
            Ok(s) => Ok(Value::Text(s)),
            Err(e) => Err(excel_error_to_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "THAINUMSTRING",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Number],
        implementation: thainumstring_fn,
    }
}

fn thainumstring_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(number, |v| {
        let number = coerce_finite_number(ctx, v)?;
        match crate::functions::text::thai::thainumstring(number) {
            Ok(s) => Ok(Value::Text(s)),
            Err(e) => Err(excel_error_to_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "THAISTRINGLENGTH",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: thaistringlength_fn,
    }
}

fn thaistringlength_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(text, |v| {
        let text = v.coerce_to_string_with_ctx(ctx)?;
        Ok(Value::Number(
            crate::functions::text::thai::thai_string_length(&text) as f64,
        ))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "THAIYEAR",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: thaiyear_fn,
    }
}

fn thaiyear_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let serial = array_lift::eval_arg(ctx, &args[0]);
    let system = ctx.date_system();
    array_lift::lift1(serial, |v| {
        let serial = coerce_serial_floor(ctx, v)?;
        match crate::functions::date_time::thai::thaiyear(serial, system) {
            Ok(year) => Ok(Value::Number(year as f64)),
            Err(e) => Err(excel_error_to_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "ISTHAIDIGIT",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Text],
        implementation: isthaidigit_fn,
    }
}

fn isthaidigit_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(text, |v| {
        let text = v.coerce_to_string_with_ctx(ctx)?;
        Ok(Value::Bool(crate::functions::text::thai::is_thai_digit(
            &text,
        )))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "ROUNDBAHTDOWN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: roundbahtdown_fn,
    }
}

fn roundbahtdown_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(number, |v| {
        let number = coerce_finite_number(ctx, v)?;
        match crate::functions::text::thai::roundbahtdown(number) {
            Ok(n) => Ok(Value::Number(n)),
            Err(e) => Err(excel_error_to_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "ROUNDBAHTUP",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number],
        implementation: roundbahtup_fn,
    }
}

fn roundbahtup_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(number, |v| {
        let number = coerce_finite_number(ctx, v)?;
        match crate::functions::text::thai::roundbahtup(number) {
            Ok(n) => Ok(Value::Number(n)),
            Err(e) => Err(excel_error_to_kind(e)),
        }
    })
}

fn coerce_finite_number(ctx: &dyn FunctionContext, v: &Value) -> Result<f64, ErrorKind> {
    let n = v.coerce_to_number_with_ctx(ctx)?;
    if !n.is_finite() {
        return Err(ErrorKind::Num);
    }
    Ok(n)
}

fn coerce_serial_floor(ctx: &dyn FunctionContext, v: &Value) -> Result<i32, ErrorKind> {
    let n = coerce_finite_number(ctx, v)?;
    let serial = n.floor();
    if serial < (i32::MIN as f64) || serial > (i32::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(serial as i32)
}

fn excel_error_to_kind(err: ExcelError) -> ErrorKind {
    match err {
        ExcelError::Div0 => ErrorKind::Div0,
        ExcelError::Value => ErrorKind::Value,
        ExcelError::Num => ErrorKind::Num,
    }
}
