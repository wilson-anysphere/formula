use crate::error::ExcelError;
use crate::eval::CompiledExpr;
use crate::functions::array_lift;
use crate::functions::{ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

fn excel_error_kind(err: ExcelError) -> ErrorKind {
    match err {
        ExcelError::Div0 => ErrorKind::Div0,
        ExcelError::Value => ErrorKind::Value,
        ExcelError::Num => ErrorKind::Num,
    }
}

inventory::submit! {
    FunctionSpec {
        name: "ROMAN",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: roman_fn,
    }
}

fn roman_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = array_lift::eval_arg(ctx, &args[0]);
    let form = if args.len() == 2 {
        array_lift::eval_arg(ctx, &args[1])
    } else {
        Value::Number(0.0)
    };

    array_lift::lift2(number, form, |number, form| {
        let number = number.coerce_to_number_with_ctx(ctx)?;
        let form = form.coerce_to_number_with_ctx(ctx)?;

        if !number.is_finite() || !form.is_finite() {
            return Err(ErrorKind::Num);
        }

        let number = number.trunc() as i64;
        let form = form.trunc() as i64;

        match crate::functions::math::roman(number, Some(form)) {
            Ok(s) => Ok(Value::Text(s)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "ARABIC",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: arabic_fn,
    }
}

fn arabic_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(text, |text| {
        let text = text.coerce_to_string_with_ctx(ctx)?;
        match crate::functions::math::arabic(&text) {
            Ok(n) => Ok(Value::Number(n as f64)),
            Err(e) => Err(excel_error_kind(e)),
        }
    })
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
