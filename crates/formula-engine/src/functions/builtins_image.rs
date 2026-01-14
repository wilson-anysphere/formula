use crate::eval::CompiledExpr;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, RecordValue, Value};

inventory::submit! {
    FunctionSpec {
        name: "IMAGE",
        min_args: 1,
        max_args: 5,
        volatility: Volatility::Volatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Any,
        arg_types: &[
            ValueType::Text,
            ValueType::Text,
            ValueType::Number,
            ValueType::Number,
            ValueType::Number,
        ],
        implementation: image_fn,
    }
}

fn image_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // IMAGE(source, [alt_text], [sizing], [height], [width])
    //
    // The core engine does not fetch or decode images; we return a deterministic rich-value
    // record that preserves a stable display string and exposes standard fields (Excel-compatible
    // rich value behavior).
    let source = match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };

    let alt_text = if args.len() >= 2 {
        match eval_scalar_arg(ctx, &args[1]).coerce_to_string_with_ctx(ctx) {
            Ok(s) => Some(s),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    let sizing = if args.len() >= 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_i64_with_ctx(ctx) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        }
    } else {
        0
    };

    // Excel's documented sizing modes are 0..=3.
    if !(0..=3).contains(&sizing) {
        return Value::Error(ErrorKind::Value);
    }

    let height = if args.len() >= 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_number_with_ctx(ctx) {
            Ok(n) => Some(n),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    let width = if args.len() >= 5 {
        match eval_scalar_arg(ctx, &args[4]).coerce_to_number_with_ctx(ctx) {
            Ok(n) => Some(n),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    // Height/width are only validated/required in custom sizing mode.
    if sizing == 3 {
        let (Some(height), Some(width)) = (height, width) else {
            return Value::Error(ErrorKind::Value);
        };

        if !height.is_finite() || !width.is_finite() {
            return Value::Error(ErrorKind::Num);
        }

        // Excel rejects non-positive custom dimensions.
        if height <= 0.0 || width <= 0.0 {
            return Value::Error(ErrorKind::Num);
        }
    }

    let display_field = if alt_text.is_some() {
        "alt_text"
    } else {
        "source"
    };
    let display = alt_text.clone().unwrap_or_else(|| source.clone());
    let alt_value = alt_text.map(Value::Text).unwrap_or(Value::Blank);

    // Match Excel: `IMAGE` behaves like a rich value with field access.
    let mut record = RecordValue::with_fields_iter(
        display,
        [
            ("source", Value::Text(source)),
            ("alt_text", alt_value),
            ("sizing", Value::Number(sizing as f64)),
            ("height", height.map(Value::Number).unwrap_or(Value::Blank)),
            ("width", width.map(Value::Number).unwrap_or(Value::Blank)),
        ],
    );
    record.display_field = Some(display_field.to_string());

    Value::Record(record)
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
