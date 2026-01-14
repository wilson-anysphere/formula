use crate::eval::CompiledExpr;
use crate::functions::engineering::{self, FixedBase, BIT_MAX};
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

// ------------------------------------------------------------------
// Fixed-width base conversion functions (BIFF FTAB)
// ------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "BIN2DEC",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: bin2dec_fn,
    }
}

fn bin2dec_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    fixed_base_to_decimal_fn(ctx, args, FixedBase::Bin)
}

inventory::submit! {
    FunctionSpec {
        name: "BIN2OCT",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: bin2oct_fn,
    }
}

fn bin2oct_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    fixed_base_to_fixed_base_fn(ctx, args, FixedBase::Bin, FixedBase::Oct)
}

inventory::submit! {
    FunctionSpec {
        name: "BIN2HEX",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: bin2hex_fn,
    }
}

fn bin2hex_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    fixed_base_to_fixed_base_fn(ctx, args, FixedBase::Bin, FixedBase::Hex)
}

inventory::submit! {
    FunctionSpec {
        name: "OCT2DEC",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: oct2dec_fn,
    }
}

fn oct2dec_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    fixed_base_to_decimal_fn(ctx, args, FixedBase::Oct)
}

inventory::submit! {
    FunctionSpec {
        name: "OCT2BIN",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: oct2bin_fn,
    }
}

fn oct2bin_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    fixed_base_to_fixed_base_fn(ctx, args, FixedBase::Oct, FixedBase::Bin)
}

inventory::submit! {
    FunctionSpec {
        name: "OCT2HEX",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: oct2hex_fn,
    }
}

fn oct2hex_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    fixed_base_to_fixed_base_fn(ctx, args, FixedBase::Oct, FixedBase::Hex)
}

inventory::submit! {
    FunctionSpec {
        name: "HEX2DEC",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: hex2dec_fn,
    }
}

fn hex2dec_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    fixed_base_to_decimal_fn(ctx, args, FixedBase::Hex)
}

inventory::submit! {
    FunctionSpec {
        name: "HEX2BIN",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: hex2bin_fn,
    }
}

fn hex2bin_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    fixed_base_to_fixed_base_fn(ctx, args, FixedBase::Hex, FixedBase::Bin)
}

inventory::submit! {
    FunctionSpec {
        name: "HEX2OCT",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: hex2oct_fn,
    }
}

fn hex2oct_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    fixed_base_to_fixed_base_fn(ctx, args, FixedBase::Hex, FixedBase::Oct)
}

inventory::submit! {
    FunctionSpec {
        name: "DEC2BIN",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: dec2bin_fn,
    }
}

fn dec2bin_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    decimal_to_fixed_base_fn(ctx, args, FixedBase::Bin)
}

inventory::submit! {
    FunctionSpec {
        name: "DEC2OCT",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: dec2oct_fn,
    }
}

fn dec2oct_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    decimal_to_fixed_base_fn(ctx, args, FixedBase::Oct)
}

inventory::submit! {
    FunctionSpec {
        name: "DEC2HEX",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: dec2hex_fn,
    }
}

fn dec2hex_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    decimal_to_fixed_base_fn(ctx, args, FixedBase::Hex)
}

fn fixed_base_to_decimal_fn(
    ctx: &dyn FunctionContext,
    args: &[CompiledExpr],
    base: FixedBase,
) -> Value {
    let text = match eval_text(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match engineering::fixed_base_to_decimal(&text, base) {
        Ok(v) => Value::Number(v as f64),
        Err(e) => Value::Error(e),
    }
}

fn fixed_base_to_fixed_base_fn(
    ctx: &dyn FunctionContext,
    args: &[CompiledExpr],
    src: FixedBase,
    dst: FixedBase,
) -> Value {
    let text = match eval_text(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let places = if args.len() == 2 {
        match eval_places(ctx, &args[1]) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    match engineering::fixed_base_to_fixed_base(&text, src, dst, places) {
        Ok(v) => Value::Text(v),
        Err(e) => Value::Error(e),
    }
}

fn decimal_to_fixed_base_fn(
    ctx: &dyn FunctionContext,
    args: &[CompiledExpr],
    dst: FixedBase,
) -> Value {
    let number = match eval_i64_trunc(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let places = if args.len() == 2 {
        match eval_places(ctx, &args[1]) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    match engineering::fixed_decimal_to_fixed_base(number, dst, places) {
        Ok(v) => Value::Text(v),
        Err(e) => Value::Error(e),
    }
}

// ------------------------------------------------------------------
// Modern base conversion functions
// ------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "BASE",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Number, ValueType::Number, ValueType::Number],
        implementation: base_fn,
    }
}

fn base_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = match eval_u64_trunc_nonnegative(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let radix = match eval_u32_trunc(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let min_length = if args.len() == 3 {
        match eval_usize_trunc_nonnegative(ctx, &args[2]) {
            Ok(v) => Some(v),
            Err(e) => return Value::Error(e),
        }
    } else {
        None
    };

    match engineering::base_from_decimal(number, radix, min_length) {
        Ok(v) => Value::Text(v),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "DECIMAL",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: decimal_fn,
    }
}

fn decimal_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = match eval_text(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let radix = match eval_u32_trunc(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    match engineering::decimal_from_text(&text, radix) {
        Ok(v) => Value::Number(v),
        Err(e) => Value::Error(e),
    }
}

// ------------------------------------------------------------------
// BIT* functions
// ------------------------------------------------------------------

inventory::submit! {
    FunctionSpec {
        name: "BITAND",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: bitand_fn,
    }
}

fn bitand_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    bit_binary_op(ctx, args, engineering::bitand)
}

inventory::submit! {
    FunctionSpec {
        name: "BITOR",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: bitor_fn,
    }
}

fn bitor_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    bit_binary_op(ctx, args, engineering::bitor)
}

inventory::submit! {
    FunctionSpec {
        name: "BITXOR",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: bitxor_fn,
    }
}

fn bitxor_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    bit_binary_op(ctx, args, engineering::bitxor)
}

fn bit_binary_op(
    ctx: &dyn FunctionContext,
    args: &[CompiledExpr],
    op: fn(u64, u64) -> u64,
) -> Value {
    let a = match eval_bit_u64(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let b = match eval_bit_u64(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    Value::Number(op(a, b) as f64)
}

inventory::submit! {
    FunctionSpec {
        name: "BITLSHIFT",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: bitlshift_fn,
    }
}

fn bitlshift_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    bit_shift_op(ctx, args, engineering::bitlshift)
}

inventory::submit! {
    FunctionSpec {
        name: "BITRSHIFT",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: bitrshift_fn,
    }
}

fn bitrshift_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    bit_shift_op(ctx, args, engineering::bitrshift)
}

fn bit_shift_op(
    ctx: &dyn FunctionContext,
    args: &[CompiledExpr],
    op: fn(u64, i32) -> Result<u64, ErrorKind>,
) -> Value {
    let value = match eval_bit_u64(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let shift = match eval_i32_exact(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    match op(value, shift) {
        Ok(v) => Value::Number(v as f64),
        Err(e) => Value::Error(e),
    }
}

// ------------------------------------------------------------------
// Argument helpers
// ------------------------------------------------------------------

fn eval_text(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<String, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    v.coerce_to_string()
}

fn eval_places(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<usize, ErrorKind> {
    let n = eval_finite_number(ctx, expr)?;
    let t = n.trunc();
    if t < 0.0 || t > (usize::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(t as usize)
}

fn eval_finite_number(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<f64, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    let n = v.coerce_to_number_with_ctx(ctx)?;
    if n.is_finite() {
        Ok(n)
    } else {
        Err(ErrorKind::Num)
    }
}

fn eval_i64_trunc(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<i64, ErrorKind> {
    let n = eval_finite_number(ctx, expr)?;
    let t = n.trunc();
    if t < (i64::MIN as f64) || t > (i64::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(t as i64)
}

fn eval_u64_trunc_nonnegative(
    ctx: &dyn FunctionContext,
    expr: &CompiledExpr,
) -> Result<u64, ErrorKind> {
    let n = eval_finite_number(ctx, expr)?;
    let t = n.trunc();
    if t < 0.0 {
        return Err(ErrorKind::Num);
    }
    // Restrict to Excel's exact-integer domain for BASE/DECIMAL parity.
    const MAX_SAFE_INT: f64 = ((1u64 << 53) - 1) as f64;
    if t > MAX_SAFE_INT {
        return Err(ErrorKind::Num);
    }
    Ok(t as u64)
}

fn eval_u32_trunc(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<u32, ErrorKind> {
    let n = eval_finite_number(ctx, expr)?;
    let t = n.trunc();
    if t < 0.0 || t > (u32::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(t as u32)
}

fn eval_usize_trunc_nonnegative(
    ctx: &dyn FunctionContext,
    expr: &CompiledExpr,
) -> Result<usize, ErrorKind> {
    let n = eval_finite_number(ctx, expr)?;
    let t = n.trunc();
    if t < 0.0 || t > (usize::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(t as usize)
}

fn eval_i32_exact(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<i32, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    let n = v.coerce_to_number_with_ctx(ctx)?;
    if !n.is_finite() {
        return Err(ErrorKind::Num);
    }
    if n.trunc() != n {
        return Err(ErrorKind::Num);
    }
    if n < (i32::MIN as f64) || n > (i32::MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(n as i32)
}

fn eval_bit_u64(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<u64, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    let n = v.coerce_to_number_with_ctx(ctx)?;
    if !n.is_finite() {
        return Err(ErrorKind::Num);
    }
    if n.trunc() != n {
        return Err(ErrorKind::Num);
    }
    if n < 0.0 || n > (BIT_MAX as f64) {
        return Err(ErrorKind::Num);
    }
    Ok(n as u64)
}
