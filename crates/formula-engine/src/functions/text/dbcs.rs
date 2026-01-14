//! Legacy DBCS / byte-count text functions.
//!
//! Excel exposes `*B` variants of several text functions (LENB, LEFTB, MIDB, RIGHTB,
//! FINDB, SEARCHB, REPLACEB). In DBCS locales (e.g. Japanese), these functions
//! operate on *byte counts* instead of character counts, and the definition of a
//! "byte" depends on the active workbook locale / code page.
//!
//! The formula engine currently assumes an en-US workbook locale and Unicode
//! strings. Under that single-byte locale, the `*B` functions behave identically
//! to their non-`B` equivalents.
//!
//! `ASC` / `DBCS` perform half-width / full-width conversions in Japanese locales.
//! We currently implement these as identity transforms (no conversions).
//!
//! `PHONETIC` depends on per-cell phonetic guide metadata (furigana).
//! When phonetic metadata is present for a referenced cell, `PHONETIC(reference)`
//! returns that stored string. When phonetic metadata is absent (the common
//! case), Excel falls back to the referenced cell’s displayed text, so the
//! engine returns the referenced value coerced to text using the current
//! locale-aware formatting rules.
//!
//! Once workbook locale + codepage + phonetic metadata are modeled, this module
//! can be extended to implement real Excel semantics for DBCS workbooks.

use crate::eval::CompiledExpr;
use crate::eval::MAX_MATERIALIZED_ARRAY_CELLS;
use crate::functions::array_lift;
use crate::functions::{call_function, ArgValue, FunctionContext, Reference};
use crate::value::{Array, ErrorKind, Value};

pub(crate) fn findb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "FIND", args)
}

pub(crate) fn searchb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "SEARCH", args)
}

pub(crate) fn replaceb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "REPLACE", args)
}

pub(crate) fn leftb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "LEFT", args)
}

pub(crate) fn rightb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "RIGHT", args)
}

pub(crate) fn midb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "MID", args)
}

pub(crate) fn lenb_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // en-US: byte counts match character counts.
    call_function(ctx, "LEN", args)
}

pub(crate) fn asc_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(text, |text| {
        Ok(Value::Text(text.coerce_to_string_with_ctx(ctx)?))
    })
}

pub(crate) fn dbcs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(text, |text| {
        Ok(Value::Text(text.coerce_to_string_with_ctx(ctx)?))
    })
}

pub(crate) fn phonetic_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    match ctx.eval_arg(&args[0]) {
        ArgValue::Reference(reference) => phonetic_from_reference(ctx, reference),
        // TODO: Verify Excel's behavior for scalar/non-reference arguments (e.g. `PHONETIC("abc")`).
        // Historically, the engine treated PHONETIC as a string-coercion placeholder; preserve that
        // behavior until we have an Excel oracle case for scalar arguments.
        ArgValue::Scalar(value) => array_lift::lift1(value, |v| {
            Ok(Value::Text(v.coerce_to_string_with_ctx(ctx)?))
        }),
        ArgValue::ReferenceUnion(_) => Value::Error(ErrorKind::Value),
    }
}

fn phonetic_from_reference(ctx: &dyn FunctionContext, reference: Reference) -> Value {
    let reference = reference.normalized();
    ctx.record_reference(&reference);

    if reference.is_single_cell() {
        let cell_value = ctx.get_cell_value(&reference.sheet_id, reference.start);
        if let Value::Error(e) = &cell_value {
            return Value::Error(*e);
        }
        if let Some(phonetic) = ctx.get_cell_phonetic(&reference.sheet_id, reference.start) {
            return Value::Text(phonetic.to_string());
        }
        return match cell_value.coerce_to_string_with_ctx(ctx) {
            Ok(s) => Value::Text(s),
            Err(e) => Value::Error(e),
        };
    }

    // Preserve the existing array/broadcast behavior for multi-cell references.
    let rows = (reference.end.row - reference.start.row + 1) as usize;
    let cols = (reference.end.col - reference.start.col + 1) as usize;
    let total = match rows.checked_mul(cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if total > MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut out = Vec::new();
    if out.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
    for addr in reference.iter_cells() {
        let cell_value = ctx.get_cell_value(&reference.sheet_id, addr);
        // Error values are preserved per element (matching `array_lift` behavior).
        if let Value::Error(e) = cell_value {
            out.push(Value::Error(e));
            continue;
        }
        if let Some(phonetic) = ctx.get_cell_phonetic(&reference.sheet_id, addr) {
            out.push(Value::Text(phonetic.to_string()));
            continue;
        }
        out.push(match cell_value.coerce_to_string_with_ctx(ctx) {
            Ok(s) => Value::Text(s),
            Err(e) => Value::Error(e),
        });
    }
    Value::Array(Array::new(rows, cols, out))
}
