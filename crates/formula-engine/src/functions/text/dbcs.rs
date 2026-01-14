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
//! `PHONETIC` depends on per-cell phonetic guide metadata (furigana). When phonetic
//! metadata is available, the engine returns it; otherwise it falls back to
//! returning the referenced value coerced to text.
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
    // If the argument is a reference, attempt to read the target cells' phonetic guides.
    // Otherwise, this behaves like `TEXT(value)` (coerce to text with locale-aware number
    // formatting).
    match ctx.eval_arg(&args[0]) {
        ArgValue::Scalar(v) => {
            array_lift::lift1(v, |v| Ok(Value::Text(v.coerce_to_string_with_ctx(ctx)?)))
        }
        ArgValue::Reference(reference) => phonetic_from_reference(ctx, reference),
        ArgValue::ReferenceUnion(_) => Value::Error(ErrorKind::Value),
    }
}

fn phonetic_from_reference(ctx: &dyn FunctionContext, reference: Reference) -> Value {
    let reference = reference.normalized();
    ctx.record_reference(&reference);

    if reference.is_single_cell() {
        if let Some(phonetic) = ctx.get_cell_phonetic(&reference.sheet_id, reference.start) {
            return Value::Text(phonetic.to_string());
        }
        let value = ctx.get_cell_value(&reference.sheet_id, reference.start);
        return match value.coerce_to_string_with_ctx(ctx) {
            Ok(text) => Value::Text(text),
            Err(e) => Value::Error(e),
        };
    }

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
        if let Some(phonetic) = ctx.get_cell_phonetic(&reference.sheet_id, addr) {
            out.push(Value::Text(phonetic.to_string()));
            continue;
        }

        let value = ctx.get_cell_value(&reference.sheet_id, addr);
        out.push(match value.coerce_to_string_with_ctx(ctx) {
            Ok(text) => Value::Text(text),
            Err(e) => Value::Error(e),
        });
    }

    Value::Array(Array::new(rows, cols, out))
}
