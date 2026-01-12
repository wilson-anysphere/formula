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
//! `PHONETIC` depends on per-cell phonetic guide metadata, which the engine does
//! not model yet. We implement a deterministic placeholder that returns the
//! referenced value coerced to text.
//!
//! Once workbook locale + codepage + phonetic metadata are modeled, this module
//! can be extended to implement real Excel semantics for DBCS workbooks.

use crate::eval::CompiledExpr;
use crate::functions::array_lift;
use crate::functions::{call_function, FunctionContext};
use crate::value::Value;

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
    array_lift::lift1(text, |text| Ok(Value::Text(text.coerce_to_string()?)))
}

pub(crate) fn dbcs_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(text, |text| Ok(Value::Text(text.coerce_to_string()?)))
}

pub(crate) fn phonetic_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    // Placeholder: return the referenced value coerced to text.
    //
    // NOTE: Real Excel uses a phonetic guide (furigana) stored in cell metadata,
    // which is not currently modeled in the engine.
    let reference = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(reference, |v| Ok(Value::Text(v.coerce_to_string()?)))
}

