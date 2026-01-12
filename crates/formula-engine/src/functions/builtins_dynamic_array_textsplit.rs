use crate::eval::CompiledExpr;
use crate::eval::Expr;
use crate::functions::{eval_scalar_arg, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{Array, ErrorKind, Value};

inventory::submit! {
    FunctionSpec {
        name: "TEXTSPLIT",
        min_args: 2,
        max_args: 6,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Any,
        arg_types: &[
            ValueType::Text,
            ValueType::Any,
            ValueType::Any,
            ValueType::Bool,
            ValueType::Number,
            ValueType::Any,
        ],
        implementation: textsplit_fn,
    }
}

#[derive(Debug, Clone, Copy)]
enum MatchMode {
    CaseSensitive,
    CaseInsensitiveAscii,
}

fn textsplit_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };

    let col_delimiters = match eval_delimiter_set(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    if col_delimiters.iter().any(|d| d.is_empty()) {
        return Value::Error(ErrorKind::Value);
    }

    let row_delimiters = if args.len() >= 3 {
        match eval_optional_row_delimiters(ctx, &args[2]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        Vec::new()
    };

    let ignore_empty = if args.len() >= 4 {
        match eval_scalar_arg(ctx, &args[3]).coerce_to_bool_with_ctx(ctx) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        false
    };

    let match_mode = if args.len() >= 5 {
        match eval_scalar_arg(ctx, &args[4]).coerce_to_i64_with_ctx(ctx) {
            Ok(0) => MatchMode::CaseSensitive,
            Ok(1) => MatchMode::CaseInsensitiveAscii,
            Ok(_) => return Value::Error(ErrorKind::Value),
            Err(e) => return Value::Error(e),
        }
    } else {
        MatchMode::CaseSensitive
    };

    let pad_with = if args.len() >= 6 {
        if matches!(args[5], Expr::Blank) {
            Value::Error(ErrorKind::NA)
        } else {
        let v = ctx.eval_scalar(&args[5]);
        match v {
            Value::Array(_) | Value::Spill { .. } => return Value::Error(ErrorKind::Value),
            other => other,
        }
        }
    } else {
        Value::Error(ErrorKind::NA)
    };

    let row_segments = if row_delimiters.is_empty() {
        vec![text.clone()]
    } else {
        split_on_any(&text, &row_delimiters, ignore_empty, match_mode)
    };

    let mut rows: Vec<Vec<String>> = Vec::new();
    for row in row_segments {
        let cols = split_on_any(&row, &col_delimiters, ignore_empty, match_mode);
        rows.push(cols);
    }

    if rows.is_empty() {
        return Value::Error(ErrorKind::Calc);
    }

    let out_cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    if out_cols == 0 {
        return Value::Error(ErrorKind::Calc);
    }

    let out_rows = rows.len();
    let mut values = Vec::with_capacity(out_rows.saturating_mul(out_cols));
    for row in rows {
        for col_idx in 0..out_cols {
            if let Some(cell) = row.get(col_idx) {
                values.push(Value::Text(cell.clone()));
            } else {
                values.push(pad_with.clone());
            }
        }
    }

    Value::Array(Array::new(out_rows, out_cols, values))
}

fn eval_optional_row_delimiters(
    ctx: &dyn FunctionContext,
    expr: &CompiledExpr,
) -> Result<Vec<String>, ErrorKind> {
    let delimiters = eval_delimiter_set(ctx, expr)?;

    // Excel treats a missing/blank row_delimiter as "no row split".
    if delimiters.iter().all(|d| d.is_empty()) {
        return Ok(Vec::new());
    }

    // Disallow empty delimiters when combined with other delimiter values.
    if delimiters.iter().any(|d| d.is_empty()) {
        return Err(ErrorKind::Value);
    }

    Ok(delimiters)
}

fn eval_delimiter_set(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<Vec<String>, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    match v {
        Value::Array(arr) => {
            let mut out = Vec::with_capacity(arr.values.len());
            for v in arr.iter() {
                out.push(v.coerce_to_string_with_ctx(ctx)?);
            }
            Ok(out)
        }
        other => Ok(vec![other.coerce_to_string_with_ctx(ctx)?]),
    }
}

fn split_on_any(text: &str, delimiters: &[String], ignore_empty: bool, match_mode: MatchMode) -> Vec<String> {
    match match_mode {
        MatchMode::CaseSensitive => {
            let delims: Vec<&str> = delimiters.iter().map(|d| d.as_str()).collect();
            split_on_any_impl(text, text, &delims, ignore_empty)
        }
        MatchMode::CaseInsensitiveAscii => {
            let haystack = text.to_ascii_lowercase();
            let lowered: Vec<String> = delimiters.iter().map(|s| s.to_ascii_lowercase()).collect();
            let delims: Vec<&str> = lowered.iter().map(|d| d.as_str()).collect();
            split_on_any_impl(text, haystack.as_str(), &delims, ignore_empty)
        }
    }
}

fn split_on_any_impl(original: &str, haystack: &str, delimiters: &[&str], ignore_empty: bool) -> Vec<String> {
    let mut segments = Vec::new();
    let mut cursor = 0usize;
    let mut segment_start = 0usize;

    while let Some((match_pos, match_len)) = find_next_delim(haystack, cursor, delimiters) {
        let piece = original[segment_start..match_pos].to_string();
        segments.push(piece);
        cursor = match_pos + match_len;
        segment_start = cursor;
    }

    segments.push(original[segment_start..].to_string());

    if ignore_empty {
        segments.retain(|s| !s.is_empty());
    }

    segments
}

fn find_next_delim(haystack: &str, from: usize, delimiters: &[&str]) -> Option<(usize, usize)> {
    let mut best: Option<(usize, usize)> = None;
    for delim in delimiters {
        if delim.is_empty() {
            continue;
        }
        let Some(rel) = haystack[from..].find(delim) else {
            continue;
        };
        let abs = from + rel;
        let cand = (abs, delim.len());
        best = match best {
            None => Some(cand),
            Some((best_pos, best_len)) => {
                if abs < best_pos || (abs == best_pos && delim.len() > best_len) {
                    Some(cand)
                } else {
                    Some((best_pos, best_len))
                }
            }
        };
    }
    best
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
