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
    CaseInsensitive,
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
            Ok(1) => MatchMode::CaseInsensitive,
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
    let total = match out_rows.checked_mul(out_cols) {
        Some(v) => v,
        None => return Value::Error(ErrorKind::Spill),
    };
    if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
        return Value::Error(ErrorKind::Spill);
    }
    let mut values = Vec::new();
    if values.try_reserve_exact(total).is_err() {
        return Value::Error(ErrorKind::Num);
    }
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

fn eval_delimiter_set(
    ctx: &dyn FunctionContext,
    expr: &CompiledExpr,
) -> Result<Vec<String>, ErrorKind> {
    let v = eval_scalar_arg(ctx, expr);
    match v {
        Value::Array(arr) => {
            let mut out = Vec::new();
            if out.try_reserve_exact(arr.values.len()).is_err() {
                return Err(ErrorKind::Num);
            }
            for v in arr.iter() {
                out.push(v.coerce_to_string_with_ctx(ctx)?);
            }
            Ok(out)
        }
        other => Ok(vec![other.coerce_to_string_with_ctx(ctx)?]),
    }
}

fn split_on_any(
    text: &str,
    delimiters: &[String],
    ignore_empty: bool,
    match_mode: MatchMode,
) -> Vec<String> {
    match match_mode {
        MatchMode::CaseSensitive => {
            let delims: Vec<&str> = delimiters.iter().map(|d| d.as_str()).collect();
            split_on_any_impl(text, text, &delims, ignore_empty)
        }
        MatchMode::CaseInsensitive => {
            // ASCII fast path: preserve existing TEXTSPLIT behavior and avoid Unicode case-fold allocations.
            if text.is_ascii() && delimiters.iter().all(|d| d.is_ascii()) {
                let haystack = text.to_ascii_lowercase();
                let lowered: Vec<String> =
                    delimiters.iter().map(|s| s.to_ascii_lowercase()).collect();
                let delims: Vec<&str> = lowered.iter().map(|d| d.as_str()).collect();
                return split_on_any_impl(text, haystack.as_str(), &delims, ignore_empty);
            }

            split_on_any_unicode_case_insensitive(text, delimiters, ignore_empty)
        }
    }
}

fn split_on_any_impl(
    original: &str,
    haystack: &str,
    delimiters: &[&str],
    ignore_empty: bool,
) -> Vec<String> {
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

fn fold_char_uppercase(c: char, out: &mut Vec<char>) {
    if c.is_ascii() {
        out.push(c.to_ascii_uppercase());
    } else {
        out.extend(c.to_uppercase());
    }
}

fn fold_str_uppercase(s: &str) -> Vec<char> {
    let mut out = Vec::new();
    for c in s.chars() {
        fold_char_uppercase(c, &mut out);
    }
    out
}

fn match_delim_at_unicode_case_insensitive(
    haystack_chars: &[char],
    start: usize,
    delim_folded: &[char],
) -> Option<usize> {
    if delim_folded.is_empty() {
        return None;
    }

    let mut di = 0usize;
    let mut hi = start;
    while di < delim_folded.len() {
        let ch = *haystack_chars.get(hi)?;
        if ch.is_ascii() {
            let fc = ch.to_ascii_uppercase();
            if di >= delim_folded.len() || fc != delim_folded[di] {
                return None;
            }
            di += 1;
        } else {
            for fc in ch.to_uppercase() {
                // If the delimiter would end "mid-character" after case folding (e.g. trying to match
                // "S" against "ß" which folds to "SS"), treat it as not a match. This keeps delimiter
                // matches aligned to original character boundaries.
                if di >= delim_folded.len() {
                    return None;
                }
                if fc != delim_folded[di] {
                    return None;
                }
                di += 1;
            }
        }
        hi += 1;
    }

    Some(hi)
}

fn split_on_any_unicode_case_insensitive(
    text: &str,
    delimiters: &[String],
    ignore_empty: bool,
) -> Vec<String> {
    let hay_chars: Vec<char> = text.chars().collect();
    let mut char_starts: Vec<usize> = text.char_indices().map(|(i, _)| i).collect();
    char_starts.push(text.len());

    let folded_delimiters: Vec<Vec<char>> =
        delimiters.iter().map(|d| fold_str_uppercase(d)).collect();

    let mut segments = Vec::new();
    let mut cursor = 0usize;
    let mut segment_start = 0usize;

    while cursor < hay_chars.len() {
        // Find the next delimiter match by scanning forward. This is O(n * delimiters) but keeps
        // indices aligned to original text character boundaries even when Unicode case folding
        // changes length (e.g. ß -> SS).
        let mut found: Option<(usize, usize)> = None;
        for i in cursor..hay_chars.len() {
            let mut best_end: Option<usize> = None;
            for delim in &folded_delimiters {
                if let Some(end) = match_delim_at_unicode_case_insensitive(&hay_chars, i, delim) {
                    best_end = match best_end {
                        None => Some(end),
                        Some(prev) if end > prev => Some(end),
                        Some(prev) => Some(prev),
                    };
                }
            }
            if let Some(end) = best_end {
                found = Some((i, end - i));
                break;
            }
        }

        let Some((match_pos, match_len)) = found else {
            break;
        };

        let piece = text[char_starts[segment_start]..char_starts[match_pos]].to_string();
        segments.push(piece);
        cursor = match_pos + match_len;
        segment_start = cursor;
    }

    segments.push(text[char_starts[segment_start]..].to_string());

    if ignore_empty {
        segments.retain(|s| !s.is_empty());
    }

    segments
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
