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
        match split_on_any(&text, &row_delimiters, ignore_empty, match_mode) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    };

    let mut rows: Vec<Vec<String>> = Vec::new();
    for row in row_segments {
        let cols = match split_on_any(&row, &col_delimiters, ignore_empty, match_mode) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        };
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
) -> Result<Vec<String>, ErrorKind> {
    match match_mode {
        MatchMode::CaseSensitive => Ok(split_on_any_case_sensitive(text, text, delimiters, ignore_empty)),
        MatchMode::CaseInsensitive => {
            // ASCII fast path: preserve existing TEXTSPLIT behavior and avoid Unicode case-fold allocations.
            if text.is_ascii() && delimiters.iter().all(|d| d.is_ascii()) {
                return Ok(split_on_any_ascii_case_insensitive(text, delimiters, ignore_empty));
            }

            split_on_any_unicode_case_insensitive(text, delimiters, ignore_empty)
        }
    }
}

fn split_on_any_ascii_case_insensitive(
    original: &str,
    delimiters: &[String],
    ignore_empty: bool,
) -> Vec<String> {
    let haystack = original.as_bytes();
    let mut segments = Vec::new();
    let mut cursor = 0usize;
    let mut segment_start = 0usize;

    while let Some((match_pos, match_len)) =
        find_next_delim_ascii_case_insensitive(haystack, cursor, delimiters)
    {
        if !ignore_empty || match_pos > segment_start {
            segments.push(original[segment_start..match_pos].to_string());
        }
        cursor = match_pos + match_len;
        segment_start = cursor;
    }

    if !ignore_empty || segment_start < original.len() {
        segments.push(original[segment_start..].to_string());
    }

    segments
}

fn find_next_delim_ascii_case_insensitive(
    haystack: &[u8],
    from: usize,
    delimiters: &[String],
) -> Option<(usize, usize)> {
    let mut best: Option<(usize, usize)> = None;
    for needle in delimiters {
        if needle.is_empty() {
            continue;
        }
        let needle = needle.as_bytes();
        if from > haystack.len() || haystack.len() - from < needle.len() {
            continue;
        }
        let mut found = None;
        for start in from..=haystack.len() - needle.len() {
            if haystack[start..start + needle.len()].eq_ignore_ascii_case(needle) {
                found = Some(start);
                break;
            }
        }
        let Some(pos) = found else {
            continue;
        };
        let cand = (pos, needle.len());
        best = match best {
            None => Some(cand),
            Some((best_pos, best_len)) => {
                if pos < best_pos || (pos == best_pos && needle.len() > best_len) {
                    Some(cand)
                } else {
                    Some((best_pos, best_len))
                }
            }
        };
    }
    best
}

fn split_on_any_case_sensitive(
    original: &str,
    haystack: &str,
    delimiters: &[String],
    ignore_empty: bool,
) -> Vec<String> {
    let mut segments = Vec::new();
    let mut cursor = 0usize;
    let mut segment_start = 0usize;

    while let Some((match_pos, match_len)) = find_next_delim(haystack, cursor, delimiters) {
        if !ignore_empty || match_pos > segment_start {
            segments.push(original[segment_start..match_pos].to_string());
        }
        cursor = match_pos + match_len;
        segment_start = cursor;
    }

    if !ignore_empty || segment_start < original.len() {
        segments.push(original[segment_start..].to_string());
    }

    segments
}

fn find_next_delim(haystack: &str, from: usize, delimiters: &[String]) -> Option<(usize, usize)> {
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

fn match_delim_at_unicode_case_insensitive(
    haystack_folded: &[char],
    folded_starts: &[usize],
    start_char: usize,
    delim_folded: &[char],
) -> Option<usize> {
    if delim_folded.is_empty() {
        return None;
    }

    let start_folded = *folded_starts.get(start_char)?;
    let end_folded = start_folded.checked_add(delim_folded.len())?;
    let hay_window = haystack_folded.get(start_folded..end_folded)?;
    if hay_window != delim_folded {
        return None;
    }

    // If the delimiter would end "mid-character" after case folding (e.g. trying to match "S"
    // against "ß" which folds to "SS"), treat it as not a match. This keeps delimiter matches
    // aligned to original character boundaries.
    if end_folded == haystack_folded.len() {
        return Some(folded_starts.len());
    }
    folded_starts.binary_search(&end_folded).ok()
}

fn split_on_any_unicode_case_insensitive(
    text: &str,
    delimiters: &[String],
    ignore_empty: bool,
) -> Result<Vec<String>, ErrorKind> {
    let mut char_starts: Vec<usize> = Vec::new();
    let mut hay_folded: Vec<char> = Vec::new();
    let mut folded_starts: Vec<usize> = Vec::new();
    for (byte_idx, ch) in text.char_indices() {
        char_starts.push(byte_idx);
        folded_starts.push(hay_folded.len());
        if ch.is_ascii() {
            hay_folded.push(ch.to_ascii_uppercase());
        } else {
            hay_folded.extend(ch.to_uppercase());
        }
    }
    char_starts.push(text.len());
    let hay_len = folded_starts.len();

    let mut folded_delimiters: Vec<Vec<char>> = Vec::new();
    if folded_delimiters.try_reserve_exact(delimiters.len()).is_err() {
        debug_assert!(
            false,
            "allocation failed (TEXTSPLIT folded delimiters, len={})",
            delimiters.len()
        );
        return Err(ErrorKind::Num);
    }
    for d in delimiters {
        let folded = crate::value::fold_to_uppercase_chars(d);
        if folded.is_empty() && !d.is_empty() {
            debug_assert!(false, "allocation failed (TEXTSPLIT fold delimiter)");
            return Err(ErrorKind::Num);
        }
        folded_delimiters.push(folded);
    }

    let mut segments = Vec::new();
    let mut cursor = 0usize;
    let mut segment_start = 0usize;

    while cursor < hay_len {
        // Find the next delimiter match by scanning forward. This is O(n * delimiters) but keeps
        // indices aligned to original text character boundaries even when Unicode case folding
        // changes length (e.g. ß -> SS).
        let mut found: Option<(usize, usize)> = None;
        for i in cursor..hay_len {
            let mut best_end: Option<usize> = None;
            for delim in &folded_delimiters {
                if let Some(end) =
                    match_delim_at_unicode_case_insensitive(&hay_folded, &folded_starts, i, delim)
                {
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

        let start = char_starts[segment_start];
        let end = char_starts[match_pos];
        if !ignore_empty || end > start {
            segments.push(text[start..end].to_string());
        }
        cursor = match_pos + match_len;
        segment_start = cursor;
    }

    let tail_start = char_starts[segment_start];
    if !ignore_empty || tail_start < text.len() {
        segments.push(text[tail_start..].to_string());
    }

    Ok(segments)
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
