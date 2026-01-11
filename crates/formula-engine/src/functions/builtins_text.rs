use crate::eval::CompiledExpr;
use crate::error::ExcelError;
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{ErrorKind, Value};

const VAR_ARGS: usize = 255;

inventory::submit! {
    FunctionSpec {
        name: "CONCAT",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Any],
        implementation: concat_fn,
    }
}

fn concat_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut out = String::new();
    for arg in args {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => match v.coerce_to_string() {
                Ok(s) => out.push_str(&s),
                Err(e) => return Value::Error(e),
            },
            ArgValue::Reference(r) => {
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(r.sheet_id, addr);
                    match v.coerce_to_string() {
                        Ok(s) => out.push_str(&s),
                        Err(e) => return Value::Error(e),
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                for r in ranges {
                    for addr in r.iter_cells() {
                        let v = ctx.get_cell_value(r.sheet_id, addr);
                        match v.coerce_to_string() {
                            Ok(s) => out.push_str(&s),
                            Err(e) => return Value::Error(e),
                        }
                    }
                }
            }
        }
    }
    Value::Text(out)
}

inventory::submit! {
    FunctionSpec {
        name: "CONCATENATE",
        min_args: 1,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Any],
        implementation: concatenate_fn,
    }
}

fn concatenate_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let mut out = String::new();
    for arg in args {
        let v = eval_scalar_arg(ctx, arg);
        match v.coerce_to_string() {
            Ok(s) => out.push_str(&s),
            Err(e) => return Value::Error(e),
        }
    }
    Value::Text(out)
}

inventory::submit! {
    FunctionSpec {
        name: "LEFT",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: left_fn,
    }
}

fn left_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = match eval_scalar_arg(ctx, &args[0]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let n = if args.len() == 2 {
        match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        }
    } else {
        1
    };
    if n < 0 {
        return Value::Error(ErrorKind::Value);
    }
    Value::Text(slice_chars(&text, 0, n as usize))
}

inventory::submit! {
    FunctionSpec {
        name: "RIGHT",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: right_fn,
    }
}

fn right_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = match eval_scalar_arg(ctx, &args[0]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let n = if args.len() == 2 {
        match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        }
    } else {
        1
    };
    if n < 0 {
        return Value::Error(ErrorKind::Value);
    }
    let chars: Vec<char> = text.chars().collect();
    let len = chars.len();
    let start = len.saturating_sub(n as usize);
    Value::Text(chars[start..].iter().collect())
}

inventory::submit! {
    FunctionSpec {
        name: "MID",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number, ValueType::Number],
        implementation: mid_fn,
    }
}

fn mid_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = match eval_scalar_arg(ctx, &args[0]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let start = match eval_scalar_arg(ctx, &args[1]).coerce_to_i64() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    let len = match eval_scalar_arg(ctx, &args[2]).coerce_to_i64() {
        Ok(n) => n,
        Err(e) => return Value::Error(e),
    };
    if start < 1 || len < 0 {
        return Value::Error(ErrorKind::Value);
    }
    let start_idx = (start - 1) as usize;
    Value::Text(slice_chars(&text, start_idx, len as usize))
}

inventory::submit! {
    FunctionSpec {
        name: "LEN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: len_fn,
    }
}

fn len_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = match eval_scalar_arg(ctx, &args[0]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    Value::Number(text.chars().count() as f64)
}

inventory::submit! {
    FunctionSpec {
        name: "TRIM",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: trim_fn,
    }
}

fn trim_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = match eval_scalar_arg(ctx, &args[0]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    Value::Text(excel_trim(&text))
}

inventory::submit! {
    FunctionSpec {
        name: "UPPER",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: upper_fn,
    }
}

fn upper_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = match eval_scalar_arg(ctx, &args[0]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    Value::Text(text.to_uppercase())
}

inventory::submit! {
    FunctionSpec {
        name: "LOWER",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: lower_fn,
    }
}

fn lower_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = match eval_scalar_arg(ctx, &args[0]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    Value::Text(text.to_lowercase())
}

inventory::submit! {
    FunctionSpec {
        name: "FIND",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Number],
        implementation: find_fn,
    }
}

fn find_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let needle = match eval_scalar_arg(ctx, &args[0]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let haystack = match eval_scalar_arg(ctx, &args[1]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let start = if args.len() == 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_i64() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        }
    } else {
        1
    };
    find_impl(&needle, &haystack, start, false)
}

inventory::submit! {
    FunctionSpec {
        name: "SEARCH",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Number],
        implementation: search_fn,
    }
}

fn search_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let needle = match eval_scalar_arg(ctx, &args[0]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let haystack = match eval_scalar_arg(ctx, &args[1]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let start = if args.len() == 3 {
        match eval_scalar_arg(ctx, &args[2]).coerce_to_i64() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        }
    } else {
        1
    };
    find_impl(&needle, &haystack, start, true)
}

inventory::submit! {
    FunctionSpec {
        name: "SUBSTITUTE",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Text, ValueType::Number],
        implementation: substitute_fn,
    }
}

fn substitute_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = match eval_scalar_arg(ctx, &args[0]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let old_text = match eval_scalar_arg(ctx, &args[1]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let new_text = match eval_scalar_arg(ctx, &args[2]).coerce_to_string() {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };

    let instance_num = if args.len() == 4 {
        let raw = match eval_scalar_arg(ctx, &args[3]).coerce_to_i64() {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        match i32::try_from(raw) {
            Ok(n) => Some(n),
            Err(_) => return Value::Error(ErrorKind::Value),
        }
    } else {
        None
    };

    match crate::functions::text::substitute(&text, &old_text, &new_text, instance_num) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(match e {
            ExcelError::Div0 => ErrorKind::Div0,
            ExcelError::Value => ErrorKind::Value,
            ExcelError::Num => ErrorKind::Num,
        }),
    }
}

fn slice_chars(text: &str, start: usize, len: usize) -> String {
    let chars: Vec<char> = text.chars().collect();
    if start >= chars.len() {
        return String::new();
    }
    let end = (start + len).min(chars.len());
    chars[start..end].iter().collect()
}

fn excel_trim(text: &str) -> String {
    let mut out = String::new();
    let mut in_space = false;
    for ch in text.chars() {
        if ch == ' ' {
            in_space = true;
            continue;
        }
        if in_space && !out.is_empty() {
            out.push(' ');
        }
        in_space = false;
        out.push(ch);
    }
    out.trim_matches(' ').to_string()
}

fn find_impl(needle: &str, haystack: &str, start: i64, case_insensitive: bool) -> Value {
    if start < 1 {
        return Value::Error(ErrorKind::Value);
    }
    let needle_chars: Vec<char> = needle.chars().collect();
    let mut hay_chars: Vec<char> = haystack.chars().collect();
    let start_idx = (start - 1) as usize;
    if start_idx > hay_chars.len() {
        return Value::Error(ErrorKind::Value);
    }

    if needle_chars.is_empty() {
        return Value::Number(start as f64);
    }

    if case_insensitive {
        hay_chars = hay_chars.into_iter().map(fold_case).collect();
    }

    let needle_tokens = if case_insensitive {
        parse_search_pattern(&needle_chars.into_iter().map(fold_case).collect::<Vec<_>>())
    } else {
        vec![PatternToken::LiteralSeq(needle_chars)]
    };

    for i in start_idx..hay_chars.len() {
        if matches_pattern(&needle_tokens, &hay_chars, i) {
            return Value::Number((i + 1) as f64);
        }
    }
    Value::Error(ErrorKind::Value)
}

#[derive(Debug, Clone)]
enum PatternToken {
    LiteralSeq(Vec<char>),
    AnyOne,
    AnyMany,
}

fn parse_search_pattern(pattern: &[char]) -> Vec<PatternToken> {
    let mut tokens = Vec::new();
    let mut literal = Vec::new();
    let mut idx = 0;
    while idx < pattern.len() {
        let ch = pattern[idx];
        if ch == '~' {
            idx += 1;
            if idx < pattern.len() {
                literal.push(pattern[idx]);
                idx += 1;
            } else {
                literal.push('~');
            }
            continue;
        }
        match ch {
            '*' => {
                if !literal.is_empty() {
                    tokens.push(PatternToken::LiteralSeq(std::mem::take(&mut literal)));
                }
                tokens.push(PatternToken::AnyMany);
                idx += 1;
            }
            '?' => {
                if !literal.is_empty() {
                    tokens.push(PatternToken::LiteralSeq(std::mem::take(&mut literal)));
                }
                tokens.push(PatternToken::AnyOne);
                idx += 1;
            }
            _ => {
                literal.push(ch);
                idx += 1;
            }
        }
    }
    if !literal.is_empty() {
        tokens.push(PatternToken::LiteralSeq(literal));
    }
    tokens
}

fn matches_pattern(tokens: &[PatternToken], hay: &[char], start: usize) -> bool {
    let mut memo = vec![vec![None; hay.len() + 1]; tokens.len() + 1];
    match_rec(tokens, hay, start, 0, &mut memo)
}

fn match_rec(
    tokens: &[PatternToken],
    hay: &[char],
    hay_idx: usize,
    tok_idx: usize,
    memo: &mut [Vec<Option<bool>>],
) -> bool {
    if let Some(cached) = memo[tok_idx][hay_idx] {
        return cached;
    }
    let result = if tok_idx == tokens.len() {
        true
    } else {
        match &tokens[tok_idx] {
            PatternToken::LiteralSeq(seq) => {
                if hay_idx + seq.len() > hay.len() {
                    false
                } else if hay[hay_idx..hay_idx + seq.len()] == *seq {
                    match_rec(tokens, hay, hay_idx + seq.len(), tok_idx + 1, memo)
                } else {
                    false
                }
            }
            PatternToken::AnyOne => {
                if hay_idx >= hay.len() {
                    false
                } else {
                    match_rec(tokens, hay, hay_idx + 1, tok_idx + 1, memo)
                }
            }
            PatternToken::AnyMany => {
                if match_rec(tokens, hay, hay_idx, tok_idx + 1, memo) {
                    true
                } else if hay_idx < hay.len() {
                    match_rec(tokens, hay, hay_idx + 1, tok_idx, memo)
                } else {
                    false
                }
            }
        }
    };
    memo[tok_idx][hay_idx] = Some(result);
    result
}

fn fold_case(ch: char) -> char {
    if ch.is_ascii() {
        ch.to_ascii_lowercase()
    } else {
        ch
    }
}
