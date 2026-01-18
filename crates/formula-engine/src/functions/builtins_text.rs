use crate::error::ExcelError;
use crate::eval::CompiledExpr;
use crate::functions::array_lift;
use crate::functions::text::search_pattern::{
    matches_pattern_with_memo, min_required_hay_len, parse_search_pattern_folded, PatternToken,
};
use crate::functions::{eval_scalar_arg, ArgValue, ArraySupport, FunctionContext, FunctionSpec};
use crate::functions::{ThreadSafety, ValueType, Volatility};
use crate::value::{casefold_owned, lowercase_owned, Array, ErrorKind, Value};
use std::collections::HashSet;

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
            ArgValue::Scalar(v) => match v {
                Value::Array(arr) => {
                    for v in arr.iter() {
                        match v.coerce_to_string_with_ctx(ctx) {
                            Ok(s) => out.push_str(&s),
                            Err(e) => return Value::Error(e),
                        }
                    }
                }
                other => match other.coerce_to_string_with_ctx(ctx) {
                    Ok(s) => out.push_str(&s),
                    Err(e) => return Value::Error(e),
                },
            },
            ArgValue::Reference(r) => {
                let r = r.normalized();
                ctx.record_reference(&r);
                for addr in r.iter_cells() {
                    let v = ctx.get_cell_value(&r.sheet_id, addr);
                    match v.coerce_to_string_with_ctx(ctx) {
                        Ok(s) => out.push_str(&s),
                        Err(e) => return Value::Error(e),
                    }
                }
            }
            ArgValue::ReferenceUnion(ranges) => {
                // Excel reference unions behave like set union: overlapping cells should only be
                // visited once. Preserve stable ordering by iterating ranges in `eval_arg` order
                // and skipping already-seen `(sheet, cell)` pairs.
                let mut seen: HashSet<(crate::functions::SheetId, crate::eval::CellAddr)> =
                    HashSet::new();
                for r in ranges {
                    let r = r.normalized();
                    ctx.record_reference(&r);
                    for addr in r.iter_cells() {
                        if !seen.insert((r.sheet_id.clone(), addr)) {
                            continue;
                        }
                        let v = ctx.get_cell_value(&r.sheet_id, addr);
                        match v.coerce_to_string_with_ctx(ctx) {
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
        match v.coerce_to_string_with_ctx(ctx) {
            Ok(s) => out.push_str(&s),
            Err(e) => return Value::Error(e),
        }
    }
    Value::Text(out)
}

inventory::submit! {
    FunctionSpec {
        name: "HYPERLINK",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::ScalarOnly,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Any, ValueType::Any],
        implementation: hyperlink_fn,
    }
}

fn hyperlink_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let link_location = eval_scalar_arg(ctx, &args[0]);
    if let Value::Error(e) = link_location {
        return Value::Error(e);
    }

    let friendly_name = match args.get(1) {
        Some(expr) if matches!(expr, CompiledExpr::Blank) => None,
        Some(expr) => {
            let v = eval_scalar_arg(ctx, expr);
            if let Value::Error(e) = v {
                return Value::Error(e);
            }
            Some(v)
        }
        None => None,
    };

    let display = match friendly_name {
        Some(v) => v,
        None => link_location,
    };

    match display.coerce_to_string_with_ctx(ctx) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "LEFT",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: left_fn,
    }
}

fn left_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    let n = if args.len() == 2 {
        array_lift::eval_arg(ctx, &args[1])
    } else {
        Value::Number(1.0)
    };
    array_lift::lift2(text, n, |text, n| {
        let text = text.coerce_to_string_with_ctx(ctx)?;
        let n = n.coerce_to_i64_with_ctx(ctx)?;
        if n < 0 {
            return Err(ErrorKind::Value);
        }
        Ok(Value::Text(slice_chars(&text, 0, n as usize)))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "RIGHT",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number],
        implementation: right_fn,
    }
}

fn right_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    let n = if args.len() == 2 {
        array_lift::eval_arg(ctx, &args[1])
    } else {
        Value::Number(1.0)
    };
    array_lift::lift2(text, n, |text, n| {
        let text = text.coerce_to_string_with_ctx(ctx)?;
        let n = n.coerce_to_i64_with_ctx(ctx)?;
        if n < 0 {
            return Err(ErrorKind::Value);
        }
        let n = n as usize;
        if n == 0 {
            return Ok(Value::Text(String::new()));
        }
        let start_byte = match text.char_indices().rev().nth(n - 1) {
            Some((i, _)) => i,
            None => 0,
        };
        Ok(Value::Text(text[start_byte..].to_string()))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "MID",
        min_args: 3,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number, ValueType::Number],
        implementation: mid_fn,
    }
}

fn mid_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    let start = array_lift::eval_arg(ctx, &args[1]);
    let len = array_lift::eval_arg(ctx, &args[2]);
    array_lift::lift3(text, start, len, |text, start, len| {
        let text = text.coerce_to_string_with_ctx(ctx)?;
        let start = start.coerce_to_i64_with_ctx(ctx)?;
        let len = len.coerce_to_i64_with_ctx(ctx)?;
        if start < 1 || len < 0 {
            return Err(ErrorKind::Value);
        }
        let start_idx = (start - 1) as usize;
        Ok(Value::Text(slice_chars(&text, start_idx, len as usize)))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "LEN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: len_fn,
    }
}

fn len_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(text, |text| {
        let text = text.coerce_to_string_with_ctx(ctx)?;
        Ok(Value::Number(text.chars().count() as f64))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "TRIM",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: trim_fn,
    }
}

fn trim_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(text, |text| {
        let text = text.coerce_to_string_with_ctx(ctx)?;
        Ok(Value::Text(excel_trim(&text)))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "UPPER",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: upper_fn,
    }
}

fn upper_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(text, |text| {
        Ok(Value::Text(casefold_owned(
            text.coerce_to_string_with_ctx(ctx)?,
        )))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "LOWER",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: lower_fn,
    }
}

fn lower_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    array_lift::lift1(text, |text| {
        Ok(Value::Text(lowercase_owned(
            text.coerce_to_string_with_ctx(ctx)?,
        )))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "FIND",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Number],
        implementation: find_fn,
    }
}

fn find_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let needle = array_lift::eval_arg(ctx, &args[0]);
    let haystack = array_lift::eval_arg(ctx, &args[1]);
    let start = if args.len() == 3 {
        array_lift::eval_arg(ctx, &args[2])
    } else {
        Value::Number(1.0)
    };
    array_lift::lift3(needle, haystack, start, |needle, haystack, start| {
        let needle = needle.coerce_to_string_with_ctx(ctx)?;
        let haystack = haystack.coerce_to_string_with_ctx(ctx)?;
        let start = start.coerce_to_i64_with_ctx(ctx)?;
        Ok(find_impl(&needle, &haystack, start, false))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "SEARCH",
        min_args: 2,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Number],
        implementation: search_fn,
    }
}

fn search_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let needle = array_lift::eval_arg(ctx, &args[0]);
    let haystack = array_lift::eval_arg(ctx, &args[1]);
    let start = if args.len() == 3 {
        array_lift::eval_arg(ctx, &args[2])
    } else {
        Value::Number(1.0)
    };
    array_lift::lift3(needle, haystack, start, |needle, haystack, start| {
        let needle = needle.coerce_to_string_with_ctx(ctx)?;
        let haystack = haystack.coerce_to_string_with_ctx(ctx)?;
        let start = start.coerce_to_i64_with_ctx(ctx)?;
        Ok(find_impl(&needle, &haystack, start, true))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "SUBSTITUTE",
        min_args: 3,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Text, ValueType::Number],
        implementation: substitute_fn,
    }
}

fn substitute_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = array_lift::eval_arg(ctx, &args[0]);
    let old_text = array_lift::eval_arg(ctx, &args[1]);
    let new_text = array_lift::eval_arg(ctx, &args[2]);

    if args.len() == 4 {
        let instance_num = array_lift::eval_arg(ctx, &args[3]);
        return array_lift::lift4(
            text,
            old_text,
            new_text,
            instance_num,
            |text, old_text, new_text, instance_num| {
                let text = text.coerce_to_string_with_ctx(ctx)?;
                let old_text = old_text.coerce_to_string_with_ctx(ctx)?;
                let new_text = new_text.coerce_to_string_with_ctx(ctx)?;
                let raw = instance_num.coerce_to_i64_with_ctx(ctx)?;
                let instance_num = i32::try_from(raw).map_err(|_| ErrorKind::Value)?;

                match crate::functions::text::substitute(
                    &text,
                    &old_text,
                    &new_text,
                    Some(instance_num),
                ) {
                    Ok(s) => Ok(Value::Text(s)),
                    Err(e) => Err(match e {
                        ExcelError::Div0 => ErrorKind::Div0,
                        ExcelError::Value => ErrorKind::Value,
                        ExcelError::Num => ErrorKind::Num,
                    }),
                }
            },
        );
    }

    array_lift::lift3(text, old_text, new_text, |text, old_text, new_text| {
        let text = text.coerce_to_string_with_ctx(ctx)?;
        let old_text = old_text.coerce_to_string_with_ctx(ctx)?;
        let new_text = new_text.coerce_to_string_with_ctx(ctx)?;
        match crate::functions::text::substitute(&text, &old_text, &new_text, None) {
            Ok(s) => Ok(Value::Text(s)),
            Err(e) => Err(match e {
                ExcelError::Div0 => ErrorKind::Div0,
                ExcelError::Value => ErrorKind::Value,
                ExcelError::Num => ErrorKind::Num,
            }),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "VALUE",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text],
        implementation: value_fn,
    }
}

fn value_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let arg = match eval_matrix_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let cfg = ctx.value_locale();
    let now_utc = ctx.now_utc();
    let system = ctx.date_system();

    elementwise_unary(arg, |v| match v {
        Value::Error(e) => Value::Error(*e),
        Value::Number(n) => {
            if n.is_finite() {
                Value::Number(*n)
            } else {
                Value::Error(ErrorKind::Num)
            }
        }
        other => {
            let text = match other.coerce_to_string_with_ctx(ctx) {
                Ok(s) => s,
                Err(e) => return Value::Error(e),
            };
            match crate::functions::text::value_with_locale(&text, cfg, now_utc, system) {
                Ok(n) => Value::Number(n),
                Err(e) => Value::Error(excel_error_to_kind(e)),
            }
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "NUMBERVALUE",
        min_args: 1,
        max_args: 3,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Number,
        arg_types: &[ValueType::Text, ValueType::Text, ValueType::Text],
        implementation: numbervalue_fn,
    }
}

fn numbervalue_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number_text = match eval_matrix_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let separators = ctx.value_locale().separators;

    let decimal_sep = if args.len() >= 2 {
        match eval_matrix_arg(ctx, &args[1]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        MatrixArg::Scalar(Value::Text(separators.decimal_sep.to_string()))
    };

    let group_sep = if args.len() >= 3 {
        match eval_matrix_arg(ctx, &args[2]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        MatrixArg::Scalar(Value::Text(separators.thousands_sep.to_string()))
    };

    elementwise_ternary(number_text, decimal_sep, group_sep, |raw, dec, group| {
        if let Value::Error(e) = raw {
            return Value::Error(*e);
        }

        if let Value::Number(n) = raw {
            if n.is_finite() {
                return Value::Number(*n);
            }
            return Value::Error(ErrorKind::Num);
        }

        let text = match raw.coerce_to_string_with_ctx(ctx) {
            Ok(s) => s,
            Err(e) => return Value::Error(e),
        };
        let decimal = match coerce_single_char(dec, ctx) {
            Ok(ch) => ch,
            Err(e) => return Value::Error(e),
        };
        let group = match coerce_optional_single_char(group, ctx) {
            Ok(ch) => ch,
            Err(e) => return Value::Error(e),
        };

        if group == Some(decimal) {
            return Value::Error(ErrorKind::Value);
        }

        match crate::functions::text::numbervalue(&text, Some(decimal), group) {
            Ok(n) => Value::Number(n),
            Err(e) => Value::Error(excel_error_to_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "TEXT",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Any, ValueType::Text],
        implementation: text_fn,
    }
}

fn text_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let value = match eval_matrix_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let format_text = match eval_matrix_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let date_system = ctx.date_system();
    elementwise_binary(value, format_text, |value, fmt| {
        let fmt = match fmt.coerce_to_string_with_ctx(ctx) {
            Ok(s) => s,
            Err(e) => return Value::Error(e),
        };
        match crate::functions::text::text(value, &fmt, date_system, ctx.value_locale()) {
            Ok(s) => Value::Text(s),
            Err(e) => Value::Error(e),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "DOLLAR",
        min_args: 1,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Number, ValueType::Number],
        implementation: dollar_fn,
    }
}

fn dollar_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let number = match eval_matrix_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    let decimals = if args.len() >= 2 {
        match eval_matrix_arg(ctx, &args[1]) {
            Ok(v) => v,
            Err(e) => return Value::Error(e),
        }
    } else {
        MatrixArg::Scalar(Value::Number(2.0))
    };

    elementwise_binary(number, decimals, |number, decimals| {
        let number = match number.coerce_to_number_with_ctx(ctx) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        if !number.is_finite() {
            return Value::Error(ErrorKind::Num);
        }

        let decimals_raw = match decimals.coerce_to_i64_with_ctx(ctx) {
            Ok(n) => n,
            Err(e) => return Value::Error(e),
        };
        let decimals = match i32::try_from(decimals_raw) {
            Ok(n) => n,
            Err(_) => return Value::Error(ErrorKind::Num),
        };

        match crate::functions::text::dollar(number, Some(decimals), ctx.value_locale()) {
            Ok(s) => Value::Text(s),
            Err(e) => Value::Error(excel_error_to_kind(e)),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "TEXTJOIN",
        min_args: 3,
        max_args: VAR_ARGS,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Bool, ValueType::Any],
        implementation: textjoin_fn,
    }
}

fn textjoin_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let delimiter = match eval_scalar_arg(ctx, &args[0]).coerce_to_string_with_ctx(ctx) {
        Ok(s) => s,
        Err(e) => return Value::Error(e),
    };
    let ignore_empty = match eval_scalar_arg(ctx, &args[1]).coerce_to_bool_with_ctx(ctx) {
        Ok(b) => b,
        Err(e) => return Value::Error(e),
    };

    let mut values = Vec::new();
    for arg in &args[2..] {
        match ctx.eval_arg(arg) {
            ArgValue::Scalar(v) => flatten_textjoin_value(&mut values, v),
            ArgValue::Reference(r) => flatten_textjoin_reference(ctx, &mut values, r),
            ArgValue::ReferenceUnion(ranges) => {
                flatten_textjoin_reference_union(ctx, &mut values, &ranges)
            }
        }
    }

    match crate::functions::text::textjoin(
        &delimiter,
        ignore_empty,
        &values,
        ctx.date_system(),
        ctx.value_locale(),
    ) {
        Ok(s) => Value::Text(s),
        Err(e) => Value::Error(e),
    }
}

inventory::submit! {
    FunctionSpec {
        name: "CLEAN",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: clean_fn,
    }
}

fn clean_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = match eval_matrix_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    elementwise_unary(text, |v| {
        let text = match v.coerce_to_string_with_ctx(ctx) {
            Ok(s) => s,
            Err(e) => return Value::Error(e),
        };
        match crate::functions::text::clean(&text) {
            Ok(s) => Value::Text(s),
            Err(e) => Value::Error(e),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "EXACT",
        min_args: 2,
        max_args: 2,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Bool,
        arg_types: &[ValueType::Text, ValueType::Text],
        implementation: exact_fn,
    }
}

fn exact_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text1 = match eval_matrix_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let text2 = match eval_matrix_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    elementwise_binary(text1, text2, |left, right| {
        let left = match left.coerce_to_string_with_ctx(ctx) {
            Ok(s) => s,
            Err(e) => return Value::Error(e),
        };
        let right = match right.coerce_to_string_with_ctx(ctx) {
            Ok(s) => s,
            Err(e) => return Value::Error(e),
        };
        Value::Bool(crate::functions::text::exact(&left, &right))
    })
}

inventory::submit! {
    FunctionSpec {
        name: "PROPER",
        min_args: 1,
        max_args: 1,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text],
        implementation: proper_fn,
    }
}

fn proper_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let text = match eval_matrix_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    elementwise_unary(text, |v| {
        let text = match v.coerce_to_string_with_ctx(ctx) {
            Ok(s) => s,
            Err(e) => return Value::Error(e),
        };
        match crate::functions::text::proper(&text) {
            Ok(s) => Value::Text(s),
            Err(e) => Value::Error(e),
        }
    })
}

inventory::submit! {
    FunctionSpec {
        name: "REPLACE",
        min_args: 4,
        max_args: 4,
        volatility: Volatility::NonVolatile,
        thread_safety: ThreadSafety::ThreadSafe,
        array_support: ArraySupport::SupportsArrays,
        return_type: ValueType::Text,
        arg_types: &[ValueType::Text, ValueType::Number, ValueType::Number, ValueType::Text],
        implementation: replace_fn,
    }
}

fn replace_fn(ctx: &dyn FunctionContext, args: &[CompiledExpr]) -> Value {
    let old_text = match eval_matrix_arg(ctx, &args[0]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let start_num = match eval_matrix_arg(ctx, &args[1]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let num_chars = match eval_matrix_arg(ctx, &args[2]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };
    let new_text = match eval_matrix_arg(ctx, &args[3]) {
        Ok(v) => v,
        Err(e) => return Value::Error(e),
    };

    elementwise_quaternary(
        old_text,
        start_num,
        num_chars,
        new_text,
        |old, start, num, new| {
            let old = match old.coerce_to_string_with_ctx(ctx) {
                Ok(s) => s,
                Err(e) => return Value::Error(e),
            };
            let start = match start
                .coerce_to_i64_with_ctx(ctx)
                .and_then(|n| i32::try_from(n).map_err(|_| ErrorKind::Value))
            {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            let num = match num
                .coerce_to_i64_with_ctx(ctx)
                .and_then(|n| i32::try_from(n).map_err(|_| ErrorKind::Value))
            {
                Ok(n) => n,
                Err(e) => return Value::Error(e),
            };
            let new = match new.coerce_to_string_with_ctx(ctx) {
                Ok(s) => s,
                Err(e) => return Value::Error(e),
            };

            match crate::functions::text::replace(&old, start, num, &new) {
                Ok(s) => Value::Text(s),
                Err(e) => Value::Error(excel_error_to_kind(e)),
            }
        },
    )
}

fn slice_chars(text: &str, start: usize, len: usize) -> String {
    if len == 0 {
        return String::new();
    }

    let end_char = start.saturating_add(len);
    let mut start_byte: Option<usize> = (start == 0).then_some(0);
    let mut end_byte: Option<usize> = None;
    let mut ci = 0usize;
    for (i, _) in text.char_indices() {
        if ci == start {
            start_byte = Some(i);
        }
        if ci == end_char {
            end_byte = Some(i);
            break;
        }
        ci += 1;
    }

    let Some(start_byte) = start_byte else {
        return String::new();
    };
    let end_byte = end_byte.unwrap_or(text.len());
    text[start_byte..end_byte].to_string()
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
    out
}

fn find_impl(needle: &str, haystack: &str, start: i64, case_insensitive: bool) -> Value {
    if start < 1 {
        return Value::Error(ErrorKind::Value);
    }
    let start_idx = (start - 1) as usize;

    if case_insensitive {
        let hay_ascii_needs_uppercasing = haystack
            .as_bytes()
            .iter()
            .any(|b| b.is_ascii_lowercase());
        let (hay_folded, folded_starts) =
            crate::value::fold_str_to_uppercase_with_starts(haystack, hay_ascii_needs_uppercasing);
        let hay_len_chars = folded_starts.len();

        if start_idx > hay_len_chars {
            return Value::Error(ErrorKind::Value);
        }
        if needle.is_empty() {
            return Value::Number(start as f64);
        }
        if start_idx == hay_len_chars {
            return Value::Error(ErrorKind::Value);
        }

        // Excel SEARCH is case-insensitive using Unicode-aware uppercasing (e.g. ß -> SS).
        // Fold the pattern into a comparable char stream while parsing (avoids an intermediate
        // `Vec<char>` allocation for the folded pattern).
        let needle_tokens = parse_search_pattern_folded(needle);
        let min_required = min_required_hay_len(&needle_tokens);

        if let [PatternToken::LiteralSeq(seq)] = needle_tokens.as_slice() {
            for orig_idx in start_idx..hay_len_chars {
                let folded_idx = folded_starts[orig_idx];
                if hay_folded.len().saturating_sub(folded_idx) < seq.len() {
                    break;
                }
                if hay_folded[folded_idx..].starts_with(seq) {
                    return Value::Number((orig_idx + 1) as f64);
                }
            }
            return Value::Error(ErrorKind::Value);
        }

        let stride = match hay_folded.len().checked_add(1) {
            Some(v) => v,
            None => return Value::Error(ErrorKind::Num),
        };
        let token_count = match needle_tokens.len().checked_add(1) {
            Some(v) => v,
            None => return Value::Error(ErrorKind::Num),
        };
        let memo_len = match token_count.checked_mul(stride) {
            Some(v) => v,
            None => return Value::Error(ErrorKind::Num),
        };
        let mut memo: Vec<Option<bool>> = Vec::new();
        if memo.try_reserve_exact(memo_len).is_err() {
            debug_assert!(false, "allocation failed (search memo, len={memo_len})");
            return Value::Error(ErrorKind::Num);
        }
        memo.resize(memo_len, None);
        for orig_idx in start_idx..hay_len_chars {
            let folded_idx = folded_starts[orig_idx];
            if hay_folded.len().saturating_sub(folded_idx) < min_required {
                break;
            }
            if matches_pattern_with_memo(&needle_tokens, &hay_folded, folded_idx, stride, &mut memo)
            {
                return Value::Number((orig_idx + 1) as f64);
            }
        }
        Value::Error(ErrorKind::Value)
    } else {
        if haystack.is_ascii() {
            let hay_len = haystack.len();
            if start_idx > hay_len {
                return Value::Error(ErrorKind::Value);
            }
            if needle.is_empty() {
                return Value::Number(start as f64);
            }
            if start_idx == hay_len {
                return Value::Error(ErrorKind::Value);
            }
            if let Some(rel) = haystack[start_idx..].find(needle) {
                return Value::Number((start_idx + rel + 1) as f64);
            }
            return Value::Error(ErrorKind::Value);
        }

        let mut hay_len_chars = 0usize;
        let mut start_byte: Option<usize> = None;
        for (i, _) in haystack.char_indices() {
            if hay_len_chars == start_idx {
                start_byte = Some(i);
            }
            hay_len_chars += 1;
        }

        if start_idx > hay_len_chars {
            return Value::Error(ErrorKind::Value);
        }
        if needle.is_empty() {
            return Value::Number(start as f64);
        }
        if start_idx == hay_len_chars {
            return Value::Error(ErrorKind::Value);
        }

        let Some(start_byte) = start_byte else {
            debug_assert!(false, "start_idx < hay_len_chars implies a start byte");
            return Value::Error(ErrorKind::Value);
        };
        let hay = &haystack[start_byte..];
        let Some(rel) = hay.find(needle) else {
            return Value::Error(ErrorKind::Value);
        };

        let prefix = &hay[..rel];
        let prefix_chars = prefix.chars().count();
        Value::Number((start_idx + prefix_chars + 1) as f64)
    }
}

#[derive(Debug, Clone)]
enum MatrixArg {
    Scalar(Value),
    Array(Array),
}

impl MatrixArg {
    fn dims(&self) -> (usize, usize) {
        match self {
            MatrixArg::Scalar(_) => (1, 1),
            MatrixArg::Array(arr) => (arr.rows, arr.cols),
        }
    }

    fn get(&self, row: usize, col: usize) -> &Value {
        static FALLBACK_ERROR: Value = Value::Error(ErrorKind::Value);

        match self {
            MatrixArg::Scalar(v) => v,
            MatrixArg::Array(arr) => {
                if arr.rows == 1 && arr.cols == 1 {
                    match arr.get(0, 0) {
                        Some(v) => v,
                        None => {
                            debug_assert!(false, "1x1 arrays should have top-left");
                            &FALLBACK_ERROR
                        }
                    }
                } else {
                    match arr.get(row, col) {
                        Some(v) => v,
                        None => {
                            debug_assert!(false, "broadcast shape should ensure in-bounds");
                            &FALLBACK_ERROR
                        }
                    }
                }
            }
        }
    }
}

fn eval_matrix_arg(ctx: &dyn FunctionContext, expr: &CompiledExpr) -> Result<MatrixArg, ErrorKind> {
    match ctx.eval_arg(expr) {
        ArgValue::Scalar(v) => match v {
            Value::Array(arr) => Ok(MatrixArg::Array(arr)),
            other => Ok(MatrixArg::Scalar(other)),
        },
        ArgValue::Reference(r) => {
            let r = r.normalized();
            ctx.record_reference(&r);
            if r.is_single_cell() {
                return Ok(MatrixArg::Scalar(ctx.get_cell_value(&r.sheet_id, r.start)));
            }
            let rows = (r.end.row - r.start.row + 1) as usize;
            let cols = (r.end.col - r.start.col + 1) as usize;
            let total = match rows.checked_mul(cols) {
                Some(v) => v,
                None => return Err(ErrorKind::Spill),
            };
            if total > crate::eval::MAX_MATERIALIZED_ARRAY_CELLS {
                return Err(ErrorKind::Spill);
            }
            let mut values = Vec::new();
            values
                .try_reserve_exact(total)
                .map_err(|_| ErrorKind::Num)?;
            for addr in r.iter_cells() {
                values.push(ctx.get_cell_value(&r.sheet_id, addr));
            }
            Ok(MatrixArg::Array(Array::new(rows, cols, values)))
        }
        ArgValue::ReferenceUnion(_) => Err(ErrorKind::Value),
    }
}

fn broadcast_shape(args: &[&MatrixArg]) -> Result<(usize, usize), ErrorKind> {
    let mut rows = 1usize;
    let mut cols = 1usize;
    for arg in args {
        let (r, c) = arg.dims();
        if r == 1 && c == 1 {
            continue;
        }
        if rows == 1 && cols == 1 {
            rows = r;
            cols = c;
            continue;
        }
        if rows != r || cols != c {
            return Err(ErrorKind::Value);
        }
    }
    Ok((rows, cols))
}

fn elementwise_unary(arg: MatrixArg, f: impl Fn(&Value) -> Value) -> Value {
    match arg {
        MatrixArg::Scalar(v) => f(&v),
        MatrixArg::Array(arr) => {
            let total = match arr.rows.checked_mul(arr.cols) {
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
            for v in arr.iter() {
                values.push(f(v));
            }
            Value::Array(Array::new(arr.rows, arr.cols, values))
        }
    }
}

fn elementwise_binary(
    left: MatrixArg,
    right: MatrixArg,
    f: impl Fn(&Value, &Value) -> Value,
) -> Value {
    let (rows, cols) = match broadcast_shape(&[&left, &right]) {
        Ok(shape) => shape,
        Err(e) => return Value::Error(e),
    };

    let returns_array = matches!(left, MatrixArg::Array(_)) || matches!(right, MatrixArg::Array(_));
    if rows == 1 && cols == 1 {
        let scalar = f(left.get(0, 0), right.get(0, 0));
        return if returns_array {
            Value::Array(Array::new(1, 1, vec![scalar]))
        } else {
            scalar
        };
    }
    let total = match rows.checked_mul(cols) {
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
    for r in 0..rows {
        for c in 0..cols {
            values.push(f(left.get(r, c), right.get(r, c)));
        }
    }
    Value::Array(Array::new(rows, cols, values))
}

fn elementwise_ternary(
    first: MatrixArg,
    second: MatrixArg,
    third: MatrixArg,
    f: impl Fn(&Value, &Value, &Value) -> Value,
) -> Value {
    let (rows, cols) = match broadcast_shape(&[&first, &second, &third]) {
        Ok(shape) => shape,
        Err(e) => return Value::Error(e),
    };

    let returns_array = matches!(first, MatrixArg::Array(_))
        || matches!(second, MatrixArg::Array(_))
        || matches!(third, MatrixArg::Array(_));
    if rows == 1 && cols == 1 {
        let scalar = f(first.get(0, 0), second.get(0, 0), third.get(0, 0));
        return if returns_array {
            Value::Array(Array::new(1, 1, vec![scalar]))
        } else {
            scalar
        };
    }
    let total = match rows.checked_mul(cols) {
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
    for r in 0..rows {
        for c in 0..cols {
            values.push(f(first.get(r, c), second.get(r, c), third.get(r, c)));
        }
    }
    Value::Array(Array::new(rows, cols, values))
}

fn elementwise_quaternary(
    first: MatrixArg,
    second: MatrixArg,
    third: MatrixArg,
    fourth: MatrixArg,
    f: impl Fn(&Value, &Value, &Value, &Value) -> Value,
) -> Value {
    let (rows, cols) = match broadcast_shape(&[&first, &second, &third, &fourth]) {
        Ok(shape) => shape,
        Err(e) => return Value::Error(e),
    };

    let returns_array = matches!(first, MatrixArg::Array(_))
        || matches!(second, MatrixArg::Array(_))
        || matches!(third, MatrixArg::Array(_))
        || matches!(fourth, MatrixArg::Array(_));
    if rows == 1 && cols == 1 {
        let scalar = f(
            first.get(0, 0),
            second.get(0, 0),
            third.get(0, 0),
            fourth.get(0, 0),
        );
        return if returns_array {
            Value::Array(Array::new(1, 1, vec![scalar]))
        } else {
            scalar
        };
    }
    let total = match rows.checked_mul(cols) {
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
    for r in 0..rows {
        for c in 0..cols {
            values.push(f(
                first.get(r, c),
                second.get(r, c),
                third.get(r, c),
                fourth.get(r, c),
            ));
        }
    }
    Value::Array(Array::new(rows, cols, values))
}

fn excel_error_to_kind(err: ExcelError) -> ErrorKind {
    match err {
        ExcelError::Div0 => ErrorKind::Div0,
        ExcelError::Value => ErrorKind::Value,
        ExcelError::Num => ErrorKind::Num,
    }
}

fn coerce_single_char(value: &Value, ctx: &dyn FunctionContext) -> Result<char, ErrorKind> {
    let s = value.coerce_to_string_with_ctx(ctx)?;
    let mut chars = s.chars();
    match (chars.next(), chars.next()) {
        (Some(ch), None) => Ok(ch),
        _ => Err(ErrorKind::Value),
    }
}

fn coerce_optional_single_char(
    value: &Value,
    ctx: &dyn FunctionContext,
) -> Result<Option<char>, ErrorKind> {
    let s = value.coerce_to_string_with_ctx(ctx)?;
    if s.is_empty() {
        return Ok(None);
    }
    let mut chars = s.chars();
    match (chars.next(), chars.next()) {
        (Some(ch), None) => Ok(Some(ch)),
        _ => Err(ErrorKind::Value),
    }
}

fn flatten_textjoin_value(out: &mut Vec<Value>, value: Value) {
    match value {
        Value::Array(arr) => out.extend(arr.values.into_iter()),
        other => out.push(other),
    }
}

fn flatten_textjoin_reference(
    ctx: &dyn FunctionContext,
    out: &mut Vec<Value>,
    reference: crate::functions::Reference,
) {
    let reference = reference.normalized();
    ctx.record_reference(&reference);
    for addr in reference.iter_cells() {
        out.push(ctx.get_cell_value(&reference.sheet_id, addr));
    }
}

fn flatten_textjoin_reference_union(
    ctx: &dyn FunctionContext,
    out: &mut Vec<Value>,
    ranges: &[crate::functions::Reference],
) {
    let mut seen = std::collections::HashSet::new();
    for range in ranges {
        let range = range.normalized();
        ctx.record_reference(&range);
        for addr in range.iter_cells() {
            if !seen.insert((range.sheet_id.clone(), addr)) {
                continue;
            }
            out.push(ctx.get_cell_value(&range.sheet_id, addr));
        }
    }
}

// On wasm targets, `inventory` registrations can be dropped by the linker if the object file
// contains no otherwise-referenced symbols. Referencing this function from a `#[used]` table in
// `functions/mod.rs` ensures the module (and its `inventory::submit!` entries) are retained.
#[cfg(target_arch = "wasm32")]
pub(super) fn __force_link() {}
