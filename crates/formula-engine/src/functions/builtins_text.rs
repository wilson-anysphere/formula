use crate::error::ExcelError;
use crate::eval::CompiledExpr;
use crate::functions::array_lift;
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
        let chars: Vec<char> = text.chars().collect();
        let len = chars.len();
        let start = len.saturating_sub(n as usize);
        Ok(Value::Text(chars[start..].iter().collect()))
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
        Value::Text(crate::functions::text::clean(&text))
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
        Value::Text(crate::functions::text::proper(&text))
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
    let hay_chars: Vec<char> = haystack.chars().collect();
    let start_idx = (start - 1) as usize;
    if start_idx > hay_chars.len() {
        return Value::Error(ErrorKind::Value);
    }

    if needle_chars.is_empty() {
        return Value::Number(start as f64);
    }

    if case_insensitive {
        // Excel SEARCH is case-insensitive using Unicode-aware uppercasing (e.g. ß -> SS).
        // Fold both pattern and haystack into a comparable char stream.
        let needle_folded: Vec<char> = if needle.is_ascii() {
            let needs_uppercasing = needle.as_bytes().iter().any(|b| b.is_ascii_lowercase());
            let mut out = Vec::with_capacity(needle.len());
            if needs_uppercasing {
                out.extend(needle.chars().map(|c| c.to_ascii_uppercase()));
            } else {
                out.extend(needle.chars());
            }
            out
        } else {
            needle.chars().flat_map(|c| c.to_uppercase()).collect()
        };
        let needle_tokens = parse_search_pattern(&needle_folded);

        let hay_ascii_needs_uppercasing = haystack.is_ascii()
            && haystack
                .as_bytes()
                .iter()
                .any(|b| b.is_ascii_lowercase());
        let mut hay_folded = Vec::with_capacity(hay_chars.len());
        let mut folded_starts = Vec::with_capacity(hay_chars.len());
        for ch in &hay_chars {
            folded_starts.push(hay_folded.len());
            if ch.is_ascii() {
                if hay_ascii_needs_uppercasing {
                    hay_folded.push(ch.to_ascii_uppercase());
                } else {
                    hay_folded.push(*ch);
                }
            } else {
                hay_folded.extend(ch.to_uppercase());
            }
        }

        for orig_idx in start_idx..hay_chars.len() {
            let folded_idx = folded_starts[orig_idx];
            if matches_pattern(&needle_tokens, &hay_folded, folded_idx) {
                return Value::Number((orig_idx + 1) as f64);
            }
        }
        Value::Error(ErrorKind::Value)
    } else {
        let needle_tokens = vec![PatternToken::LiteralSeq(needle_chars)];
        for i in start_idx..hay_chars.len() {
            if matches_pattern(&needle_tokens, &hay_chars, i) {
                return Value::Number((i + 1) as f64);
            }
        }
        Value::Error(ErrorKind::Value)
    }
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
        match self {
            MatrixArg::Scalar(v) => v,
            MatrixArg::Array(arr) => {
                if arr.rows == 1 && arr.cols == 1 {
                    arr.get(0, 0).expect("1x1 arrays have top-left")
                } else {
                    arr.get(row, col)
                        .expect("broadcast shape ensures in-bounds")
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
    match s.chars().collect::<Vec<_>>().as_slice() {
        [ch] => Ok(*ch),
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
    match s.chars().collect::<Vec<_>>().as_slice() {
        [ch] => Ok(Some(*ch)),
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
