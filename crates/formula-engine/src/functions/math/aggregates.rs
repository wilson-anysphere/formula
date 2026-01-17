use chrono::{DateTime, Utc};

use crate::coercion::ValueLocaleConfig;
use crate::date::ExcelDateSystem;
use crate::functions::math::criteria::Criteria;
use crate::simd;
use crate::value::{parse_number, NumberLocale};
use crate::{ErrorKind, Value};

/// SUMIF(range, criteria, [sum_range])
pub fn sumif(
    criteria_range: &[Value],
    criteria: &Value,
    sum_range: Option<&[Value]>,
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    system: ExcelDateSystem,
) -> Result<f64, ErrorKind> {
    let sum_range = sum_range.unwrap_or(criteria_range);
    if criteria_range.len() != sum_range.len() {
        return Err(ErrorKind::Value);
    }

    if let Value::Error(e) = criteria {
        return Err(*e);
    }
    let criteria = Criteria::parse_with_locale_config(criteria, cfg, now_utc, system)?;
    let mut sum = 0.0;
    for (crit_val, sum_val) in criteria_range.iter().zip(sum_range.iter()) {
        if criteria.matches(crit_val) {
            match sum_val {
                Value::Number(n) => sum += n,
                Value::Error(e) => return Err(*e),
                Value::Lambda(_) => return Err(ErrorKind::Value),
                _ => {}
            }
        }
    }
    Ok(sum)
}

/// SUMIFS(sum_range, criteria_range1, criteria1, ...)
pub fn sumifs(
    sum_range: &[Value],
    criteria_pairs: &[(&[Value], &Value)],
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    system: ExcelDateSystem,
) -> Result<f64, ErrorKind> {
    for (range, _) in criteria_pairs {
        if range.len() != sum_range.len() {
            return Err(ErrorKind::Value);
        }
    }

    let compiled = criteria_pairs
        .iter()
        .map(|(_, crit)| {
            if let Value::Error(e) = *crit {
                return Err(*e);
            }
            Criteria::parse_with_locale_config(*crit, cfg, now_utc, system)
        })
        .collect::<Result<Vec<_>, ErrorKind>>()?;

    let mut sum = 0.0;
    'row: for idx in 0..sum_range.len() {
        for ((range, _), crit) in criteria_pairs.iter().zip(compiled.iter()) {
            if !crit.matches(&range[idx]) {
                continue 'row;
            }
        }

        match &sum_range[idx] {
            Value::Number(n) => sum += n,
            Value::Error(e) => return Err(*e),
            Value::Lambda(_) => return Err(ErrorKind::Value),
            _ => {}
        }
    }

    Ok(sum)
}

/// COUNTIFS(criteria_range1, criteria1, ...)
pub fn countifs(
    criteria_pairs: &[(&[Value], &Value)],
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    system: ExcelDateSystem,
) -> Result<f64, ErrorKind> {
    if criteria_pairs.is_empty() {
        return Ok(0.0);
    }

    let len = criteria_pairs[0].0.len();
    for (range, _) in criteria_pairs {
        if range.len() != len {
            return Err(ErrorKind::Value);
        }
    }

    let compiled = criteria_pairs
        .iter()
        .map(|(_, crit)| {
            if let Value::Error(e) = *crit {
                return Err(*e);
            }
            Criteria::parse_with_locale_config(*crit, cfg, now_utc, system)
        })
        .collect::<Result<Vec<_>, ErrorKind>>()?;

    let mut count = 0u64;
    'row: for idx in 0..len {
        for ((range, _), crit) in criteria_pairs.iter().zip(compiled.iter()) {
            if !crit.matches(&range[idx]) {
                continue 'row;
            }
        }
        count += 1;
    }

    Ok(count as f64)
}

/// AVERAGEIF(range, criteria, [average_range])
pub fn averageif(
    criteria_range: &[Value],
    criteria: &Value,
    average_range: Option<&[Value]>,
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    system: ExcelDateSystem,
) -> Result<f64, ErrorKind> {
    let average_range = average_range.unwrap_or(criteria_range);
    if criteria_range.len() != average_range.len() {
        return Err(ErrorKind::Value);
    }

    if let Value::Error(e) = criteria {
        return Err(*e);
    }
    let criteria = Criteria::parse_with_locale_config(criteria, cfg, now_utc, system)?;
    let mut sum = 0.0;
    let mut count = 0u64;
    for (crit_val, avg_val) in criteria_range.iter().zip(average_range.iter()) {
        if criteria.matches(crit_val) {
            match avg_val {
                Value::Number(n) => {
                    sum += n;
                    count += 1;
                }
                Value::Error(e) => return Err(*e),
                Value::Lambda(_) => return Err(ErrorKind::Value),
                _ => {}
            }
        }
    }

    if count == 0 {
        return Err(ErrorKind::Div0);
    }
    Ok(sum / count as f64)
}

/// AVERAGEIFS(average_range, criteria_range1, criteria1, ...)
pub fn averageifs(
    average_range: &[Value],
    criteria_pairs: &[(&[Value], &Value)],
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    system: ExcelDateSystem,
) -> Result<f64, ErrorKind> {
    for (range, _) in criteria_pairs {
        if range.len() != average_range.len() {
            return Err(ErrorKind::Value);
        }
    }

    let compiled = criteria_pairs
        .iter()
        .map(|(_, crit)| {
            if let Value::Error(e) = *crit {
                return Err(*e);
            }
            Criteria::parse_with_locale_config(*crit, cfg, now_utc, system)
        })
        .collect::<Result<Vec<_>, ErrorKind>>()?;

    let mut sum = 0.0;
    let mut count = 0u64;
    'row: for idx in 0..average_range.len() {
        for ((range, _), crit) in criteria_pairs.iter().zip(compiled.iter()) {
            if !crit.matches(&range[idx]) {
                continue 'row;
            }
        }

        match &average_range[idx] {
            Value::Number(n) => {
                sum += n;
                count += 1;
            }
            Value::Error(e) => return Err(*e),
            Value::Lambda(_) => return Err(ErrorKind::Value),
            _ => {}
        }
    }

    if count == 0 {
        return Err(ErrorKind::Div0);
    }
    Ok(sum / count as f64)
}

/// MAXIFS(max_range, criteria_range1, criteria1, ...)
pub fn maxifs(
    max_range: &[Value],
    criteria_pairs: &[(&[Value], &Value)],
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    system: ExcelDateSystem,
) -> Result<f64, ErrorKind> {
    for (range, _) in criteria_pairs {
        if range.len() != max_range.len() {
            return Err(ErrorKind::Value);
        }
    }

    let compiled = criteria_pairs
        .iter()
        .map(|(_, crit)| {
            if let Value::Error(e) = *crit {
                return Err(*e);
            }
            Criteria::parse_with_locale_config(*crit, cfg, now_utc, system)
        })
        .collect::<Result<Vec<_>, ErrorKind>>()?;

    let mut best: Option<f64> = None;
    'row: for idx in 0..max_range.len() {
        for ((range, _), crit) in criteria_pairs.iter().zip(compiled.iter()) {
            if !crit.matches(&range[idx]) {
                continue 'row;
            }
        }

        match &max_range[idx] {
            Value::Number(n) => best = Some(best.map_or(*n, |b| b.max(*n))),
            Value::Error(e) => return Err(*e),
            Value::Lambda(_) => return Err(ErrorKind::Value),
            _ => {}
        }
    }

    Ok(best.unwrap_or(0.0))
}

/// MINIFS(min_range, criteria_range1, criteria1, ...)
pub fn minifs(
    min_range: &[Value],
    criteria_pairs: &[(&[Value], &Value)],
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    system: ExcelDateSystem,
) -> Result<f64, ErrorKind> {
    for (range, _) in criteria_pairs {
        if range.len() != min_range.len() {
            return Err(ErrorKind::Value);
        }
    }

    let compiled = criteria_pairs
        .iter()
        .map(|(_, crit)| {
            if let Value::Error(e) = *crit {
                return Err(*e);
            }
            Criteria::parse_with_locale_config(*crit, cfg, now_utc, system)
        })
        .collect::<Result<Vec<_>, ErrorKind>>()?;

    let mut best: Option<f64> = None;
    'row: for idx in 0..min_range.len() {
        for ((range, _), crit) in criteria_pairs.iter().zip(compiled.iter()) {
            if !crit.matches(&range[idx]) {
                continue 'row;
            }
        }

        match &min_range[idx] {
            Value::Number(n) => best = Some(best.map_or(*n, |b| b.min(*n))),
            Value::Error(e) => return Err(*e),
            Value::Lambda(_) => return Err(ErrorKind::Value),
            _ => {}
        }
    }

    Ok(best.unwrap_or(0.0))
}

/// SUMPRODUCT(array1, [array2], ...)
pub fn sumproduct(arrays: &[&[Value]], locale: NumberLocale) -> Result<f64, ErrorKind> {
    if arrays.is_empty() {
        return Ok(0.0);
    }

    // Excel broadcasts 1x1 scalars across the other array length.
    let len = arrays.iter().map(|arr| arr.len()).max().unwrap_or(0);
    if len == 0 {
        return Err(ErrorKind::Value);
    }
    for arr in arrays {
        if arr.len() != len && arr.len() != 1 {
            return Err(ErrorKind::Value);
        }
    }

    // Hot path: SUMPRODUCT over two arrays is common, and can be SIMD-accelerated once the
    // Value -> f64 coercions are done.
    if arrays.len() == 2 {
        const BLOCK: usize = 1024;

        let a = arrays[0];
        let b = arrays[1];

        // Both arrays already have matching length: SIMD hot path.
        if a.len() == len && b.len() == len {
            let mut buf_a = [0.0_f64; BLOCK];
            let mut buf_b = [0.0_f64; BLOCK];
            let mut buf_len = 0usize;

            let mut sum = 0.0;
            let mut saw_nan = false;

            for idx in 0..len {
                // Preserve Excel-like error precedence: first array is coerced before the second.
                let xa = coerce_sumproduct_number(&a[idx], locale)?;
                let xb = coerce_sumproduct_number(&b[idx], locale)?;

                if xa.is_nan() || xb.is_nan() {
                    saw_nan = true;
                }

                if !saw_nan {
                    buf_a[buf_len] = xa;
                    buf_b[buf_len] = xb;
                    buf_len += 1;

                    if buf_len == BLOCK {
                        sum += simd::sumproduct_ignore_nan_f64(&buf_a, &buf_b);
                        buf_len = 0;
                    }
                }
            }

            if saw_nan {
                return Ok(f64::NAN);
            }

            if buf_len > 0 {
                sum += simd::sumproduct_ignore_nan_f64(&buf_a[..buf_len], &buf_b[..buf_len]);
            }

            return Ok(sum);
        }

        // Broadcast fast path: one side is a scalar, the other is a longer array.
        if a.len() == 1 && b.len() == len {
            let xa = coerce_sumproduct_number(&a[0], locale)?;
            let mut saw_nan = xa.is_nan();

            let mut buf_a = [0.0_f64; BLOCK];
            let mut buf_b = [0.0_f64; BLOCK];
            let mut buf_len = 0usize;
            let mut sum = 0.0;

            for idx in 0..len {
                let xb = coerce_sumproduct_number(&b[idx], locale)?;

                if saw_nan {
                    continue;
                }

                if xb.is_nan() {
                    saw_nan = true;
                    continue;
                }

                buf_a[buf_len] = xa;
                buf_b[buf_len] = xb;
                buf_len += 1;
                if buf_len == BLOCK {
                    sum += simd::sumproduct_ignore_nan_f64(&buf_a, &buf_b);
                    buf_len = 0;
                }
            }

            if saw_nan {
                return Ok(f64::NAN);
            }
            if buf_len > 0 {
                sum += simd::sumproduct_ignore_nan_f64(&buf_a[..buf_len], &buf_b[..buf_len]);
            }
            return Ok(sum);
        }

        if a.len() == len && b.len() == 1 {
            // Preserve error precedence: for idx=0 we must coerce `a[0]` before `b[0]`.
            let xa0 = coerce_sumproduct_number(&a[0], locale)?;
            let xb = coerce_sumproduct_number(&b[0], locale)?;

            let mut saw_nan = xa0.is_nan() || xb.is_nan();

            let mut buf_a = [0.0_f64; BLOCK];
            let mut buf_b = [0.0_f64; BLOCK];
            let mut buf_len = 0usize;
            let mut sum = 0.0;

            if !saw_nan {
                buf_a[buf_len] = xa0;
                buf_b[buf_len] = xb;
                buf_len += 1;
            }

            for idx in 1..len {
                let xa = coerce_sumproduct_number(&a[idx], locale)?;

                if saw_nan {
                    continue;
                }

                if xa.is_nan() {
                    saw_nan = true;
                    continue;
                }

                buf_a[buf_len] = xa;
                buf_b[buf_len] = xb;
                buf_len += 1;
                if buf_len == BLOCK {
                    sum += simd::sumproduct_ignore_nan_f64(&buf_a, &buf_b);
                    buf_len = 0;
                }
            }

            if saw_nan {
                return Ok(f64::NAN);
            }
            if buf_len > 0 {
                sum += simd::sumproduct_ignore_nan_f64(&buf_a[..buf_len], &buf_b[..buf_len]);
            }
            return Ok(sum);
        }

        debug_assert!(false, "broadcast validation should have handled all length combinations");
        return Err(ErrorKind::Value);
    }

    let mut sum = 0.0;
    for idx in 0..len {
        let mut prod = 1.0;
        for arr in arrays {
            let v = if arr.len() == 1 { &arr[0] } else { &arr[idx] };
            let n = coerce_sumproduct_number(v, locale)?;
            prod *= n;
        }
        sum += prod;
    }

    Ok(sum)
}

pub(crate) fn coerce_sumproduct_number(
    value: &Value,
    locale: NumberLocale,
) -> Result<f64, ErrorKind> {
    match value {
        Value::Number(n) => Ok(*n),
        Value::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
        Value::Text(s) => match parse_number(s, locale) {
            Ok(n) => Ok(n),
            Err(crate::error::ExcelError::Value) => Ok(0.0),
            Err(crate::error::ExcelError::Div0) => Err(ErrorKind::Div0),
            Err(crate::error::ExcelError::Num) => Err(ErrorKind::Num),
        },
        Value::Entity(entity) => match parse_number(entity.display.as_str(), locale) {
            Ok(n) => Ok(n),
            Err(crate::error::ExcelError::Value) => Ok(0.0),
            Err(crate::error::ExcelError::Div0) => Err(ErrorKind::Div0),
            Err(crate::error::ExcelError::Num) => Err(ErrorKind::Num),
        },
        Value::Record(record) => match parse_number(record.display.as_str(), locale) {
            Ok(n) => Ok(n),
            Err(crate::error::ExcelError::Value) => Ok(0.0),
            Err(crate::error::ExcelError::Div0) => Err(ErrorKind::Div0),
            Err(crate::error::ExcelError::Num) => Err(ErrorKind::Num),
        },
        Value::Blank => Ok(0.0),
        Value::Error(e) => Err(*e),
        Value::Lambda(_) => Err(ErrorKind::Value),
        Value::Reference(_) | Value::ReferenceUnion(_) | Value::Array(_) | Value::Spill { .. } => {
            Ok(0.0)
        }
    }
}

/// SUBTOTAL(function_num, ref1, [ref2], ...)
///
/// This implements the common `function_num` set (1-11 / 101-111). Hidden rows
/// / filtered ranges are handled by the caller (range iterator).
pub fn subtotal(function_num: i32, values: &[Value]) -> Result<f64, ErrorKind> {
    let base = if function_num >= 100 {
        function_num - 100
    } else {
        function_num
    };

    match base {
        1 => average(values, false),
        2 => Ok(count(values) as f64),
        3 => Ok(counta(values) as f64),
        4 => max(values, false),
        5 => min(values, false),
        6 => product(values, false),
        7 => stdev(values, false),
        8 => stdevp(values, false),
        9 => sum(values, false),
        10 => var(values, false),
        11 => varp(values, false),
        _ => Err(ErrorKind::Value),
    }
}

/// AGGREGATE(function_num, options, ref1, [ref2])
///
/// This intentionally implements the most common aggregation subtypes (1-11).
/// `options` only controls whether errors are ignored.
pub fn aggregate(function_num: i32, options: i32, values: &[Value]) -> Result<f64, ErrorKind> {
    let ignore_errors = matches!(options, 2 | 3 | 6 | 7);
    match function_num {
        1 => average(values, ignore_errors),
        2 => Ok(count(values) as f64),
        3 => Ok(counta_with_errors(values, ignore_errors) as f64),
        4 => max(values, ignore_errors),
        5 => min(values, ignore_errors),
        6 => product(values, ignore_errors),
        7 => stdev(values, ignore_errors),
        8 => stdevp(values, ignore_errors),
        9 => sum(values, ignore_errors),
        10 => var(values, ignore_errors),
        11 => varp(values, ignore_errors),
        _ => Err(ErrorKind::Value),
    }
}

fn sum(values: &[Value], ignore_errors: bool) -> Result<f64, ErrorKind> {
    let mut out = 0.0;
    for value in values {
        match value {
            Value::Number(n) => out += n,
            Value::Error(e) if !ignore_errors => return Err(*e),
            Value::Lambda(_) if !ignore_errors => return Err(ErrorKind::Value),
            _ => {}
        }
    }
    Ok(out)
}

fn average(values: &[Value], ignore_errors: bool) -> Result<f64, ErrorKind> {
    let mut sum = 0.0;
    let mut count = 0usize;
    for value in values {
        match value {
            Value::Number(n) => {
                sum += n;
                count += 1;
            }
            Value::Error(e) if !ignore_errors => return Err(*e),
            Value::Lambda(_) if !ignore_errors => return Err(ErrorKind::Value),
            _ => {}
        }
    }

    if count == 0 {
        return Err(ErrorKind::Div0);
    }
    Ok(sum / count as f64)
}

fn count(values: &[Value]) -> usize {
    values
        .iter()
        .filter(|v| matches!(v, Value::Number(_)))
        .count()
}

fn counta(values: &[Value]) -> usize {
    values.iter().filter(|v| !matches!(v, Value::Blank)).count()
}

fn counta_with_errors(values: &[Value], ignore_errors: bool) -> usize {
    values
        .iter()
        .filter(|v| match v {
            Value::Blank => false,
            Value::Error(_) => !ignore_errors,
            _ => true,
        })
        .count()
}

fn max(values: &[Value], ignore_errors: bool) -> Result<f64, ErrorKind> {
    let mut best: Option<f64> = None;
    for value in values {
        match value {
            Value::Number(n) => best = Some(best.map_or(*n, |b| b.max(*n))),
            Value::Error(e) if !ignore_errors => return Err(*e),
            Value::Lambda(_) if !ignore_errors => return Err(ErrorKind::Value),
            _ => {}
        }
    }
    Ok(best.unwrap_or(0.0))
}

fn min(values: &[Value], ignore_errors: bool) -> Result<f64, ErrorKind> {
    let mut best: Option<f64> = None;
    for value in values {
        match value {
            Value::Number(n) => best = Some(best.map_or(*n, |b| b.min(*n))),
            Value::Error(e) if !ignore_errors => return Err(*e),
            Value::Lambda(_) if !ignore_errors => return Err(ErrorKind::Value),
            _ => {}
        }
    }
    Ok(best.unwrap_or(0.0))
}

fn product(values: &[Value], ignore_errors: bool) -> Result<f64, ErrorKind> {
    let mut out = 1.0;
    let mut saw_number = false;
    for value in values {
        match value {
            Value::Number(n) => {
                saw_number = true;
                out *= n;
            }
            Value::Error(e) if !ignore_errors => return Err(*e),
            Value::Lambda(_) if !ignore_errors => return Err(ErrorKind::Value),
            _ => {}
        }
    }
    if !saw_number {
        return Ok(1.0);
    }
    if out.is_finite() {
        Ok(out)
    } else {
        Err(ErrorKind::Num)
    }
}

fn variance(values: &[Value], ignore_errors: bool) -> Result<(usize, f64, f64), ErrorKind> {
    let mut nums = Vec::new();
    for value in values {
        match value {
            Value::Number(n) => nums.push(*n),
            Value::Error(e) if !ignore_errors => return Err(*e),
            Value::Lambda(_) if !ignore_errors => return Err(ErrorKind::Value),
            _ => {}
        }
    }

    let n = nums.len();
    if n == 0 {
        return Err(ErrorKind::Div0);
    }
    let mean = nums.iter().sum::<f64>() / n as f64;
    let sse = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>();
    Ok((n, mean, sse))
}

fn var(values: &[Value], ignore_errors: bool) -> Result<f64, ErrorKind> {
    let (n, _, sse) = variance(values, ignore_errors)?;
    if n < 2 {
        return Err(ErrorKind::Div0);
    }
    Ok(sse / (n as f64 - 1.0))
}

fn varp(values: &[Value], ignore_errors: bool) -> Result<f64, ErrorKind> {
    let (n, _, sse) = variance(values, ignore_errors)?;
    Ok(sse / n as f64)
}

fn stdev(values: &[Value], ignore_errors: bool) -> Result<f64, ErrorKind> {
    Ok(var(values, ignore_errors)?.sqrt())
}

fn stdevp(values: &[Value], ignore_errors: bool) -> Result<f64, ErrorKind> {
    Ok(varp(values, ignore_errors)?.sqrt())
}
