use crate::coercion::datetime::parse_value_text;
use crate::date::ExcelDateSystem;
use crate::functions::wildcard::WildcardPattern;
use crate::locale::ValueLocaleConfig;
use crate::value::format_number_general_with_options;
use crate::{ErrorKind, Value};
use chrono::{DateTime, Utc};
use formula_format::Locale;
use std::borrow::Cow;
use std::cmp::Ordering;

#[derive(Debug, Clone, Copy)]
struct LookupContext {
    value_locale: ValueLocaleConfig,
    date_system: ExcelDateSystem,
    now_utc: DateTime<Utc>,
}

impl LookupContext {
    fn new(
        value_locale: ValueLocaleConfig,
        date_system: ExcelDateSystem,
        now_utc: DateTime<Utc>,
    ) -> Self {
        Self {
            value_locale,
            date_system,
            now_utc,
        }
    }

    fn default() -> Self {
        Self::new(
            ValueLocaleConfig::en_us(),
            ExcelDateSystem::EXCEL_1900,
            Utc::now(),
        )
    }
}

fn coerce_to_string_with_general_options(
    value: &Value,
    locale: Locale,
    date_system: ExcelDateSystem,
) -> Result<String, ErrorKind> {
    match value {
        Value::Text(s) => Ok(s.clone()),
        Value::Entity(entity) => Ok(entity.display.clone()),
        Value::Record(record) => {
            if let Some(display_field) = record.display_field.as_deref() {
                if let Some(value) = record.get_field_case_insensitive(display_field) {
                    return coerce_to_string_with_general_options(&value, locale, date_system);
                }
            }
            Ok(record.display.clone())
        }
        Value::Number(n) => Ok(format_number_general_with_options(*n, locale, date_system)),
        Value::Bool(b) => Ok(if *b { "TRUE" } else { "FALSE" }.to_string()),
        Value::Blank => Ok(String::new()),
        Value::Error(e) => Err(*e),
        Value::Reference(_)
        | Value::ReferenceUnion(_)
        | Value::Array(_)
        | Value::Lambda(_)
        | Value::Spill { .. } => Err(ErrorKind::Value),
    }
}

fn values_equal_for_lookup(ctx: &LookupContext, lookup_value: &Value, candidate: &Value) -> bool {
    // Excel compares text case-insensitively. Rich values (entities/records) behave like text by
    // using their display string; for records, this includes `display_field` semantics.
    fn text_like_str<'a>(ctx: &LookupContext, v: &'a Value) -> Option<Cow<'a, str>> {
        match v {
            Value::Text(s) => Some(Cow::Borrowed(s.as_str())),
            Value::Entity(entity) => Some(Cow::Borrowed(entity.display.as_str())),
            Value::Record(record) => {
                // Fast path: if no display_field is present, we can borrow the stored display.
                if record.display_field.is_none() {
                    return Some(Cow::Borrowed(record.display.as_str()));
                }

                match coerce_to_string_with_general_options(
                    v,
                    ctx.value_locale.separators,
                    ctx.date_system,
                ) {
                    Ok(s) => Some(Cow::Owned(s)),
                    // Be conservative: if display_field coercion fails (e.g. display field points
                    // at an error/reference/etc), fall back to the raw display string.
                    Err(_) => Some(Cow::Borrowed(record.display.as_str())),
                }
            }
            _ => None,
        }
    }

    match (lookup_value, candidate) {
        (Value::Number(a), Value::Number(b)) => a == b,
        (a, b) if text_like_str(ctx, a).is_some() && text_like_str(ctx, b).is_some() => {
            let a = text_like_str(ctx, a).unwrap();
            let b = text_like_str(ctx, b).unwrap();
            crate::value::eq_case_insensitive(a.as_ref(), b.as_ref())
        }
        (Value::Bool(a), Value::Bool(b)) => a == b,
        (Value::Error(a), Value::Error(b)) => a == b,
        (Value::Number(a), b) | (b, Value::Number(a)) if text_like_str(ctx, b).is_some() => {
            let b = text_like_str(ctx, b).unwrap();
            let trimmed = b.trim();
            if trimmed.is_empty() {
                false
            } else {
                parse_value_text(trimmed, ctx.value_locale, ctx.now_utc, ctx.date_system)
                    .is_ok_and(|parsed| parsed == *a)
            }
        }
        (Value::Bool(a), Value::Number(b)) | (Value::Number(b), Value::Bool(a)) => {
            (*b == 0.0 && !a) || (*b == 1.0 && *a)
        }
        (Value::Blank, Value::Blank) => true,
        _ => false,
    }
}

fn lookup_cmp_with_equality(ctx: &LookupContext, a: &Value, b: &Value) -> Ordering {
    if values_equal_for_lookup(ctx, a, b) {
        Ordering::Equal
    } else {
        lookup_cmp(ctx, a, b)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MatchMode {
    /// 0: exact match
    Exact,
    /// -1: exact match or next smaller item
    ExactOrNextSmaller,
    /// 1: exact match or next larger item
    ExactOrNextLarger,
    /// 2: wildcard match (text patterns)
    Wildcard,
}

impl TryFrom<i64> for MatchMode {
    type Error = ErrorKind;

    fn try_from(value: i64) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(MatchMode::Exact),
            -1 => Ok(MatchMode::ExactOrNextSmaller),
            1 => Ok(MatchMode::ExactOrNextLarger),
            2 => Ok(MatchMode::Wildcard),
            _ => Err(ErrorKind::Value),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SearchMode {
    /// 1: search first-to-last
    FirstToLast,
    /// -1: search last-to-first
    LastToFirst,
    /// 2: binary search (ascending)
    BinaryAscending,
    /// -2: binary search (descending)
    BinaryDescending,
}

impl TryFrom<i64> for SearchMode {
    type Error = ErrorKind;

    fn try_from(value: i64) -> Result<Self, Self::Error> {
        match value {
            1 => Ok(SearchMode::FirstToLast),
            -1 => Ok(SearchMode::LastToFirst),
            2 => Ok(SearchMode::BinaryAscending),
            -2 => Ok(SearchMode::BinaryDescending),
            _ => Err(ErrorKind::Value),
        }
    }
}

fn lookup_cmp(ctx: &LookupContext, a: &Value, b: &Value) -> Ordering {
    fn text_like_str<'a>(ctx: &LookupContext, v: &'a Value) -> Option<Cow<'a, str>> {
        match v {
            Value::Text(s) => Some(Cow::Borrowed(s.as_str())),
            Value::Entity(entity) => Some(Cow::Borrowed(entity.display.as_str())),
            Value::Record(record) => {
                if record.display_field.is_none() {
                    return Some(Cow::Borrowed(record.display.as_str()));
                }

                match coerce_to_string_with_general_options(
                    v,
                    ctx.value_locale.separators,
                    ctx.date_system,
                ) {
                    Ok(s) => Some(Cow::Owned(s)),
                    Err(_) => Some(Cow::Borrowed(record.display.as_str())),
                }
            }
            _ => None,
        }
    }

    // Blank coerces to the other type for comparisons (Excel semantics).
    match (a, b) {
        (Value::Blank, Value::Number(y)) => {
            return 0.0_f64.partial_cmp(y).unwrap_or(Ordering::Equal)
        }
        (Value::Number(x), Value::Blank) => {
            return x.partial_cmp(&0.0_f64).unwrap_or(Ordering::Equal)
        }
        (Value::Blank, Value::Bool(y)) => return false.cmp(y),
        (Value::Bool(x), Value::Blank) => return x.cmp(&false),
        (Value::Blank, other) => {
            if let Some(other) = text_like_str(ctx, other) {
                return crate::value::cmp_case_insensitive("", other.as_ref());
            }
        }
        (other, Value::Blank) => {
            if let Some(other) = text_like_str(ctx, other) {
                return crate::value::cmp_case_insensitive(other.as_ref(), "");
            }
        }
        _ => {}
    }

    fn type_rank(v: &Value) -> u8 {
        match v {
            Value::Number(_) => 0,
            Value::Text(_) | Value::Entity(_) | Value::Record(_) => 1,
            Value::Bool(_) => 2,
            Value::Blank => 3,
            Value::Error(_) => 4,
            // Non-scalar values aren't comparable for lookup ordering; keep comparisons
            // deterministic by sorting them after all scalar types. Using a catch-all here avoids
            // build breakages if new `Value` variants are introduced.
            _ => 5,
        }
    }

    let ra = type_rank(a);
    let rb = type_rank(b);
    if ra != rb {
        return ra.cmp(&rb);
    }

    match (a, b) {
        (Value::Number(x), Value::Number(y)) => x.partial_cmp(y).unwrap_or(Ordering::Equal),
        (a, b) if text_like_str(ctx, a).is_some() && text_like_str(ctx, b).is_some() => {
            let a = text_like_str(ctx, a).unwrap();
            let b = text_like_str(ctx, b).unwrap();
            crate::value::cmp_case_insensitive(a.as_ref(), b.as_ref())
        }
        (Value::Bool(x), Value::Bool(y)) => x.cmp(y),
        (Value::Blank, Value::Blank) => Ordering::Equal,
        (Value::Error(x), Value::Error(y)) => x.code().cmp(&y.code()),
        _ => Ordering::Equal,
    }
}

fn lower_bound_by(
    values: &[Value],
    needle: &Value,
    cmp: impl Fn(&Value, &Value) -> Ordering,
) -> usize {
    let mut lo = 0usize;
    let mut hi = values.len();
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        match cmp(&values[mid], needle) {
            Ordering::Less => lo = mid + 1,
            Ordering::Equal | Ordering::Greater => hi = mid,
        }
    }
    lo
}

fn upper_bound_by(
    values: &[Value],
    needle: &Value,
    cmp: impl Fn(&Value, &Value) -> Ordering,
) -> usize {
    let mut lo = 0usize;
    let mut hi = values.len();
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        match cmp(&values[mid], needle) {
            Ordering::Greater => hi = mid,
            Ordering::Less | Ordering::Equal => lo = mid + 1,
        }
    }
    lo
}

/// XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])
///
/// Returns a 1-based index on success, or `#N/A` when no match is found.
pub fn xmatch_with_modes(
    lookup_value: &Value,
    lookup_array: &[Value],
    match_mode: MatchMode,
    search_mode: SearchMode,
) -> Result<i32, ErrorKind> {
    let ctx = LookupContext::default();
    xmatch_with_modes_impl(&ctx, lookup_value, lookup_array, match_mode, search_mode)
}

pub(crate) fn xmatch_with_modes_with_locale(
    lookup_value: &Value,
    lookup_array: &[Value],
    match_mode: MatchMode,
    search_mode: SearchMode,
    value_locale: ValueLocaleConfig,
    date_system: ExcelDateSystem,
    now_utc: DateTime<Utc>,
) -> Result<i32, ErrorKind> {
    let ctx = LookupContext::new(value_locale, date_system, now_utc);
    xmatch_with_modes_impl(&ctx, lookup_value, lookup_array, match_mode, search_mode)
}

fn xmatch_with_modes_impl(
    ctx: &LookupContext,
    lookup_value: &Value,
    lookup_array: &[Value],
    match_mode: MatchMode,
    search_mode: SearchMode,
) -> Result<i32, ErrorKind> {
    if matches!(lookup_value, Value::Lambda(_)) {
        return Err(ErrorKind::Value);
    }
    let pos = match search_mode {
        SearchMode::FirstToLast => {
            xmatch_linear(ctx, lookup_value, lookup_array, match_mode, false)?
        }
        SearchMode::LastToFirst => {
            xmatch_linear(ctx, lookup_value, lookup_array, match_mode, true)?
        }
        SearchMode::BinaryAscending => {
            xmatch_binary(ctx, lookup_value, lookup_array, match_mode, false)?
        }
        SearchMode::BinaryDescending => {
            xmatch_binary(ctx, lookup_value, lookup_array, match_mode, true)?
        }
    };

    let pos = pos.checked_add(1).unwrap_or(usize::MAX);
    Ok(i32::try_from(pos).unwrap_or(i32::MAX))
}

/// Like [`xmatch_with_modes`], but allows providing the lookup array via a random-access accessor.
///
/// This is used to avoid materializing large reference ranges into a `Vec<Value>` when the caller
/// can compute element values on-demand (e.g. via worksheet cell lookup).
pub fn xmatch_with_modes_accessor(
    lookup_value: &Value,
    len: usize,
    mut value_at: impl FnMut(usize) -> Value,
    match_mode: MatchMode,
    search_mode: SearchMode,
) -> Result<i32, ErrorKind> {
    let ctx = LookupContext::default();
    xmatch_with_modes_accessor_impl(
        &ctx,
        lookup_value,
        len,
        &mut value_at,
        match_mode,
        search_mode,
    )
}

pub(crate) fn xmatch_with_modes_accessor_with_locale(
    lookup_value: &Value,
    len: usize,
    mut value_at: impl FnMut(usize) -> Value,
    match_mode: MatchMode,
    search_mode: SearchMode,
    value_locale: ValueLocaleConfig,
    date_system: ExcelDateSystem,
    now_utc: DateTime<Utc>,
) -> Result<i32, ErrorKind> {
    let ctx = LookupContext::new(value_locale, date_system, now_utc);
    xmatch_with_modes_accessor_impl(
        &ctx,
        lookup_value,
        len,
        &mut value_at,
        match_mode,
        search_mode,
    )
}

fn xmatch_with_modes_accessor_impl(
    ctx: &LookupContext,
    lookup_value: &Value,
    len: usize,
    value_at: &mut impl FnMut(usize) -> Value,
    match_mode: MatchMode,
    search_mode: SearchMode,
) -> Result<i32, ErrorKind> {
    if matches!(lookup_value, Value::Lambda(_)) {
        return Err(ErrorKind::Value);
    }
    let pos = match search_mode {
        SearchMode::FirstToLast => {
            xmatch_linear_accessor(ctx, lookup_value, len, value_at, match_mode, false)?
        }
        SearchMode::LastToFirst => {
            xmatch_linear_accessor(ctx, lookup_value, len, value_at, match_mode, true)?
        }
        SearchMode::BinaryAscending => {
            xmatch_binary_accessor(ctx, lookup_value, len, value_at, match_mode, false)?
        }
        SearchMode::BinaryDescending => {
            xmatch_binary_accessor(ctx, lookup_value, len, value_at, match_mode, true)?
        }
    };

    let pos = pos.checked_add(1).unwrap_or(usize::MAX);
    Ok(i32::try_from(pos).unwrap_or(i32::MAX))
}

fn xmatch_linear(
    ctx: &LookupContext,
    lookup_value: &Value,
    lookup_array: &[Value],
    match_mode: MatchMode,
    reverse: bool,
) -> Result<usize, ErrorKind> {
    let iter: Box<dyn Iterator<Item = (usize, &Value)>> = if reverse {
        Box::new(lookup_array.iter().enumerate().rev())
    } else {
        Box::new(lookup_array.iter().enumerate())
    };

    match match_mode {
        MatchMode::Exact => {
            for (idx, candidate) in iter {
                if values_equal_for_lookup(ctx, lookup_value, candidate) {
                    return Ok(idx);
                }
            }
            Err(ErrorKind::NA)
        }
        MatchMode::Wildcard => {
            // Excel applies wildcard matching to text patterns.
            let pattern = match coerce_to_string_with_general_options(
                lookup_value,
                ctx.value_locale.separators,
                ctx.date_system,
            ) {
                Ok(s) => s,
                Err(e) => return Err(e),
            };
            let pattern = WildcardPattern::new(&pattern);
            for (idx, candidate) in iter {
                let text = match candidate {
                    Value::Error(_) => continue,
                    Value::Text(s) => Cow::Borrowed(s.as_str()),
                    other => match coerce_to_string_with_general_options(
                        other,
                        ctx.value_locale.separators,
                        ctx.date_system,
                    ) {
                        Ok(s) => Cow::Owned(s),
                        Err(_) => continue,
                    },
                };
                if pattern.matches(text.as_ref()) {
                    return Ok(idx);
                }
            }
            Err(ErrorKind::NA)
        }
        MatchMode::ExactOrNextSmaller => {
            let mut best: Option<usize> = None;
            for (idx, candidate) in iter {
                let ord = lookup_cmp_with_equality(ctx, candidate, lookup_value);
                if ord == Ordering::Less || ord == Ordering::Equal {
                    best = match best {
                        None => Some(idx),
                        Some(best_idx) => {
                            let best_val = &lookup_array[best_idx];
                            match lookup_cmp_with_equality(ctx, candidate, best_val) {
                                Ordering::Greater => Some(idx),
                                Ordering::Equal => {
                                    // For "next smaller", choose the last occurrence of the winning
                                    // value (insertion point is after duplicates).
                                    if idx > best_idx {
                                        Some(idx)
                                    } else {
                                        Some(best_idx)
                                    }
                                }
                                Ordering::Less => Some(best_idx),
                            }
                        }
                    };
                }
            }
            best.ok_or(ErrorKind::NA)
        }
        MatchMode::ExactOrNextLarger => {
            let mut best: Option<usize> = None;
            for (idx, candidate) in iter {
                let ord = lookup_cmp_with_equality(ctx, candidate, lookup_value);
                if ord == Ordering::Greater || ord == Ordering::Equal {
                    best = match best {
                        None => Some(idx),
                        Some(best_idx) => {
                            let best_val = &lookup_array[best_idx];
                            match lookup_cmp_with_equality(ctx, candidate, best_val) {
                                Ordering::Less => Some(idx),
                                Ordering::Equal => {
                                    // For "next larger", choose the first occurrence of the winning
                                    // value (insertion point is before duplicates).
                                    if idx < best_idx {
                                        Some(idx)
                                    } else {
                                        Some(best_idx)
                                    }
                                }
                                Ordering::Greater => Some(best_idx),
                            }
                        }
                    };
                }
            }
            best.ok_or(ErrorKind::NA)
        }
    }
}

fn xmatch_linear_accessor(
    ctx: &LookupContext,
    lookup_value: &Value,
    len: usize,
    value_at: &mut impl FnMut(usize) -> Value,
    match_mode: MatchMode,
    reverse: bool,
) -> Result<usize, ErrorKind> {
    let iter: Box<dyn Iterator<Item = usize>> = if reverse {
        Box::new((0..len).rev())
    } else {
        Box::new(0..len)
    };

    match match_mode {
        MatchMode::Exact => {
            for idx in iter {
                let candidate = value_at(idx);
                if values_equal_for_lookup(ctx, lookup_value, &candidate) {
                    return Ok(idx);
                }
            }
            Err(ErrorKind::NA)
        }
        MatchMode::Wildcard => {
            let pattern = match coerce_to_string_with_general_options(
                lookup_value,
                ctx.value_locale.separators,
                ctx.date_system,
            ) {
                Ok(s) => s,
                Err(e) => return Err(e),
            };
            let pattern = WildcardPattern::new(&pattern);
            for idx in iter {
                let candidate = value_at(idx);
                let text = match &candidate {
                    Value::Error(_) => continue,
                    Value::Text(s) => Cow::Borrowed(s.as_str()),
                    other => match coerce_to_string_with_general_options(
                        other,
                        ctx.value_locale.separators,
                        ctx.date_system,
                    ) {
                        Ok(s) => Cow::Owned(s),
                        Err(_) => continue,
                    },
                };
                if pattern.matches(text.as_ref()) {
                    return Ok(idx);
                }
            }
            Err(ErrorKind::NA)
        }
        MatchMode::ExactOrNextSmaller => {
            let mut best: Option<(usize, Value)> = None;
            for idx in iter {
                let candidate = value_at(idx);
                let ord = lookup_cmp_with_equality(ctx, &candidate, lookup_value);
                if ord == Ordering::Less || ord == Ordering::Equal {
                    best = match best {
                        None => Some((idx, candidate)),
                        Some((best_idx, best_val)) => {
                            match lookup_cmp_with_equality(ctx, &candidate, &best_val) {
                                Ordering::Greater => Some((idx, candidate)),
                                Ordering::Equal => {
                                    if idx > best_idx {
                                        Some((idx, candidate))
                                    } else {
                                        Some((best_idx, best_val))
                                    }
                                }
                                Ordering::Less => Some((best_idx, best_val)),
                            }
                        }
                    };
                }
            }
            best.map(|(idx, _)| idx).ok_or(ErrorKind::NA)
        }
        MatchMode::ExactOrNextLarger => {
            let mut best: Option<(usize, Value)> = None;
            for idx in iter {
                let candidate = value_at(idx);
                let ord = lookup_cmp_with_equality(ctx, &candidate, lookup_value);
                if ord == Ordering::Greater || ord == Ordering::Equal {
                    best = match best {
                        None => Some((idx, candidate)),
                        Some((best_idx, best_val)) => {
                            match lookup_cmp_with_equality(ctx, &candidate, &best_val) {
                                Ordering::Less => Some((idx, candidate)),
                                Ordering::Equal => {
                                    if idx < best_idx {
                                        Some((idx, candidate))
                                    } else {
                                        Some((best_idx, best_val))
                                    }
                                }
                                Ordering::Greater => Some((best_idx, best_val)),
                            }
                        }
                    };
                }
            }
            best.map(|(idx, _)| idx).ok_or(ErrorKind::NA)
        }
    }
}

fn xmatch_binary(
    ctx: &LookupContext,
    lookup_value: &Value,
    lookup_array: &[Value],
    match_mode: MatchMode,
    descending: bool,
) -> Result<usize, ErrorKind> {
    if lookup_array.is_empty() {
        return Err(ErrorKind::NA);
    }

    if matches!(match_mode, MatchMode::Wildcard) {
        return Err(ErrorKind::Value);
    }

    let cmp = |a: &Value, b: &Value| {
        if descending {
            lookup_cmp(ctx, b, a)
        } else {
            lookup_cmp(ctx, a, b)
        }
    };

    let effective_mode = if descending {
        match match_mode {
            MatchMode::Exact => MatchMode::Exact,
            MatchMode::ExactOrNextSmaller => MatchMode::ExactOrNextLarger,
            MatchMode::ExactOrNextLarger => MatchMode::ExactOrNextSmaller,
            MatchMode::Wildcard => MatchMode::Wildcard,
        }
    } else {
        match_mode
    };

    let lb = lower_bound_by(lookup_array, lookup_value, cmp);

    match effective_mode {
        MatchMode::Exact => {
            if lb < lookup_array.len()
                && values_equal_for_lookup(ctx, lookup_value, &lookup_array[lb])
            {
                return Ok(lb);
            }
            Err(ErrorKind::NA)
        }
        MatchMode::ExactOrNextLarger => {
            if lb < lookup_array.len()
                && values_equal_for_lookup(ctx, lookup_value, &lookup_array[lb])
            {
                return Ok(lb);
            }
            if lb < lookup_array.len() {
                return Ok(lb);
            }
            Err(ErrorKind::NA)
        }
        MatchMode::ExactOrNextSmaller => {
            let ub = upper_bound_by(lookup_array, lookup_value, cmp);
            if ub == 0 {
                return Err(ErrorKind::NA);
            }
            Ok(ub - 1)
        }
        MatchMode::Wildcard => Err(ErrorKind::Value),
    }
}

fn lower_bound_by_accessor(
    len: usize,
    needle: &Value,
    cmp: impl Fn(&Value, &Value) -> Ordering,
    value_at: &mut impl FnMut(usize) -> Value,
) -> usize {
    let mut lo = 0usize;
    let mut hi = len;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let mid_val = value_at(mid);
        match cmp(&mid_val, needle) {
            Ordering::Less => lo = mid + 1,
            Ordering::Equal | Ordering::Greater => hi = mid,
        }
    }
    lo
}

fn upper_bound_by_accessor(
    len: usize,
    needle: &Value,
    cmp: impl Fn(&Value, &Value) -> Ordering,
    value_at: &mut impl FnMut(usize) -> Value,
) -> usize {
    let mut lo = 0usize;
    let mut hi = len;
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        let mid_val = value_at(mid);
        match cmp(&mid_val, needle) {
            Ordering::Greater => hi = mid,
            Ordering::Less | Ordering::Equal => lo = mid + 1,
        }
    }
    lo
}

fn xmatch_binary_accessor(
    ctx: &LookupContext,
    lookup_value: &Value,
    len: usize,
    value_at: &mut impl FnMut(usize) -> Value,
    match_mode: MatchMode,
    descending: bool,
) -> Result<usize, ErrorKind> {
    if len == 0 {
        return Err(ErrorKind::NA);
    }

    if matches!(match_mode, MatchMode::Wildcard) {
        return Err(ErrorKind::Value);
    }

    let cmp = |a: &Value, b: &Value| {
        if descending {
            lookup_cmp(ctx, b, a)
        } else {
            lookup_cmp(ctx, a, b)
        }
    };

    let effective_mode = if descending {
        match match_mode {
            MatchMode::Exact => MatchMode::Exact,
            MatchMode::ExactOrNextSmaller => MatchMode::ExactOrNextLarger,
            MatchMode::ExactOrNextLarger => MatchMode::ExactOrNextSmaller,
            MatchMode::Wildcard => MatchMode::Wildcard,
        }
    } else {
        match_mode
    };

    let lb = lower_bound_by_accessor(len, lookup_value, cmp, value_at);

    match effective_mode {
        MatchMode::Exact => {
            if lb < len {
                let candidate = value_at(lb);
                if values_equal_for_lookup(ctx, lookup_value, &candidate) {
                    return Ok(lb);
                }
            }
            Err(ErrorKind::NA)
        }
        MatchMode::ExactOrNextLarger => {
            if lb < len {
                let candidate = value_at(lb);
                if values_equal_for_lookup(ctx, lookup_value, &candidate) {
                    return Ok(lb);
                }
                return Ok(lb);
            }
            Err(ErrorKind::NA)
        }
        MatchMode::ExactOrNextSmaller => {
            let ub = upper_bound_by_accessor(len, lookup_value, cmp, value_at);
            if ub == 0 {
                return Err(ErrorKind::NA);
            }
            Ok(ub - 1)
        }
        MatchMode::Wildcard => Err(ErrorKind::Value),
    }
}

/// XMATCH(lookup_value, lookup_array)
///
/// Wrapper for the default mode: exact match, searching first-to-last.
pub fn xmatch(lookup_value: &Value, lookup_array: &[Value]) -> Result<i32, ErrorKind> {
    xmatch_with_modes(
        lookup_value,
        lookup_array,
        MatchMode::Exact,
        SearchMode::FirstToLast,
    )
}

/// XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])
///
/// Implements the most common mode: exact match, searching first-to-last.
pub fn xlookup(
    lookup_value: &Value,
    lookup_array: &[Value],
    return_array: &[Value],
    if_not_found: Option<Value>,
) -> Result<Value, ErrorKind> {
    xlookup_with_modes(
        lookup_value,
        lookup_array,
        return_array,
        if_not_found,
        MatchMode::Exact,
        SearchMode::FirstToLast,
    )
}

/// XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
pub fn xlookup_with_modes(
    lookup_value: &Value,
    lookup_array: &[Value],
    return_array: &[Value],
    if_not_found: Option<Value>,
    match_mode: MatchMode,
    search_mode: SearchMode,
) -> Result<Value, ErrorKind> {
    if lookup_array.len() != return_array.len() {
        return Err(ErrorKind::Value);
    }

    match xmatch_with_modes(lookup_value, lookup_array, match_mode, search_mode) {
        Ok(pos) => {
            let idx = usize::try_from(pos - 1).map_err(|_| ErrorKind::Value)?;
            return_array.get(idx).cloned().ok_or(ErrorKind::Value)
        }
        Err(ErrorKind::NA) => if_not_found.ok_or(ErrorKind::NA),
        Err(e) => Err(e),
    }
}

/// LOOKUP(lookup_value, lookup_vector, [result_vector])
///
/// Implements Excel's legacy LOOKUP vector form:
/// - Always uses approximate matching ("exact or next smaller").
/// - On duplicates, returns the last matching item.
/// - `result_vector` defaults to `lookup_vector`.
///
/// Returns `#N/A` when `lookup_value` is smaller than the smallest item.
pub fn lookup_vector(
    lookup_value: &Value,
    lookup_vector: &[Value],
    result_vector: Option<&[Value]>,
) -> Result<Value, ErrorKind> {
    if matches!(lookup_value, Value::Lambda(_)) {
        return Err(ErrorKind::Value);
    }

    let result_vector = result_vector.unwrap_or(lookup_vector);
    if lookup_vector.len() != result_vector.len() {
        return Err(ErrorKind::Value);
    }

    let pos = xmatch_with_modes(
        lookup_value,
        lookup_vector,
        MatchMode::ExactOrNextSmaller,
        SearchMode::BinaryAscending,
    )?;
    let idx = usize::try_from(pos - 1).map_err(|_| ErrorKind::Value)?;
    result_vector.get(idx).cloned().ok_or(ErrorKind::Value)
}

/// LOOKUP(lookup_value, array)
///
/// Implements Excel's legacy LOOKUP array form:
/// - If `array` has more rows than columns (or is square), search the first column and
///   return the corresponding value from the last column.
/// - Otherwise, search the first row and return the corresponding value from the last row.
pub fn lookup_array(lookup_value: &Value, array: &crate::value::Array) -> Result<Value, ErrorKind> {
    if matches!(lookup_value, Value::Lambda(_)) {
        return Err(ErrorKind::Value);
    }
    if array.rows == 0 || array.cols == 0 {
        return Err(ErrorKind::NA);
    }

    let search_first_col = array.rows >= array.cols;
    let len = if search_first_col {
        array.rows
    } else {
        array.cols
    };
    let last_row = array.rows.saturating_sub(1);
    let last_col = array.cols.saturating_sub(1);

    let pos = xmatch_with_modes_accessor(
        lookup_value,
        len,
        |idx| {
            if search_first_col {
                array.get(idx, 0).cloned().unwrap_or(Value::Blank)
            } else {
                array.get(0, idx).cloned().unwrap_or(Value::Blank)
            }
        },
        MatchMode::ExactOrNextSmaller,
        SearchMode::BinaryAscending,
    )?;
    let idx = usize::try_from(pos - 1).map_err(|_| ErrorKind::Value)?;

    Ok(if search_first_col {
        array.get(idx, last_col).cloned().unwrap_or(Value::Blank)
    } else {
        array.get(last_row, idx).cloned().unwrap_or(Value::Blank)
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::value::{EntityValue, RecordValue};

    #[test]
    fn xmatch_matches_numeric_text_via_value_parsing() {
        let array = vec![Value::from("1,234.5")];
        assert_eq!(xmatch(&Value::Number(1234.5), &array).unwrap(), 1);
    }

    #[test]
    fn xmatch_matches_numeric_entity_display_using_value_locale() {
        // Regression test: number vs entity display parsing must use the workbook value locale.
        let array = vec![Value::Entity(EntityValue::new("1,5"))];
        let pos = xmatch_with_modes_with_locale(
            &Value::Number(1.5),
            &array,
            MatchMode::Exact,
            SearchMode::FirstToLast,
            ValueLocaleConfig::de_de(),
            ExcelDateSystem::EXCEL_1900,
            Utc::now(),
        )
        .unwrap();
        assert_eq!(pos, 1);
    }

    #[test]
    fn xmatch_matches_numeric_record_display_using_value_locale() {
        let array = vec![Value::Record(RecordValue::new("1,5"))];
        let pos = xmatch_with_modes_with_locale(
            &Value::Number(1.5),
            &array,
            MatchMode::Exact,
            SearchMode::FirstToLast,
            ValueLocaleConfig::de_de(),
            ExcelDateSystem::EXCEL_1900,
            Utc::now(),
        )
        .unwrap();
        assert_eq!(pos, 1);
    }

    #[test]
    fn xmatch_matches_date_text_via_value_parsing() {
        let array = vec![Value::from("2020-01-01")];
        let serial = parse_value_text(
            "2020-01-01",
            ValueLocaleConfig::en_us(),
            Utc::now(),
            ExcelDateSystem::EXCEL_1900,
        )
        .unwrap();
        assert_eq!(xmatch(&Value::Number(serial), &array).unwrap(), 1);
    }

    #[test]
    fn xmatch_approximate_modes_use_insertion_points_for_duplicates() {
        let array = vec![
            Value::Number(1.0),
            Value::Number(2.0),
            Value::Number(2.0),
            Value::Number(2.0),
            Value::Number(3.0),
        ];

        assert_eq!(
            xmatch_with_modes(
                &Value::Number(2.0),
                &array,
                MatchMode::ExactOrNextSmaller,
                SearchMode::FirstToLast
            )
            .unwrap(),
            4
        );
        assert_eq!(
            xmatch_with_modes(
                &Value::Number(2.0),
                &array,
                MatchMode::ExactOrNextLarger,
                SearchMode::FirstToLast
            )
            .unwrap(),
            2
        );
        assert_eq!(
            xmatch_with_modes(
                &Value::Number(2.0),
                &array,
                MatchMode::ExactOrNextSmaller,
                SearchMode::BinaryAscending
            )
            .unwrap(),
            4
        );
        assert_eq!(
            xmatch_with_modes(
                &Value::Number(2.0),
                &array,
                MatchMode::ExactOrNextLarger,
                SearchMode::BinaryAscending
            )
            .unwrap(),
            2
        );
    }

    #[test]
    fn xmatch_wildcard_reverse_search_returns_last_match() {
        let array = vec![
            Value::from("apple"),
            Value::from("banana"),
            Value::from("apricot"),
        ];
        assert_eq!(
            xmatch_with_modes(
                &Value::from("a*"),
                &array,
                MatchMode::Wildcard,
                SearchMode::LastToFirst
            )
            .unwrap(),
            3
        );
    }

    #[test]
    fn xmatch_binary_search_rejects_wildcard_mode() {
        let array = vec![Value::from("apple"), Value::from("banana")];
        assert_eq!(
            xmatch_with_modes(
                &Value::from("a*"),
                &array,
                MatchMode::Wildcard,
                SearchMode::BinaryAscending
            )
            .unwrap_err(),
            ErrorKind::Value
        );
    }

    #[test]
    fn xlookup_len_mismatch_is_value_error() {
        let lookup_array = vec![Value::from("A"), Value::from("B")];
        let return_array = vec![Value::Number(1.0)];
        assert_eq!(
            xlookup(&Value::from("A"), &lookup_array, &return_array, None).unwrap_err(),
            ErrorKind::Value
        );
    }
}
