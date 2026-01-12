use crate::{ErrorKind, Value};
use crate::functions::wildcard::WildcardPattern;
use crate::coercion::number::parse_number_strict;
use std::borrow::Cow;
use std::cmp::Ordering;

fn text_eq_case_insensitive(a: &str, b: &str) -> bool {
    if a.is_ascii() && b.is_ascii() {
        return a.eq_ignore_ascii_case(b);
    }

    a.chars()
        .flat_map(|c| c.to_uppercase())
        .eq(b.chars().flat_map(|c| c.to_uppercase()))
}

fn values_equal_for_lookup(lookup_value: &Value, candidate: &Value) -> bool {
    match (lookup_value, candidate) {
        (Value::Number(a), Value::Number(b)) => a == b,
        (Value::Text(a), Value::Text(b)) => text_eq_case_insensitive(a, b),
        (Value::Bool(a), Value::Bool(b)) => a == b,
        (Value::Error(a), Value::Error(b)) => a == b,
        (Value::Number(a), Value::Text(b)) | (Value::Text(b), Value::Number(a)) => {
            parse_number_strict(b, '.', Some(',')).is_ok_and(|parsed| parsed == *a)
        }
        (Value::Bool(a), Value::Number(b)) | (Value::Number(b), Value::Bool(a)) => {
            (*b == 0.0 && !a) || (*b == 1.0 && *a)
        }
        (Value::Blank, Value::Blank) => true,
        _ => false,
    }
}

fn lookup_cmp_with_equality(a: &Value, b: &Value) -> Ordering {
    if values_equal_for_lookup(a, b) {
        Ordering::Equal
    } else {
        lookup_cmp(a, b)
    }
}

fn cmp_ascii_case_insensitive(a: &str, b: &str) -> Ordering {
    let mut a_iter = a.as_bytes().iter();
    let mut b_iter = b.as_bytes().iter();
    loop {
        match (a_iter.next(), b_iter.next()) {
            (Some(&ac), Some(&bc)) => {
                let ac = ac.to_ascii_uppercase();
                let bc = bc.to_ascii_uppercase();
                match ac.cmp(&bc) {
                    Ordering::Equal => continue,
                    ord => return ord,
                }
            }
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (None, None) => return Ordering::Equal,
        }
    }
}

fn cmp_case_insensitive(a: &str, b: &str) -> Ordering {
    if a.is_ascii() && b.is_ascii() {
        return cmp_ascii_case_insensitive(a, b);
    }

    // Compare using Unicode-aware uppercasing so matches behave like Excel (e.g. ß -> SS).
    // This intentionally uses the same `char::to_uppercase` logic as criteria matching.
    let mut a_iter = a.chars().flat_map(|c| c.to_uppercase());
    let mut b_iter = b.chars().flat_map(|c| c.to_uppercase());
    loop {
        match (a_iter.next(), b_iter.next()) {
            (Some(ac), Some(bc)) => match ac.cmp(&bc) {
                Ordering::Equal => continue,
                ord => return ord,
            },
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (None, None) => return Ordering::Equal,
        }
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

fn lookup_cmp(a: &Value, b: &Value) -> Ordering {
    // Blank coerces to the other type for comparisons (Excel semantics).
    match (a, b) {
        (Value::Blank, Value::Number(y)) => return 0.0_f64.partial_cmp(y).unwrap_or(Ordering::Equal),
        (Value::Number(x), Value::Blank) => return x.partial_cmp(&0.0_f64).unwrap_or(Ordering::Equal),
        (Value::Blank, Value::Bool(y)) => return false.cmp(y),
        (Value::Bool(x), Value::Blank) => return x.cmp(&false),
        (Value::Blank, Value::Text(y)) => return cmp_case_insensitive("", y),
        (Value::Text(x), Value::Blank) => return cmp_case_insensitive(x, ""),
        _ => {}
    }

    fn type_rank(v: &Value) -> u8 {
        match v {
            Value::Number(_) => 0,
            Value::Text(_) => 1,
            Value::Bool(_) => 2,
            Value::Blank => 3,
            Value::Error(_) => 4,
            Value::Reference(_)
            | Value::ReferenceUnion(_)
            | Value::Array(_)
            | Value::Lambda(_)
            | Value::Spill { .. } => 5,
        }
    }

    let ra = type_rank(a);
    let rb = type_rank(b);
    if ra != rb {
        return ra.cmp(&rb);
    }

    match (a, b) {
        (Value::Number(x), Value::Number(y)) => x.partial_cmp(y).unwrap_or(Ordering::Equal),
        (Value::Text(x), Value::Text(y)) => cmp_case_insensitive(x, y),
        (Value::Bool(x), Value::Bool(y)) => x.cmp(y),
        (Value::Blank, Value::Blank) => Ordering::Equal,
        (Value::Error(x), Value::Error(y)) => x.as_code().cmp(y.as_code()),
        _ => Ordering::Equal,
    }
}

fn lower_bound_by(values: &[Value], needle: &Value, cmp: impl Fn(&Value, &Value) -> Ordering) -> usize {
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

fn upper_bound_by(values: &[Value], needle: &Value, cmp: impl Fn(&Value, &Value) -> Ordering) -> usize {
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
    if matches!(lookup_value, Value::Lambda(_)) {
        return Err(ErrorKind::Value);
    }
    let pos = match search_mode {
        SearchMode::FirstToLast => xmatch_linear(lookup_value, lookup_array, match_mode, false)?,
        SearchMode::LastToFirst => xmatch_linear(lookup_value, lookup_array, match_mode, true)?,
        SearchMode::BinaryAscending => xmatch_binary(lookup_value, lookup_array, match_mode, false)?,
        SearchMode::BinaryDescending => xmatch_binary(lookup_value, lookup_array, match_mode, true)?,
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
    if matches!(lookup_value, Value::Lambda(_)) {
        return Err(ErrorKind::Value);
    }
    let pos = match search_mode {
        SearchMode::FirstToLast => xmatch_linear_accessor(lookup_value, len, &mut value_at, match_mode, false)?,
        SearchMode::LastToFirst => xmatch_linear_accessor(lookup_value, len, &mut value_at, match_mode, true)?,
        SearchMode::BinaryAscending => xmatch_binary_accessor(lookup_value, len, &mut value_at, match_mode, false)?,
        SearchMode::BinaryDescending => xmatch_binary_accessor(lookup_value, len, &mut value_at, match_mode, true)?,
    };

    let pos = pos.checked_add(1).unwrap_or(usize::MAX);
    Ok(i32::try_from(pos).unwrap_or(i32::MAX))
}

fn xmatch_linear(
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
                if values_equal_for_lookup(lookup_value, candidate) {
                    return Ok(idx);
                }
            }
            Err(ErrorKind::NA)
        }
        MatchMode::Wildcard => {
            // Excel applies wildcard matching to text patterns.
            let pattern = match lookup_value.coerce_to_string() {
                Ok(s) => s,
                Err(e) => return Err(e),
            };
            let pattern = WildcardPattern::new(&pattern);
            for (idx, candidate) in iter {
                let text = match candidate {
                    Value::Error(_) => continue,
                    Value::Text(s) => Cow::Borrowed(s.as_str()),
                    other => match other.coerce_to_string() {
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
                let ord = lookup_cmp_with_equality(candidate, lookup_value);
                if ord == Ordering::Less || ord == Ordering::Equal {
                    best = match best {
                        None => Some(idx),
                        Some(best_idx) => {
                            let best_val = &lookup_array[best_idx];
                            match lookup_cmp_with_equality(candidate, best_val) {
                                Ordering::Greater => Some(idx),
                                Ordering::Equal => {
                                    // For "next smaller", choose the last occurrence of the winning
                                    // value (insertion point is after duplicates).
                                    if idx > best_idx { Some(idx) } else { Some(best_idx) }
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
                let ord = lookup_cmp_with_equality(candidate, lookup_value);
                if ord == Ordering::Greater || ord == Ordering::Equal {
                    best = match best {
                        None => Some(idx),
                        Some(best_idx) => {
                            let best_val = &lookup_array[best_idx];
                            match lookup_cmp_with_equality(candidate, best_val) {
                                Ordering::Less => Some(idx),
                                Ordering::Equal => {
                                    // For "next larger", choose the first occurrence of the winning
                                    // value (insertion point is before duplicates).
                                    if idx < best_idx { Some(idx) } else { Some(best_idx) }
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
                if values_equal_for_lookup(lookup_value, &candidate) {
                    return Ok(idx);
                }
            }
            Err(ErrorKind::NA)
        }
        MatchMode::Wildcard => {
            let pattern = match lookup_value.coerce_to_string() {
                Ok(s) => s,
                Err(e) => return Err(e),
            };
            let pattern = WildcardPattern::new(&pattern);
            for idx in iter {
                let candidate = value_at(idx);
                let text = match &candidate {
                    Value::Error(_) => continue,
                    Value::Text(s) => Cow::Borrowed(s.as_str()),
                    other => match other.coerce_to_string() {
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
                let ord = lookup_cmp_with_equality(&candidate, lookup_value);
                if ord == Ordering::Less || ord == Ordering::Equal {
                    best = match best {
                        None => Some((idx, candidate)),
                        Some((best_idx, best_val)) => match lookup_cmp_with_equality(&candidate, &best_val) {
                            Ordering::Greater => Some((idx, candidate)),
                            Ordering::Equal => {
                                if idx > best_idx {
                                    Some((idx, candidate))
                                } else {
                                    Some((best_idx, best_val))
                                }
                            }
                            Ordering::Less => Some((best_idx, best_val)),
                        },
                    };
                }
            }
            best.map(|(idx, _)| idx).ok_or(ErrorKind::NA)
        }
        MatchMode::ExactOrNextLarger => {
            let mut best: Option<(usize, Value)> = None;
            for idx in iter {
                let candidate = value_at(idx);
                let ord = lookup_cmp_with_equality(&candidate, lookup_value);
                if ord == Ordering::Greater || ord == Ordering::Equal {
                    best = match best {
                        None => Some((idx, candidate)),
                        Some((best_idx, best_val)) => match lookup_cmp_with_equality(&candidate, &best_val) {
                            Ordering::Less => Some((idx, candidate)),
                            Ordering::Equal => {
                                if idx < best_idx {
                                    Some((idx, candidate))
                                } else {
                                    Some((best_idx, best_val))
                                }
                            }
                            Ordering::Greater => Some((best_idx, best_val)),
                        },
                    };
                }
            }
            best.map(|(idx, _)| idx).ok_or(ErrorKind::NA)
        }
    }
}

fn xmatch_binary(
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

    let cmp = if descending {
        |a: &Value, b: &Value| lookup_cmp(b, a)
    } else {
        lookup_cmp
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
            if lb < lookup_array.len() && values_equal_for_lookup(lookup_value, &lookup_array[lb]) {
                return Ok(lb);
            }
            Err(ErrorKind::NA)
        }
        MatchMode::ExactOrNextLarger => {
            if lb < lookup_array.len() && values_equal_for_lookup(lookup_value, &lookup_array[lb]) {
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

    let cmp = if descending {
        |a: &Value, b: &Value| lookup_cmp(b, a)
    } else {
        lookup_cmp
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
                if values_equal_for_lookup(lookup_value, &candidate) {
                    return Ok(lb);
                }
            }
            Err(ErrorKind::NA)
        }
        MatchMode::ExactOrNextLarger => {
            if lb < len {
                let candidate = value_at(lb);
                if values_equal_for_lookup(lookup_value, &candidate) {
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
    xmatch_with_modes(lookup_value, lookup_array, MatchMode::Exact, SearchMode::FirstToLast)
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn xmatch_matches_numeric_text_via_value_parsing() {
        let array = vec![Value::from("1,234.5")];
        assert_eq!(xmatch(&Value::Number(1234.5), &array).unwrap(), 1);
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
        let array = vec![Value::from("apple"), Value::from("banana"), Value::from("apricot")];
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
