use chrono::{NaiveDate, NaiveDateTime};

use crate::locale::{DateOrder, ValueLocaleConfig};

pub(crate) fn parse_text_number(text: &str, value_locale: ValueLocaleConfig) -> Option<f64> {
    // Sort/filter parsing should treat empty/whitespace-only strings as non-numeric.
    if text.trim().is_empty() {
        return None;
    }

    let separators = value_locale.separators;
    // For locales where the thousands separator is `.`, Excel's VALUE-like coercion is careful
    // not to interpret dot-separated dates (e.g. `2020.01.01`) as a number by stripping all
    // separators. Mirror that heuristic here so we don't accidentally treat date-like strings
    // as numeric values during sort/filter type detection.
    if separators.thousands_sep == '.' && text.contains(separators.thousands_sep) {
        if let Some(compact) = compact_for_grouping_validation(text) {
            if !has_valid_thousands_grouping(&compact, separators.decimal_sep, separators.thousands_sep) {
                return None;
            }
        }
    }
    crate::coercion::number::parse_number_strict(
        text,
        separators.decimal_sep,
        Some(separators.thousands_sep),
    )
    .ok()
}

pub(crate) fn parse_text_datetime(text: &str, value_locale: ValueLocaleConfig) -> Option<NaiveDateTime> {
    let s = text.trim();
    if s.is_empty() {
        return None;
    }

    // Common ISO-ish formats first.
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y-%m-%d %H:%M:%S") {
        return Some(dt);
    }
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y-%m-%d %H:%M") {
        return Some(dt);
    }
    if let Ok(date) = NaiveDate::parse_from_str(s, "%Y-%m-%d") {
        return Some(date.and_hms_opt(0, 0, 0)?);
    }

    // Unambiguous year-first numeric dates should be accepted regardless of the locale's primary
    // date order (mirrors `coercion::datetime` behavior).
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y/%m/%d %H:%M:%S") {
        return Some(dt);
    }
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y/%m/%d %H:%M") {
        return Some(dt);
    }
    if let Ok(date) = NaiveDate::parse_from_str(s, "%Y/%m/%d") {
        return Some(date.and_hms_opt(0, 0, 0)?);
    }

    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y.%m.%d %H:%M:%S") {
        return Some(dt);
    }
    if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y.%m.%d %H:%M") {
        return Some(dt);
    }
    if let Ok(date) = NaiveDate::parse_from_str(s, "%Y.%m.%d") {
        return Some(date.and_hms_opt(0, 0, 0)?);
    }

    // Localized numeric dates.
    match value_locale.date_order {
        DateOrder::MDY => {
            if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%m/%d/%Y %H:%M:%S") {
                return Some(dt);
            }
            if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%m/%d/%Y %H:%M") {
                return Some(dt);
            }
            if let Ok(date) = NaiveDate::parse_from_str(s, "%m/%d/%Y") {
                return Some(date.and_hms_opt(0, 0, 0)?);
            }
        }
        DateOrder::DMY => {
            if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%d/%m/%Y %H:%M:%S") {
                return Some(dt);
            }
            if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%d/%m/%Y %H:%M") {
                return Some(dt);
            }
            if let Ok(date) = NaiveDate::parse_from_str(s, "%d/%m/%Y") {
                return Some(date.and_hms_opt(0, 0, 0)?);
            }

            // Many DMY locales use `.` as a date separator.
            if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%d.%m.%Y %H:%M:%S") {
                return Some(dt);
            }
            if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%d.%m.%Y %H:%M") {
                return Some(dt);
            }
            if let Ok(date) = NaiveDate::parse_from_str(s, "%d.%m.%Y") {
                return Some(date.and_hms_opt(0, 0, 0)?);
            }
        }
        DateOrder::YMD => {
            if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y/%m/%d %H:%M:%S") {
                return Some(dt);
            }
            if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y/%m/%d %H:%M") {
                return Some(dt);
            }
            if let Ok(date) = NaiveDate::parse_from_str(s, "%Y/%m/%d") {
                return Some(date.and_hms_opt(0, 0, 0)?);
            }

            if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y.%m.%d %H:%M:%S") {
                return Some(dt);
            }
            if let Ok(dt) = NaiveDateTime::parse_from_str(s, "%Y.%m.%d %H:%M") {
                return Some(dt);
            }
            if let Ok(date) = NaiveDate::parse_from_str(s, "%Y.%m.%d") {
                return Some(date.and_hms_opt(0, 0, 0)?);
            }
        }
    }

    None
}

fn compact_for_grouping_validation(text: &str) -> Option<String> {
    let mut s = text.trim();
    if s.is_empty() {
        return None;
    }

    // Parentheses indicate accounting negative numbers; they're not relevant to grouping.
    if s.starts_with('(') && s.ends_with(')') && s.len() >= 2 {
        s = s[1..s.len() - 1].trim();
    }

    if let Some(rest) = s.strip_prefix('-') {
        s = rest.trim_start();
    } else if let Some(rest) = s.strip_prefix('+') {
        s = rest.trim_start();
    }

    // Strip a small set of common currency symbols (mirrors coercion::number).
    s = s
        .trim_start_matches(|c: char| matches!(c, '$' | '€' | '£' | '¥'))
        .trim();

    // Strip trailing percent signs.
    loop {
        let trimmed = s.trim_end();
        if let Some(rest) = trimmed.strip_suffix('%') {
            s = rest;
            continue;
        }
        s = trimmed;
        break;
    }

    let compact: String = s.chars().filter(|c| !c.is_whitespace()).collect();
    if compact.is_empty() {
        None
    } else {
        Some(compact)
    }
}

fn has_valid_thousands_grouping(compact: &str, decimal_sep: char, group_sep: char) -> bool {
    if group_sep == decimal_sep {
        return false;
    }

    // Ignore exponent part when validating grouping in the mantissa.
    let mantissa = compact
        .split_once('e')
        .map(|(m, _)| m)
        .unwrap_or_else(|| compact.split_once('E').map(|(m, _)| m).unwrap_or(compact));

    if !mantissa.contains(group_sep) {
        return true;
    }

    // Fractional part must not contain grouping separators.
    let (integer, fractional) = mantissa
        .split_once(decimal_sep)
        .map(|(i, f)| (i, Some(f)))
        .unwrap_or((mantissa, None));
    if fractional.is_some_and(|f| f.contains(group_sep)) {
        return false;
    }

    if integer.starts_with(group_sep) || integer.ends_with(group_sep) {
        return false;
    }

    let segments: Vec<&str> = integer.split(group_sep).collect();
    if segments.len() <= 1 {
        return true;
    }

    if segments[0].is_empty() || segments[0].len() > 3 {
        return false;
    }

    for seg in &segments {
        if seg.is_empty() || !seg.chars().all(|c| c.is_ascii_digit()) {
            return false;
        }
    }

    for seg in &segments[1..] {
        if seg.len() != 3 {
            return false;
        }
    }

    true
}
