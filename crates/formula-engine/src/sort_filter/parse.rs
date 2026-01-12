use chrono::{NaiveDate, NaiveDateTime};

use crate::locale::{DateOrder, ValueLocaleConfig};

pub(crate) fn parse_text_number(text: &str, value_locale: ValueLocaleConfig) -> Option<f64> {
    // Sort/filter parsing should treat empty/whitespace-only strings as non-numeric.
    if text.trim().is_empty() {
        return None;
    }

    let separators = value_locale.separators;
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

