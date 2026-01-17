use chrono::{DateTime, Datelike, Utc};

use std::borrow::Cow;

use crate::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use crate::error::{ExcelError, ExcelResult};

use super::{DateOrder, ValueLocaleConfig};

pub fn parse_datevalue_text(
    text: &str,
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
    match try_parse_date(text, cfg, now_utc, system) {
        Some(result) => result,
        None => Err(ExcelError::Value),
    }
}

pub fn parse_timevalue_text(text: &str, _cfg: ValueLocaleConfig) -> ExcelResult<f64> {
    match try_parse_time(text) {
        Some(result) => result,
        None => Err(ExcelError::Value),
    }
}

pub fn parse_value_text(
    text: &str,
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    let decimal_sep = cfg.separators.decimal_sep;
    let group_sep = cfg.separators.thousands_sep;

    // Excel's VALUE is eager to treat many strings as dates/times. For locales where the
    // thousands separator is `.`, we need to be careful not to interpret dot-separated dates
    // (e.g. `2020.01.01`) as a number by stripping all separators. A pragmatic approximation is
    // to reject invalid thousands grouping (groups must be 3 digits) before attempting numeric
    // parsing.
    let mut try_number = true;
    if group_sep == '.' && text.contains(group_sep) {
        if let Some(compact) = compact_for_grouping_validation(text) {
            if !has_valid_thousands_grouping(compact.as_ref(), decimal_sep, group_sep) {
                try_number = false;
            }
        }
    }

    if try_number {
        match super::number::parse_number_strict(text, decimal_sep, Some(group_sep)) {
            Ok(n) => return Ok(n),
            Err(ExcelError::Num) => return Err(ExcelError::Num),
            Err(ExcelError::Div0) => return Err(ExcelError::Div0),
            Err(ExcelError::Value) => {}
        }
    }

    let date = match try_parse_date(text, cfg, now_utc, system) {
        Some(Ok(serial)) => Some(serial as f64),
        Some(Err(e)) => return Err(e),
        None => None,
    };

    let time = match try_parse_time(text) {
        Some(Ok(fraction)) => Some(fraction),
        Some(Err(e)) => return Err(e),
        None => None,
    };

    match (date, time) {
        (Some(d), Some(t)) => Ok(d + t),
        (Some(d), None) => Ok(d),
        (None, Some(t)) => Ok(t),
        (None, None) => Err(ExcelError::Value),
    }
}

fn compact_for_grouping_validation(text: &str) -> Option<Cow<'_, str>> {
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

    if s.is_empty() {
        return None;
    }

    let has_whitespace = if s.is_ascii() {
        s.as_bytes().iter().any(|b| b.is_ascii_whitespace())
    } else {
        s.chars().any(|c| c.is_whitespace())
    };
    if !has_whitespace {
        return Some(Cow::Borrowed(s));
    }

    let compact: String = s.chars().filter(|c| !c.is_whitespace()).collect();
    if compact.is_empty() {
        return None;
    }

    Some(Cow::Owned(compact))
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

    let mut segs = integer.split(group_sep);
    let Some(first) = segs.next() else {
        return true;
    };
    let Some(second) = segs.next() else {
        return true;
    };

    if first.is_empty() || first.len() > 3 || !first.chars().all(|c| c.is_ascii_digit()) {
        return false;
    }
    if second.is_empty() || second.len() != 3 || !second.chars().all(|c| c.is_ascii_digit()) {
        return false;
    }
    for seg in segs {
        if seg.is_empty() || seg.len() != 3 || !seg.chars().all(|c| c.is_ascii_digit()) {
            return false;
        }
    }

    true
}

fn try_parse_date(
    text: &str,
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    system: ExcelDateSystem,
) -> Option<ExcelResult<i32>> {
    let raw = text.trim();
    if raw.is_empty() {
        return None;
    }

    let now_year = now_utc.year();
    // Token-based patterns like `2020-01-01`, `1/2/2020`, `2-Jan-2020`.
    for token in raw.split_whitespace() {
        if let Some(result) = parse_date_token(token, cfg, now_year) {
            return Some(result.and_then(|(y, m, d)| date_parts_to_serial(y, m, d, system)));
        }
    }

    // Month name forms that span tokens like `Jan 2 2020` or `2 Jan 2020`.
    let mut it = raw.split_whitespace();
    let mut t0 = it.next()?;
    let mut t1 = it.next();
    let mut t2 = it.next();
    loop {
        if let Some(result) = parse_month_name_sequence_opt(t0, t1, t2, now_year) {
            return Some(result.and_then(|(y, m, d)| date_parts_to_serial(y, m, d, system)));
        }
        let Some(next0) = t1 else {
            break;
        };
        t0 = next0;
        t1 = t2;
        t2 = it.next();
    }

    None
}

fn date_parts_to_serial(
    year: i32,
    month: u8,
    day: u8,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
    ymd_to_serial(ExcelDate::new(year, month, day), system).map_err(|_| ExcelError::Value)
}

fn parse_date_token(
    token: &str,
    cfg: ValueLocaleConfig,
    now_year: i32,
) -> Option<ExcelResult<(i32, u8, u8)>> {
    let token = token.trim_matches(',');
    if token.is_empty() {
        return None;
    }

    for sep in ['-', '/', '.'] {
        if !token.contains(sep) {
            continue;
        }

        let mut it = token.split(sep);
        let Some(a) = it.next() else {
            continue;
        };
        let Some(b) = it.next() else {
            continue;
        };
        let c = it.next();
        if it.next().is_some() {
            continue;
        }

        if let Some(c) = c {
            let parts = [a, b, c];
            let numeric_parts = parts.iter().filter(|p| is_ascii_digit_str(p)).count();
            let month_parts = parts
                .iter()
                .filter(|p| parse_month_name(p).is_some())
                .count();
            if numeric_parts == 0 && month_parts == 0 {
                return None;
            }

            if parts.iter().all(|p| is_ascii_digit_str(p)) {
                return Some(parse_numeric_3part_date(&parts, cfg));
            }
            if let Some(result) = parse_month_name_parts(&parts, now_year) {
                return Some(result);
            }

            // Date-ish pattern, but invalid.
            if numeric_parts >= 2 || month_parts >= 1 {
                return Some(Err(ExcelError::Value));
            }

            return None;
        }

            // Excel accepts `m/d` or `d/m` without a year, using the current year.
            // Restrict this to the locale's date separator to avoid confusing decimal numbers
            // like `1.5` in locales where `.` is the decimal separator.
            let allow_missing_year = sep == '/' || sep == cfg.separators.date_sep;
            if !allow_missing_year {
                continue;
            }

        let parts = [a, b];
        if parts.iter().all(|p| is_ascii_digit_str(p)) {
            return Some(parse_numeric_2part_date(&parts, cfg, now_year));
        }

        if let Some(result) = parse_month_name_parts(&parts, now_year) {
            return Some(result);
        }

        let numeric_parts = parts.iter().filter(|p| is_ascii_digit_str(p)).count();
        let month_parts = parts
            .iter()
            .filter(|p| parse_month_name(p).is_some())
            .count();
        if numeric_parts == 0 && month_parts == 0 {
            return None;
        }

        if numeric_parts >= 1 || month_parts >= 1 {
            return Some(Err(ExcelError::Value));
        }

        return None;
    }

    None
}

fn parse_numeric_3part_date(parts: &[&str], cfg: ValueLocaleConfig) -> ExcelResult<(i32, u8, u8)> {
    debug_assert_eq!(parts.len(), 3);
    let a = parts[0].trim();
    let b = parts[1].trim();
    let c = parts[2].trim();

    if a.len() == 4 {
        // ISO-ish: yyyy-mm-dd
        let year = parse_year_component(a)?;
        let month = parse_u8_component(b)?;
        let day = parse_u8_component(c)?;
        return Ok((year, month, day));
    }

    match cfg.date_order {
        DateOrder::MDY => {
            let month = parse_u8_component(a)?;
            let day = parse_u8_component(b)?;
            let year = parse_year_component(c)?;
            Ok((year, month, day))
        }
        DateOrder::DMY => {
            let day = parse_u8_component(a)?;
            let month = parse_u8_component(b)?;
            let year = parse_year_component(c)?;
            Ok((year, month, day))
        }
        DateOrder::YMD => {
            let year = parse_year_component(a)?;
            let month = parse_u8_component(b)?;
            let day = parse_u8_component(c)?;
            Ok((year, month, day))
        }
    }
}

fn parse_numeric_2part_date(
    parts: &[&str],
    cfg: ValueLocaleConfig,
    now_year: i32,
) -> ExcelResult<(i32, u8, u8)> {
    debug_assert_eq!(parts.len(), 2);
    let a = parts[0].trim();
    let b = parts[1].trim();

    match cfg.date_order {
        DateOrder::MDY | DateOrder::YMD => {
            let month = parse_u8_component(a)?;
            let day = parse_u8_component(b)?;
            Ok((now_year, month, day))
        }
        DateOrder::DMY => {
            let day = parse_u8_component(a)?;
            let month = parse_u8_component(b)?;
            Ok((now_year, month, day))
        }
    }
}

fn parse_month_name_parts(parts: &[&str], now_year: i32) -> Option<ExcelResult<(i32, u8, u8)>> {
    match parts.len() {
        3 => {
            let a = parts[0].trim();
            let b = parts[1].trim();
            let c = parts[2].trim();

            if let Some(month) = parse_month_name(a) {
                let day = match parse_u8_component(b) {
                    Ok(v) => v,
                    Err(e) => return Some(Err(e)),
                };
                let year = match parse_year_component(c) {
                    Ok(v) => v,
                    Err(e) => return Some(Err(e)),
                };
                return Some(Ok((year, month, day)));
            }

            if let Some(month) = parse_month_name(b) {
                if !is_ascii_digit_str(a) || !is_ascii_digit_str(c) {
                    return Some(Err(ExcelError::Value));
                }

                let (year_str, day_str) = if a.len() == 4 { (a, c) } else { (c, a) };
                let year = match parse_year_component(year_str) {
                    Ok(v) => v,
                    Err(e) => return Some(Err(e)),
                };
                let day = match parse_u8_component(day_str) {
                    Ok(v) => v,
                    Err(e) => return Some(Err(e)),
                };
                return Some(Ok((year, month, day)));
            }

            None
        }
        2 => {
            let a = parts[0].trim();
            let b = parts[1].trim();

            if let Some(month) = parse_month_name(a) {
                let day = match parse_u8_component(b) {
                    Ok(v) => v,
                    Err(e) => return Some(Err(e)),
                };
                return Some(Ok((now_year, month, day)));
            }

            if let Some(month) = parse_month_name(b) {
                let day = match parse_u8_component(a) {
                    Ok(v) => v,
                    Err(e) => return Some(Err(e)),
                };
                return Some(Ok((now_year, month, day)));
            }

            None
        }
        _ => None,
    }
}

fn parse_month_name_sequence_opt(
    token0: &str,
    token1: Option<&str>,
    token2: Option<&str>,
    now_year: i32,
) -> Option<ExcelResult<(i32, u8, u8)>> {
    let month_token_clean = token0.trim_matches(|c: char| matches!(c, ',' | '.'));
    if let Some(month) = parse_month_name(month_token_clean) {
        let day_token = token1?;
        let day = match parse_u8_component(day_token.trim_matches(',')) {
            Ok(v) => v,
            Err(e) => return Some(Err(e)),
        };

        if let Some(year_token) = token2 {
            if year_token.chars().all(|c| c.is_ascii_digit() || c == ',') {
                let year = match parse_year_component(year_token.trim_matches(',')) {
                    Ok(v) => v,
                    Err(e) => return Some(Err(e)),
                };
                return Some(Ok((year, month, day)));
            }
        }

        return Some(Ok((now_year, month, day)));
    }

    if !is_ascii_digit_str(token0.trim_matches(',')) {
        return None;
    }
    let day = match parse_u8_component(token0.trim_matches(',')) {
        Ok(v) => v,
        Err(e) => return Some(Err(e)),
    };

    let month_token = token1?;
    let month_token_clean = month_token.trim_matches(|c: char| matches!(c, ',' | '.'));
    let month = parse_month_name(month_token_clean)?;

    if let Some(year_token) = token2 {
        if year_token.chars().all(|c| c.is_ascii_digit() || c == ',') {
            let year = match parse_year_component(year_token.trim_matches(',')) {
                Ok(v) => v,
                Err(e) => return Some(Err(e)),
            };
            return Some(Ok((year, month, day)));
        }
    }

    Some(Ok((now_year, month, day)))
}

fn parse_month_name(token: &str) -> Option<u8> {
    let token = token.trim_matches(|c: char| !c.is_ascii_alphabetic());
    if token.is_empty() {
        return None;
    }
    if token.eq_ignore_ascii_case("jan") || token.eq_ignore_ascii_case("january") {
        return Some(1);
    }
    if token.eq_ignore_ascii_case("feb") || token.eq_ignore_ascii_case("february") {
        return Some(2);
    }
    if token.eq_ignore_ascii_case("mar") || token.eq_ignore_ascii_case("march") {
        return Some(3);
    }
    if token.eq_ignore_ascii_case("apr") || token.eq_ignore_ascii_case("april") {
        return Some(4);
    }
    if token.eq_ignore_ascii_case("may") {
        return Some(5);
    }
    if token.eq_ignore_ascii_case("jun") || token.eq_ignore_ascii_case("june") {
        return Some(6);
    }
    if token.eq_ignore_ascii_case("jul") || token.eq_ignore_ascii_case("july") {
        return Some(7);
    }
    if token.eq_ignore_ascii_case("aug") || token.eq_ignore_ascii_case("august") {
        return Some(8);
    }
    if token.eq_ignore_ascii_case("sep")
        || token.eq_ignore_ascii_case("sept")
        || token.eq_ignore_ascii_case("september")
    {
        return Some(9);
    }
    if token.eq_ignore_ascii_case("oct") || token.eq_ignore_ascii_case("october") {
        return Some(10);
    }
    if token.eq_ignore_ascii_case("nov") || token.eq_ignore_ascii_case("november") {
        return Some(11);
    }
    if token.eq_ignore_ascii_case("dec") || token.eq_ignore_ascii_case("december") {
        return Some(12);
    }
    None
}

fn parse_year_component(s: &str) -> ExcelResult<i32> {
    let raw: i32 = s.parse().map_err(|_| ExcelError::Value)?;
    if raw < 0 {
        return Err(ExcelError::Value);
    }
    if (0..100).contains(&raw) {
        Ok(if raw <= 29 { 2000 + raw } else { 1900 + raw })
    } else {
        Ok(raw)
    }
}

fn parse_u8_component(s: &str) -> ExcelResult<u8> {
    let raw: u8 = s.parse().map_err(|_| ExcelError::Value)?;
    if raw == 0 {
        return Err(ExcelError::Value);
    }
    Ok(raw)
}

fn is_ascii_digit_str(s: &str) -> bool {
    let s = s.trim();
    !s.is_empty() && s.chars().all(|c| c.is_ascii_digit())
}

fn try_parse_time(text: &str) -> Option<ExcelResult<f64>> {
    let raw = text.trim();
    if raw.is_empty() {
        return None;
    }

    let mut tokens = raw.split_whitespace().peekable();
    while let Some(token) = tokens.next() {
        let token = token.trim_matches(',');
        if token.is_empty() {
            continue;
        }

        if token.contains(':') {
            if let Some(&next) = tokens.peek() {
                let suffix = next.trim_matches(',');
                if suffix.eq_ignore_ascii_case("AM") || suffix.eq_ignore_ascii_case("PM") {
                    let mut combined = String::with_capacity(token.len() + 1 + suffix.len());
                    combined.push_str(token);
                    combined.push(' ');
                    combined.push_str(suffix);
                    return Some(parse_time_candidate(&combined));
                }
            }
            return Some(parse_time_candidate(token));
        }

        if looks_like_ampm_suffix(token) {
            return Some(parse_time_candidate(token));
        }

        if is_ascii_digit_str(token) {
            if let Some(&next) = tokens.peek() {
                let suffix = next.trim_matches(',');
                if suffix.eq_ignore_ascii_case("AM") || suffix.eq_ignore_ascii_case("PM") {
                    let mut combined = String::with_capacity(token.len() + 1 + suffix.len());
                    combined.push_str(token);
                    combined.push(' ');
                    combined.push_str(suffix);
                    return Some(parse_time_candidate(&combined));
                }
            }
        }
    }

    None
}

fn looks_like_ampm_suffix(token: &str) -> bool {
    let bytes = token.as_bytes();
    if bytes.len() < 2 {
        return false;
    }
    let suffix = &bytes[bytes.len() - 2..];
    (suffix.eq_ignore_ascii_case(b"AM") || suffix.eq_ignore_ascii_case(b"PM"))
        && bytes.iter().copied().any(|b| b.is_ascii_digit())
}

fn parse_time_candidate(candidate: &str) -> ExcelResult<f64> {
    let mut s = candidate.trim();
    if s.is_empty() {
        return Err(ExcelError::Value);
    }

    let mut ampm: Option<&str> = None;
    if s.len() >= 2 {
        let bytes = s.as_bytes();
        let suffix = &bytes[bytes.len() - 2..];
        if suffix.eq_ignore_ascii_case(b"AM") {
            ampm = Some("AM");
            s = s[..s.len() - 2].trim_end();
        } else if suffix.eq_ignore_ascii_case(b"PM") {
            ampm = Some("PM");
            s = s[..s.len() - 2].trim_end();
        }
    }

    if s.is_empty() {
        return Err(ExcelError::Value);
    }

    let (mut hour, minute, second) = if s.contains(':') {
        parse_colon_time(s)?
    } else {
        if ampm.is_none() {
            return Err(ExcelError::Value);
        }
        let hour: i32 = s.parse().map_err(|_| ExcelError::Value)?;
        (hour, 0, 0.0)
    };

    if minute < 0 || minute >= 60 || second < 0.0 || second >= 60.0 {
        return Err(ExcelError::Value);
    }

    if let Some(ampm) = ampm {
        if hour < 0 || hour > 12 {
            return Err(ExcelError::Value);
        }
        if hour == 12 {
            hour = 0;
        }
        if ampm == "PM" {
            hour += 12;
        }
    }

    if hour < 0 {
        return Err(ExcelError::Value);
    }

    let total_seconds = (hour as f64) * 3600.0 + (minute as f64) * 60.0 + second;
    Ok(total_seconds / 86_400.0)
}

fn parse_colon_time(s: &str) -> ExcelResult<(i32, i32, f64)> {
    let mut it = s.split(':');
    let Some(hour_s) = it.next() else {
        return Err(ExcelError::Value);
    };
    let Some(minute_s) = it.next() else {
        return Err(ExcelError::Value);
    };
    let second_s = it.next();
    if it.next().is_some() {
        return Err(ExcelError::Value);
    }

    let hour: i32 = hour_s.trim().parse().map_err(|_| ExcelError::Value)?;
    let minute: i32 = minute_s.trim().parse().map_err(|_| ExcelError::Value)?;
    let second = match second_s {
        Some(s) => parse_seconds_component(s.trim())?,
        None => 0.0,
    };

    Ok((hour, minute, second))
}

fn parse_seconds_component(s: &str) -> ExcelResult<f64> {
    if s.is_empty() {
        return Err(ExcelError::Value);
    }
    let seconds: f64 = s.parse().map_err(|_| ExcelError::Value)?;
    if !seconds.is_finite() {
        return Err(ExcelError::Value);
    }
    let whole = seconds.trunc();
    if whole < 0.0 || whole >= 60.0 {
        return Err(ExcelError::Value);
    }
    Ok(seconds)
}

#[cfg(test)]
mod tests {
    use chrono::TimeZone;

    use super::*;
    use crate::date::ExcelDateSystem;

    #[test]
    fn parses_missing_year_using_now_year() {
        let now = Utc.with_ymd_and_hms(2024, 6, 1, 0, 0, 0).unwrap();
        let cfg = ValueLocaleConfig::en_us();
        let serial = parse_datevalue_text("1/2", cfg, now, ExcelDateSystem::EXCEL_1900).unwrap();
        let expected = date_parts_to_serial(2024, 1, 2, ExcelDateSystem::EXCEL_1900).unwrap();
        assert_eq!(serial, expected);
    }

    #[test]
    fn respects_dmy_order() {
        let now = Utc.with_ymd_and_hms(2024, 6, 1, 0, 0, 0).unwrap();
        let cfg = ValueLocaleConfig {
            date_order: DateOrder::DMY,
            ..ValueLocaleConfig::en_us()
        };
        let serial =
            parse_datevalue_text("1/2/2020", cfg, now, ExcelDateSystem::EXCEL_1900).unwrap();
        let expected = date_parts_to_serial(2020, 2, 1, ExcelDateSystem::EXCEL_1900).unwrap();
        assert_eq!(serial, expected);
    }

    #[test]
    fn parses_month_names() {
        let now = Utc.with_ymd_and_hms(2024, 6, 1, 0, 0, 0).unwrap();
        let cfg = ValueLocaleConfig::en_us();
        let serial =
            parse_datevalue_text("January 2, 2020", cfg, now, ExcelDateSystem::EXCEL_1900).unwrap();
        let expected = date_parts_to_serial(2020, 1, 2, ExcelDateSystem::EXCEL_1900).unwrap();
        assert_eq!(serial, expected);
        let serial =
            parse_datevalue_text("2-Jan-2020", cfg, now, ExcelDateSystem::EXCEL_1900).unwrap();
        assert_eq!(serial, expected);
    }

    #[test]
    fn parses_timevalue_hour_only_ampm() {
        let cfg = ValueLocaleConfig::en_us();
        assert_eq!(parse_timevalue_text("1 PM", cfg).unwrap(), 13.0 / 24.0);
        assert_eq!(parse_timevalue_text("1PM", cfg).unwrap(), 13.0 / 24.0);
    }

    #[test]
    fn parses_dot_separated_dates_without_year_for_dot_date_sep_locales() {
        let now = Utc.with_ymd_and_hms(2024, 6, 1, 0, 0, 0).unwrap();
        let cfg = ValueLocaleConfig::de_de();
        let serial = parse_datevalue_text("1.2", cfg, now, ExcelDateSystem::EXCEL_1900).unwrap();
        let expected = date_parts_to_serial(2024, 2, 1, ExcelDateSystem::EXCEL_1900).unwrap();
        assert_eq!(serial, expected);
    }
}
