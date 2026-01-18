use crate::literal::RenderedText;
use crate::{FormatOptions, LiteralLayoutOp};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DateSystem {
    Excel1900,
    Excel1904,
}

#[derive(Debug, Clone, Copy)]
struct DateTimeParts {
    year: i32,
    month: u32,
    day: u32,
    hour: u32,
    minute: u32,
    second: u32,
    subsecond_units: u32,
    subsecond_digits: usize,
    // Used for elapsed time formats like `[h]:mm:ss`
    total_seconds: i64,
    // 0=Sunday..6=Saturday
    weekday: u32,
}

/// Best-effort check for whether a number format section contains Excel date/time tokens.
///
/// This is intentionally lightweight: it does *not* fully parse Excel's number format grammar.
/// It is used by higher-level tooling (e.g. pivot cache builders) to infer when a numeric value
/// should be treated as a serial date/time based on its format.
pub fn looks_like_datetime(section: &str) -> bool {
    let mut in_quotes = false;
    let mut escape = false;
    let mut chars = section.chars().peekable();

    while let Some(ch) = chars.next() {
        if escape {
            escape = false;
            continue;
        }
        if in_quotes {
            if ch == '"' {
                in_quotes = false;
            }
            continue;
        }

        match ch {
            '"' => in_quotes = true,
            '\\' => escape = true,
            '[' => {
                // Elapsed time: [h], [m], [s]
                let mut content = String::new();
                while let Some(c) = chars.next() {
                    if c == ']' {
                        break;
                    }
                    content.push(c);
                }
                if is_elapsed_time_token(&content) {
                    return true;
                }
            }
            'y' | 'Y' | 'd' | 'D' | 'h' | 'H' | 's' | 'S' => return true,
            'm' | 'M' => return true,
            'a' | 'A' => {
                // AM/PM or A/P markers (case-insensitive).
                let mut probe = String::new();
                probe.push(ch);
                let mut clone = chars.clone();
                for _ in 0..4 {
                    if let Some(c) = clone.next() {
                        probe.push(c);
                    } else {
                        break;
                    }
                }
                if probe
                    .get(.."am/pm".len())
                    .is_some_and(|p| p.eq_ignore_ascii_case("am/pm"))
                    || probe
                        .get(.."a/p".len())
                        .is_some_and(|p| p.eq_ignore_ascii_case("a/p"))
                {
                    return true;
                }
            }
            _ => {}
        }
    }

    false
}

fn is_elapsed_time_token(lower: &str) -> bool {
    if lower.is_empty() {
        return false;
    }
    lower.chars().all(|c| matches!(c, 'h' | 'H'))
        || lower.chars().all(|c| matches!(c, 'm' | 'M'))
        || lower.chars().all(|c| matches!(c, 's' | 'S'))
}

fn parse_elapsed_bracket_token(lower: &str) -> Option<Token> {
    if lower.is_empty() {
        return None;
    }
    let count = lower.chars().count();
    if lower.chars().all(|c| matches!(c, 'h' | 'H')) {
        return Some(Token::ElapsedHours(count));
    }
    if lower.chars().all(|c| matches!(c, 'm' | 'M')) {
        return Some(Token::ElapsedMinutes(count));
    }
    if lower.chars().all(|c| matches!(c, 's' | 'S')) {
        return Some(Token::ElapsedSeconds(count));
    }
    None
}

pub(crate) fn format_datetime(serial: f64, pattern: &str, options: &FormatOptions) -> RenderedText {
    let mut tokens = tokenize(pattern);
    let has_ampm = tokens
        .iter()
        .any(|t| matches!(t, Token::AmPmLong | Token::AmPmShort));
    disambiguate_minutes(&mut tokens);

    let frac_digits = tokens
        .iter()
        .filter_map(|t| match t {
            Token::FractionalSeconds(d) => Some(*d),
            _ => None,
        })
        .max()
        .unwrap_or(0);

    let Some(parts) = serial_to_parts(serial, options.date_system, frac_digits) else {
        return RenderedText::new("#####".to_string());
    };
    render_tokens(&tokens, &parts, has_ampm, options)
}

fn serial_to_parts(serial: f64, date_system: DateSystem, frac_second_digits: usize) -> Option<DateTimeParts> {
    if !serial.is_finite() {
        return None;
    }
    if serial < 0.0 {
        // Excel shows ##### for negative date serials when formatted as dates.
        return None;
    }

    let mut days = serial.floor() as i64;
    let frac = serial - days as f64;
    let frac_second_digits = frac_second_digits.min(9);
    let scale = if frac_second_digits == 0 {
        1_i64
    } else {
        10_i64.pow(frac_second_digits as u32)
    };

    let mut total_units = (frac * 86_400.0 * scale as f64).round() as i64;
    let day_units = 86_400_i64 * scale;
    if total_units >= day_units {
        total_units = 0;
        days += 1;
    }

    let seconds_total = total_units / scale;
    let subsecond_units = (total_units % scale) as u32;

    let hour = (seconds_total / 3_600) as u32;
    let minute = ((seconds_total % 3_600) / 60) as u32;
    let second = (seconds_total % 60) as u32;

    // Convert the date part.
    let (year, month, day, weekday) = match date_system {
        DateSystem::Excel1900 => excel_1900_days_to_ymd(days)?,
        DateSystem::Excel1904 => excel_1904_days_to_ymd(days)?,
    };

    Some(DateTimeParts {
        year,
        month,
        day,
        hour,
        minute,
        second,
        subsecond_units,
        subsecond_digits: frac_second_digits,
        total_seconds: (days * 86_400) + seconds_total,
        weekday,
    })
}

fn excel_1900_days_to_ymd(days: i64) -> Option<(i32, u32, u32, u32)> {
    // In Excel's 1900 date system, day 0 is 1899-12-31.
    // Day 60 is the fictitious 1900-02-29 (Lotus bug).
    let base = days_from_civil(1899, 12, 31);

    if days == 60 {
        // Return the fake date; weekday is computed from the adjusted day count
        // (base + 60) which corresponds to 1900-03-01 in the real calendar.
        let weekday = weekday_from_days(base + 60);
        return Some((1900, 2, 29, weekday));
    }

    let adjusted = if days < 60 { days } else { days - 1 };
    let abs_days = base + adjusted;
    let (year, month, day) = civil_from_days(abs_days);
    let weekday = weekday_from_days(abs_days);
    Some((year, month, day, weekday))
}

fn excel_1904_days_to_ymd(days: i64) -> Option<(i32, u32, u32, u32)> {
    // In Excel's 1904 date system, day 0 is 1904-01-01.
    let base = days_from_civil(1904, 1, 1);
    let abs_days = base + days;
    let (year, month, day) = civil_from_days(abs_days);
    let weekday = weekday_from_days(abs_days);
    Some((year, month, day, weekday))
}

// date algorithms from Howard Hinnant (public domain).
fn days_from_civil(year: i32, month: u32, day: u32) -> i64 {
    let mut y = year as i64;
    let m = month as i64;
    let d = day as i64;

    y -= if m <= 2 { 1 } else { 0 };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let mp = m + if m > 2 { -3 } else { 9 };
    let doy = (153 * mp + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe - 719_468
}

fn civil_from_days(days: i64) -> (i32, u32, u32) {
    let z = days + 719_468;
    let era = (if z >= 0 { z } else { z - 146_096 }) / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = mp + if mp < 10 { 3 } else { -9 };
    let year = y + if m <= 2 { 1 } else { 0 };
    (year as i32, m as u32, d as u32)
}

fn weekday_from_days(days: i64) -> u32 {
    // 1970-01-01 was Thursday. Map 0=Sunday..6=Saturday.
    (days + 4).rem_euclid(7) as u32
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum Token {
    Literal(String),
    Year(usize),
    Day(usize),
    Hour(usize),
    Second(usize),
    FractionalSeconds(usize),
    MonthOrMinute(usize),
    Month(usize),
    Minute(usize),
    DateSep,
    TimeSep,
    AmPmLong,
    AmPmShort,
    ElapsedHours(usize),
    ElapsedMinutes(usize),
    ElapsedSeconds(usize),
    Underscore(char),
    Fill(char),
}

fn tokenize(pattern: &str) -> Vec<Token> {
    let mut tokens = Vec::new();
    let mut literal_buf = String::new();
    let mut in_quotes = false;
    let mut chars = pattern.chars().peekable();

    while let Some(ch) = chars.next() {
        if in_quotes {
            if ch == '"' {
                in_quotes = false;
            } else {
                literal_buf.push(ch);
            }
            continue;
        }

        match ch {
            '"' => in_quotes = true,
            '\\' => {
                if let Some(next) = chars.next() {
                    literal_buf.push(next);
                }
            }
            '_' => {
                if let Some(width_of) = chars.next() {
                    flush_literal(&mut literal_buf, &mut tokens);
                    tokens.push(Token::Underscore(width_of));
                } else {
                    literal_buf.push('_');
                }
            }
            '*' => {
                if let Some(fill_with) = chars.next() {
                    flush_literal(&mut literal_buf, &mut tokens);
                    tokens.push(Token::Fill(fill_with));
                } else {
                    literal_buf.push('*');
                }
            }
            '/' => {
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(Token::DateSep);
            }
            ':' => {
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(Token::TimeSep);
            }
            '[' => {
                let mut content = String::new();
                while let Some(c) = chars.next() {
                    if c == ']' {
                        break;
                    }
                    content.push(c);
                }
                flush_literal(&mut literal_buf, &mut tokens);
                if let Some(token) = parse_elapsed_bracket_token(&content) {
                    tokens.push(token);
                }
            }
            'a' | 'A' => {
                // AM/PM or A/P markers (case-insensitive).
                let mut probe = String::new();
                probe.push(ch);
                let mut clone = chars.clone();
                for _ in 0..4 {
                    if let Some(c) = clone.next() {
                        probe.push(c);
                    } else {
                        break;
                    }
                }

                if probe
                    .get(.."am/pm".len())
                    .is_some_and(|p| p.eq_ignore_ascii_case("am/pm"))
                {
                    // Consume `M/PM` (4 chars) from the original iterator.
                    for _ in 0..4 {
                        chars.next();
                    }
                    flush_literal(&mut literal_buf, &mut tokens);
                    tokens.push(Token::AmPmLong);
                } else if probe
                    .get(.."a/p".len())
                    .is_some_and(|p| p.eq_ignore_ascii_case("a/p"))
                {
                    // Consume `/P` (2 chars).
                    for _ in 0..2 {
                        chars.next();
                    }
                    flush_literal(&mut literal_buf, &mut tokens);
                    tokens.push(Token::AmPmShort);
                } else {
                    // Not a recognized token: treat as literal.
                    literal_buf.push(ch);
                }
            }
            'y' | 'Y' => {
                let count = consume_run(ch, &mut chars);
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(Token::Year(count));
            }
            'd' | 'D' => {
                let count = consume_run(ch, &mut chars);
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(Token::Day(count));
            }
            'h' | 'H' => {
                let count = consume_run(ch, &mut chars);
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(Token::Hour(count));
            }
            's' | 'S' => {
                let count = consume_run(ch, &mut chars);
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(Token::Second(count));

                // Fractional seconds: `ss.0`, `ss.00`, `ss.000`, etc.
                if chars.peek().copied() == Some('.') {
                    let mut clone = chars.clone();
                    let _ = clone.next(); // '.'
                    let mut zeros = 0usize;
                    while let Some('0') = clone.next() {
                        zeros += 1;
                    }
                    if zeros > 0 {
                        // Consume '.' + zeros from the original iterator.
                        let _ = chars.next();
                        for _ in 0..zeros {
                            let _ = chars.next();
                        }
                        tokens.push(Token::FractionalSeconds(zeros));
                    }
                }
            }
            'm' | 'M' => {
                let count = consume_run(ch, &mut chars);
                flush_literal(&mut literal_buf, &mut tokens);
                tokens.push(Token::MonthOrMinute(count));
            }
            _ => literal_buf.push(ch),
        }
    }

    flush_literal(&mut literal_buf, &mut tokens);
    tokens
}

fn consume_run(first: char, chars: &mut std::iter::Peekable<std::str::Chars<'_>>) -> usize {
    let mut count = 1;
    while let Some(next) = chars.peek().copied() {
        if next.eq_ignore_ascii_case(&first) {
            chars.next();
            count += 1;
        } else {
            break;
        }
    }
    count
}

fn flush_literal(buf: &mut String, tokens: &mut Vec<Token>) {
    if buf.is_empty() {
        return;
    }
    tokens.push(Token::Literal(std::mem::take(buf)));
}

fn disambiguate_minutes(tokens: &mut [Token]) {
    // Replace MonthOrMinute based on adjacent time tokens.
    for idx in 0..tokens.len() {
        let Token::MonthOrMinute(count) = tokens[idx] else {
            continue;
        };

        // In Excel formats, `mmm`/`mmmm`/`mmmmm` are always month name variants. Minutes only use
        // `m` or `mm` (disambiguated by neighboring hour/second tokens).
        if count >= 3 {
            tokens[idx] = Token::Month(count);
            continue;
        }

        let prev = prev_non_literal(tokens, idx);
        let next = next_non_literal(tokens, idx);
        let is_minute = matches!(prev, Some(Token::Hour(_)) | Some(Token::ElapsedHours(_)))
            || matches!(next, Some(Token::Second(_)) | Some(Token::ElapsedSeconds(_)));

        tokens[idx] = if is_minute {
            Token::Minute(count)
        } else {
            Token::Month(count)
        };
    }
}

fn prev_non_literal(tokens: &[Token], idx: usize) -> Option<&Token> {
    for j in (0..idx).rev() {
        match &tokens[j] {
            Token::Literal(_) | Token::DateSep | Token::TimeSep | Token::Underscore(_) | Token::Fill(_) => continue,
            t => return Some(t),
        }
    }
    None
}

fn next_non_literal(tokens: &[Token], idx: usize) -> Option<&Token> {
    for j in idx + 1..tokens.len() {
        match &tokens[j] {
            Token::Literal(_) | Token::DateSep | Token::TimeSep | Token::Underscore(_) | Token::Fill(_) => continue,
            t => return Some(t),
        }
    }
    None
}

fn render_tokens(tokens: &[Token], parts: &DateTimeParts, has_ampm: bool, options: &FormatOptions) -> RenderedText {
    let mut out = RenderedText::new(String::new());

    for (idx, token) in tokens.iter().enumerate() {
        match token {
            Token::Literal(s) => out.push_str(s),
            Token::DateSep => out.push(options.locale.date_sep),
            Token::TimeSep => out.push(options.locale.time_sep),
            Token::Year(count) => out.push_str(&format_year(parts.year, *count)),
            Token::Month(count) => out.push_str(&format_month(parts.month, *count)),
            Token::Minute(count) => out.push_str(&format_two(parts.minute, *count)),
            Token::Day(count) => out.push_str(&format_day(parts.day, parts.weekday, *count)),
            Token::Hour(count) => {
                let hour = if has_ampm {
                    let mut h = (parts.hour % 12) as u32;
                    if h == 0 {
                        h = 12;
                    }
                    h
                } else {
                    parts.hour
                };
                out.push_str(&format_two(hour, *count));
            }
            Token::Second(count) => out.push_str(&format_two(parts.second, *count)),
            Token::FractionalSeconds(digits) => {
                out.push(options.locale.decimal_sep);
                out.push_str(&format_fractional_seconds(parts, *digits));
            }
            Token::AmPmLong => out.push_str(if parts.hour < 12 { "AM" } else { "PM" }),
            Token::AmPmShort => out.push_str(if parts.hour < 12 { "A" } else { "P" }),
            Token::ElapsedHours(width) => out.push_str(&format_elapsed(parts.total_seconds / 3600, *width)),
            Token::ElapsedMinutes(width) => out.push_str(&format_elapsed(parts.total_seconds / 60, *width)),
            Token::ElapsedSeconds(width) => out.push_str(&format_elapsed(parts.total_seconds, *width)),
            Token::Underscore(width_of) => {
                out.push_layout_op(LiteralLayoutOp::Underscore {
                    byte_index: out.text.len(),
                    width_of: *width_of,
                });
                out.push(' ');
            }
            Token::Fill(fill_with) => {
                out.push_layout_op(LiteralLayoutOp::Fill {
                    byte_index: out.text.len(),
                    fill_with: *fill_with,
                });
            }
            Token::MonthOrMinute(count) => {
                // Best-effort fallback: callers should run `disambiguate_minutes` during parsing,
                // but avoid panicking if a token stream bypasses that step.
                let count = *count;
                if count >= 3 {
                    out.push_str(&format_month(parts.month, count));
                    continue;
                }

                let prev = prev_non_literal(tokens, idx);
                let next = next_non_literal(tokens, idx);
                let is_minute = matches!(prev, Some(Token::Hour(_)) | Some(Token::ElapsedHours(_)))
                    || matches!(next, Some(Token::Second(_)) | Some(Token::ElapsedSeconds(_)));
                if is_minute {
                    out.push_str(&format_two(parts.minute, count));
                } else {
                    out.push_str(&format_month(parts.month, count));
                }
            }
        }
    }

    out
}

fn format_year(year: i32, count: usize) -> String {
    if count == 2 {
        format!("{:02}", (year % 100).abs())
    } else {
        // Excel treats y, yyy, yyyy similarly in most contexts.
        format!("{:04}", year)
    }
}

fn format_two(value: u32, count: usize) -> String {
    if count <= 1 {
        value.to_string()
    } else {
        format!("{:02}", value)
    }
}

fn format_elapsed(value: i64, width: usize) -> String {
    if width <= 1 {
        value.to_string()
    } else {
        format!("{value:0width$}", width = width)
    }
}

fn format_month(month: u32, count: usize) -> String {
    const MONTHS_SHORT: [&str; 12] = [
        "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    ];
    const MONTHS_LONG: [&str; 12] = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];

    match count {
        1 => month.to_string(),
        2 => format!("{:02}", month),
        3 => MONTHS_SHORT[(month.saturating_sub(1)) as usize].to_string(),
        4 => MONTHS_LONG[(month.saturating_sub(1)) as usize].to_string(),
        _ => MONTHS_LONG[(month.saturating_sub(1)) as usize]
            .chars()
            .next()
            .unwrap_or('?')
            .to_string(),
    }
}

fn format_day(day: u32, weekday: u32, count: usize) -> String {
    const DAYS_SHORT: [&str; 7] = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    const DAYS_LONG: [&str; 7] = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ];
    match count {
        1 => day.to_string(),
        2 => format!("{:02}", day),
        3 => DAYS_SHORT[weekday as usize].to_string(),
        _ => DAYS_LONG[weekday as usize].to_string(),
    }
}

fn format_fractional_seconds(parts: &DateTimeParts, digits: usize) -> String {
    if digits == 0 {
        return String::new();
    }

    let total_digits = parts.subsecond_digits;
    if total_digits == 0 {
        return "0".repeat(digits);
    }

    let digits = digits.min(total_digits);
    let divisor = 10_u32.pow((total_digits - digits) as u32);
    let value = parts.subsecond_units / divisor;
    format!("{value:0digits$}", digits = digits)
}
