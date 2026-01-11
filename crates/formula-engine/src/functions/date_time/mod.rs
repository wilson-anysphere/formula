use std::collections::HashSet;

use crate::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use crate::error::{ExcelError, ExcelResult};

const DEFAULT_WEEKEND_MASK: u8 = (1 << 5) | (1 << 6); // Saturday + Sunday (Mon=0..Sun=6)

/// TIME(hour, minute, second)
pub fn time(hour: i32, minute: i32, second: i32) -> ExcelResult<f64> {
    if hour < 0 || minute < 0 || second < 0 {
        return Err(ExcelError::Num);
    }
    let total_seconds = i64::from(hour) * 3600 + i64::from(minute) * 60 + i64::from(second);
    Ok(total_seconds as f64 / 86400.0)
}

/// TIMEVALUE(time_text)
pub fn timevalue(time_text: &str) -> ExcelResult<f64> {
    let raw = time_text.trim();
    if raw.is_empty() {
        return Err(ExcelError::Value);
    }

    // Find the portion that looks like a time (contains ':') and preserve an
    // optional "AM"/"PM" suffix even when separated by whitespace.
    let parts: Vec<&str> = raw.split_whitespace().collect();
    if let Some((idx, token)) = parts.iter().enumerate().find(|(_, part)| part.contains(':')) {
        if idx + 1 < parts.len() {
            let suffix = parts[idx + 1];
            if suffix.eq_ignore_ascii_case("AM") || suffix.eq_ignore_ascii_case("PM") {
                let combined = format!("{token} {suffix}");
                return parse_time_token(&combined);
            }
        }
        return parse_time_token(token);
    }

    parse_time_token(raw)
}

fn parse_time_token(token: &str) -> ExcelResult<f64> {
    let mut s = token.trim().to_string();
    let mut ampm: Option<&str> = None;
    if s.to_ascii_uppercase().ends_with("AM") {
        ampm = Some("AM");
        s = s[..s.len() - 2].trim().to_string();
    } else if s.to_ascii_uppercase().ends_with("PM") {
        ampm = Some("PM");
        s = s[..s.len() - 2].trim().to_string();
    }

    let parts: Vec<&str> = s.split(':').collect();
    if parts.len() < 2 || parts.len() > 3 {
        return Err(ExcelError::Value);
    }

    let mut hour: i32 = parts[0].trim().parse().map_err(|_| ExcelError::Value)?;
    let minute: i32 = parts[1].trim().parse().map_err(|_| ExcelError::Value)?;
    let second: i32 = if parts.len() == 3 {
        parts[2].trim().parse().map_err(|_| ExcelError::Value)?
    } else {
        0
    };

    if minute < 0 || minute >= 60 || second < 0 || second >= 60 {
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

    time(hour, minute, second)
}

/// DATEVALUE(date_text)
pub fn datevalue(date_text: &str, system: ExcelDateSystem) -> ExcelResult<i32> {
    let raw = date_text.trim();
    if raw.is_empty() {
        return Err(ExcelError::Value);
    }

    let token = raw
        .split_whitespace()
        .find(|part| part.contains('-') || part.contains('/') || part.contains('.'))
        .unwrap_or(raw);

    let (year, month, day) = parse_date_token(token)?;
    ymd_to_serial(ExcelDate::new(year, month, day), system)
}

fn parse_date_token(token: &str) -> ExcelResult<(i32, u8, u8)> {
    let separators = ['-', '/', '.'];
    let sep = separators
        .iter()
        .copied()
        .find(|c| token.contains(*c))
        .ok_or(ExcelError::Value)?;

    let parts: Vec<&str> = token.split(sep).collect();
    if parts.len() != 3 {
        return Err(ExcelError::Value);
    }

    let a = parts[0].trim();
    let b = parts[1].trim();
    let c = parts[2].trim();

    if a.len() == 4 {
        // ISO-ish: yyyy-mm-dd
        let year: i32 = a.parse().map_err(|_| ExcelError::Value)?;
        let month: u8 = b.parse().map_err(|_| ExcelError::Value)?;
        let day: u8 = c.parse().map_err(|_| ExcelError::Value)?;
        return Ok((year, month, day));
    }

    // Default to US-style: mm/dd/yy(yy)
    let month: u8 = a.parse().map_err(|_| ExcelError::Value)?;
    let day: u8 = b.parse().map_err(|_| ExcelError::Value)?;
    let mut year: i32 = c.parse().map_err(|_| ExcelError::Value)?;
    if (0..100).contains(&year) {
        year = if year <= 29 { 2000 + year } else { 1900 + year };
    }
    Ok((year, month, day))
}

/// EOMONTH(start_date, months)
pub fn eomonth(start_date: i32, months: i32, system: ExcelDateSystem) -> ExcelResult<i32> {
    let start = crate::date::serial_to_ymd(start_date, system)?;
    let (year, month) = add_months(start.year, start.month, months);

    let (next_year, next_month) = add_months(year, month, 1);
    let first_next = ymd_to_serial(ExcelDate::new(next_year, next_month, 1), system)?;
    Ok(first_next - 1)
}

fn add_months(year: i32, month: u8, offset: i32) -> (i32, u8) {
    let base = year * 12 + i32::from(month - 1);
    let total = base + offset;
    let new_year = total.div_euclid(12);
    let new_month = (total.rem_euclid(12) + 1) as u8;
    (new_year, new_month)
}

/// EDATE(start_date, months)
pub fn edate(start_date: i32, months: i32, system: ExcelDateSystem) -> ExcelResult<i32> {
    let start = crate::date::serial_to_ymd(start_date, system)?;
    let (year, month) = add_months(start.year, start.month, months);

    // Preserve day-of-month where possible, otherwise clamp to the end of month.
    let mut day = start.day;
    while day > 0 {
        match ymd_to_serial(ExcelDate::new(year, month, day), system) {
            Ok(serial) => return Ok(serial),
            Err(ExcelError::Num) => day = day.saturating_sub(1),
            Err(e) => return Err(e),
        }
    }
    Err(ExcelError::Num)
}

/// WEEKDAY(serial_number, [return_type])
pub fn weekday(serial_number: i32, return_type: Option<i32>, system: ExcelDateSystem) -> ExcelResult<i32> {
    // Validate serial number is representable as a date in this system.
    let _ = crate::date::serial_to_ymd(serial_number, system)?;

    let monday0 = weekday_monday0(serial_number, system);
    let return_type = return_type.unwrap_or(1);
    match return_type {
        1 => Ok(((monday0 + 1).rem_euclid(7)) + 1),
        2 => Ok(monday0 + 1),
        3 => Ok(monday0),
        11..=17 => {
            let start = return_type - 11;
            Ok(((monday0 - start).rem_euclid(7)) + 1)
        }
        _ => Err(ExcelError::Num),
    }
}

fn weekday_monday0(serial_number: i32, system: ExcelDateSystem) -> i32 {
    match system {
        ExcelDateSystem::Excel1900 { lotus_compat } => {
            let mut adjusted = serial_number;
            if lotus_compat && serial_number > 60 {
                adjusted -= 1;
            }
            (adjusted - 1).rem_euclid(7)
        }
        ExcelDateSystem::Excel1904 => (serial_number + 4).rem_euclid(7),
    }
}

fn is_weekend_mask(serial_number: i32, system: ExcelDateSystem, weekend_mask: u8) -> bool {
    let monday0 = weekday_monday0(serial_number, system) as u8;
    weekend_mask & (1 << monday0) != 0
}

fn make_holiday_set(holidays: Option<&[i32]>) -> HashSet<i32> {
    holidays
        .unwrap_or(&[])
        .iter()
        .copied()
        .collect::<HashSet<_>>()
}

fn is_workday_with_weekend(
    serial_number: i32,
    system: ExcelDateSystem,
    weekend_mask: u8,
    holidays: &HashSet<i32>,
) -> bool {
    !is_weekend_mask(serial_number, system, weekend_mask) && !holidays.contains(&serial_number)
}

/// WORKDAY(start_date, days, [holidays])
pub fn workday(start_date: i32, days: i32, holidays: Option<&[i32]>, system: ExcelDateSystem) -> ExcelResult<i32> {
    workday_intl(start_date, days, DEFAULT_WEEKEND_MASK, holidays, system)
}

/// NETWORKDAYS(start_date, end_date, [holidays])
pub fn networkdays(start_date: i32, end_date: i32, holidays: Option<&[i32]>, system: ExcelDateSystem) -> ExcelResult<i32> {
    networkdays_intl(start_date, end_date, DEFAULT_WEEKEND_MASK, holidays, system)
}

/// WORKDAY.INTL(start_date, days, weekend_mask, [holidays])
pub fn workday_intl(
    start_date: i32,
    days: i32,
    weekend_mask: u8,
    holidays: Option<&[i32]>,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
    if weekend_mask == 0b111_1111 {
        return Err(ExcelError::Num);
    }

    let _ = crate::date::serial_to_ymd(start_date, system)?;
    let holiday_set = make_holiday_set(holidays);

    let direction = if days >= 0 { 1 } else { -1 };
    let mut remaining = days.abs();
    let mut current = start_date;

    if remaining == 0 {
        while !is_workday_with_weekend(current, system, weekend_mask, &holiday_set) {
            current += direction;
        }
        return Ok(current);
    }

    while remaining > 0 {
        current += direction;
        if is_workday_with_weekend(current, system, weekend_mask, &holiday_set) {
            remaining -= 1;
        }
    }

    Ok(current)
}

/// NETWORKDAYS.INTL(start_date, end_date, weekend_mask, [holidays])
pub fn networkdays_intl(
    start_date: i32,
    end_date: i32,
    weekend_mask: u8,
    holidays: Option<&[i32]>,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
    if weekend_mask == 0b111_1111 {
        return Err(ExcelError::Num);
    }

    let _ = crate::date::serial_to_ymd(start_date, system)?;
    let _ = crate::date::serial_to_ymd(end_date, system)?;
    let holiday_set = make_holiday_set(holidays);

    if start_date <= end_date {
        let mut count = 0i32;
        for serial in start_date..=end_date {
            if is_workday_with_weekend(serial, system, weekend_mask, &holiday_set) {
                count += 1;
            }
        }
        Ok(count)
    } else {
        let mut count = 0i32;
        for serial in end_date..=start_date {
            if is_workday_with_weekend(serial, system, weekend_mask, &holiday_set) {
                count += 1;
            }
        }
        Ok(-count)
    }
}

/// WEEKNUM(serial_number, [return_type])
pub fn weeknum(serial_number: i32, return_type: Option<i32>, system: ExcelDateSystem) -> ExcelResult<i32> {
    let date = crate::date::serial_to_ymd(serial_number, system)?;
    let return_type = return_type.unwrap_or(1);

    if return_type == 21 {
        return weeknum_iso(serial_number, system);
    }

    let week_start_monday0: i32 = match return_type {
        1 => 6,      // Sunday
        2 | 11 => 0, // Monday
        12 => 1,     // Tuesday
        13 => 2,     // Wednesday
        14 => 3,     // Thursday
        15 => 4,     // Friday
        16 => 5,     // Saturday
        17 => 6,     // Sunday
        _ => return Err(ExcelError::Num),
    };

    let year_start = ymd_to_serial(ExcelDate::new(date.year, 1, 1), system)?;
    let day_offset = serial_number - year_start;
    let start_weekday_monday0 = weekday_monday0(year_start, system);
    let start_index = (start_weekday_monday0 - week_start_monday0).rem_euclid(7);
    Ok(((day_offset + start_index) / 7) + 1)
}

fn weeknum_iso(serial_number: i32, system: ExcelDateSystem) -> ExcelResult<i32> {
    let monday0 = weekday_monday0(serial_number, system);
    let thursday_serial = i64::from(serial_number) + i64::from(3 - monday0);
    let thursday_serial = i32::try_from(thursday_serial).map_err(|_| ExcelError::Num)?;
    let thursday_date = crate::date::serial_to_ymd(thursday_serial, system)?;
    let iso_year = thursday_date.year;

    let jan4_serial = ymd_to_serial(ExcelDate::new(iso_year, 1, 4), system)?;
    let jan4_weekday = weekday_monday0(jan4_serial, system);
    let week1_start = i64::from(jan4_serial) - i64::from(jan4_weekday);
    let week_start = i64::from(serial_number) - i64::from(monday0);
    Ok(((week_start - week1_start).div_euclid(7) + 1) as i32)
}
