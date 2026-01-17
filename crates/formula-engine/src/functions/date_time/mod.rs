use std::collections::HashSet;

use chrono::{DateTime, Utc};

use crate::coercion::datetime::{parse_datevalue_text, parse_timevalue_text};
use crate::coercion::ValueLocaleConfig;
use crate::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use crate::error::{ExcelError, ExcelResult};

pub(crate) mod thai;

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
pub fn timevalue(time_text: &str, cfg: ValueLocaleConfig) -> ExcelResult<f64> {
    parse_timevalue_text(time_text, cfg)
}

/// DATEVALUE(date_text)
pub fn datevalue(
    date_text: &str,
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
    parse_datevalue_text(date_text, cfg, now_utc, system)
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

fn is_last_day_of_month(serial_number: i32, system: ExcelDateSystem) -> ExcelResult<bool> {
    let date = crate::date::serial_to_ymd(serial_number, system)?;
    let next_serial = serial_number.checked_add(1).ok_or(ExcelError::Num)?;
    let next = crate::date::serial_to_ymd(next_serial, system)?;
    Ok(date.year != next.year || date.month != next.month)
}

/// DAYS360(start_date, end_date, [method])
///
/// Returns the number of days between two dates based on a 360-day year (twelve 30-day months).
///
/// When `method` is `false` (or omitted), Excel uses the "US/NASD" method:
/// - If `start_date` is the last day of the month, treat it as day 30.
/// - If `end_date` is the last day of the month:
///   - If the adjusted `start_date` day is < 30, treat `end_date` as the 1st of the next month.
///   - Otherwise treat `end_date` as day 30 of the same month.
///
/// When `method` is `true`, Excel uses the European method:
/// - Dates that fall on the 31st are treated as day 30.
pub fn days360(
    start_date: i32,
    end_date: i32,
    method: bool,
    system: ExcelDateSystem,
) -> ExcelResult<i64> {
    let start = crate::date::serial_to_ymd(start_date, system)?;
    let end = crate::date::serial_to_ymd(end_date, system)?;

    let y1 = i64::from(start.year);
    let m1 = i64::from(start.month);
    let mut d1 = i64::from(start.day);

    let mut y2 = i64::from(end.year);
    let mut m2 = i64::from(end.month);
    let mut d2 = i64::from(end.day);

    if method {
        // European method: only adjust 31st-of-month dates.
        if d1 == 31 {
            d1 = 30;
        }
        if d2 == 31 {
            d2 = 30;
        }
    } else {
        // US/NASD method: adjust end-of-month rules, including February.
        if is_last_day_of_month(start_date, system)? {
            d1 = 30;
        }

        if is_last_day_of_month(end_date, system)? {
            if d1 < 30 {
                // Move to the 1st of the next month.
                d2 = 1;
                if m2 == 12 {
                    m2 = 1;
                    y2 = y2.checked_add(1).ok_or(ExcelError::Num)?;
                } else {
                    m2 = m2.checked_add(1).ok_or(ExcelError::Num)?;
                }
            } else {
                d2 = 30;
            }
        }
    }

    let year_diff = y2.checked_sub(y1).ok_or(ExcelError::Num)?;
    let month_diff = m2.checked_sub(m1).ok_or(ExcelError::Num)?;
    let day_diff = d2.checked_sub(d1).ok_or(ExcelError::Num)?;

    let year_term = year_diff.checked_mul(360).ok_or(ExcelError::Num)?;
    let month_term = month_diff.checked_mul(30).ok_or(ExcelError::Num)?;
    let total = year_term
        .checked_add(month_term)
        .and_then(|v| v.checked_add(day_diff))
        .ok_or(ExcelError::Num)?;
    Ok(total)
}

fn whole_years_between(
    start_date: i32,
    end_date: i32,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
    debug_assert!(end_date >= start_date);
    let start = crate::date::serial_to_ymd(start_date, system)?;
    let end = crate::date::serial_to_ymd(end_date, system)?;
    let mut years = end.year - start.year;
    if years <= 0 {
        return Ok(0);
    }

    let months = years.checked_mul(12).ok_or(ExcelError::Num)?;
    if edate(start_date, months, system)? > end_date {
        years = years.saturating_sub(1);
    }
    Ok(years)
}

/// YEARFRAC(start_date, end_date, [basis])
pub fn yearfrac(
    start_date: i32,
    end_date: i32,
    basis: i32,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    if !(0..=4).contains(&basis) {
        return Err(ExcelError::Num);
    }

    match basis {
        0 => Ok(days360(start_date, end_date, false, system)? as f64 / 360.0),
        2 => Ok((i64::from(end_date) - i64::from(start_date)) as f64 / 360.0),
        3 => Ok((i64::from(end_date) - i64::from(start_date)) as f64 / 365.0),
        4 => Ok(days360(start_date, end_date, true, system)? as f64 / 360.0),
        1 => {
            if start_date == end_date {
                return Ok(0.0);
            }

            let mut start = start_date;
            let mut end = end_date;
            let mut sign = 1.0;
            if end < start {
                sign = -1.0;
                std::mem::swap(&mut start, &mut end);
            }

            let years = whole_years_between(start, end, system)?;
            let months = years.checked_mul(12).ok_or(ExcelError::Num)?;
            let anniversary = edate(start, months, system)?;
            let remaining_days = i64::from(end) - i64::from(anniversary);
            if remaining_days == 0 {
                return Ok(sign * (years as f64));
            }

            let next_anniversary = edate(anniversary, 12, system)?;
            let denom_days = i64::from(next_anniversary) - i64::from(anniversary);
            if denom_days == 0 {
                return Err(ExcelError::Num);
            }

            Ok(sign * (years as f64 + (remaining_days as f64) / (denom_days as f64)))
        }
        _ => {
            debug_assert!(false, "YEARFRAC basis should have been validated: {basis}");
            Err(ExcelError::Num)
        }
    }
}

fn full_months_between(
    start_date: i32,
    end_date: i32,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
    debug_assert!(end_date >= start_date);
    let start = crate::date::serial_to_ymd(start_date, system)?;
    let end = crate::date::serial_to_ymd(end_date, system)?;

    let years = i64::from(end.year) - i64::from(start.year);
    let mut months = years
        .checked_mul(12)
        .and_then(|v| v.checked_add(i64::from(end.month) - i64::from(start.month)))
        .ok_or(ExcelError::Num)?;
    if end.day < start.day {
        months = months.checked_sub(1).ok_or(ExcelError::Num)?;
    }
    i32::try_from(months).map_err(|_| ExcelError::Num)
}

/// DATEDIF(start_date, end_date, unit)
pub fn datedif(
    start_date: i32,
    end_date: i32,
    unit: &str,
    system: ExcelDateSystem,
) -> ExcelResult<i64> {
    if start_date > end_date {
        return Err(ExcelError::Num);
    }

    let unit = unit.trim();
    if unit.is_empty() {
        return Err(ExcelError::Num);
    }

    if unit.eq_ignore_ascii_case("D") {
        return Ok(i64::from(end_date) - i64::from(start_date));
    }

    if unit.eq_ignore_ascii_case("Y")
        || unit.eq_ignore_ascii_case("M")
        || unit.eq_ignore_ascii_case("YM")
        || unit.eq_ignore_ascii_case("MD")
        || unit.eq_ignore_ascii_case("YD")
    {
        let full_months = full_months_between(start_date, end_date, system)?;
        let years = full_months / 12;

        if unit.eq_ignore_ascii_case("Y") {
            return Ok(i64::from(years));
        }
        if unit.eq_ignore_ascii_case("M") {
            return Ok(i64::from(full_months));
        }
        if unit.eq_ignore_ascii_case("YM") {
            return Ok(i64::from(full_months.rem_euclid(12)));
        }
        if unit.eq_ignore_ascii_case("MD") {
            let anchor = edate(start_date, full_months, system)?;
            return Ok(i64::from(end_date) - i64::from(anchor));
        }
        // `YD`
        let anchor = edate(start_date, years.saturating_mul(12), system)?;
        return Ok(i64::from(end_date) - i64::from(anchor));
    }

    Err(ExcelError::Num)
}

/// WEEKDAY(serial_number, [return_type])
pub fn weekday(
    serial_number: i32,
    return_type: Option<i32>,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
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
pub fn workday(
    start_date: i32,
    days: i32,
    holidays: Option<&[i32]>,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
    workday_intl(start_date, days, DEFAULT_WEEKEND_MASK, holidays, system)
}

/// NETWORKDAYS(start_date, end_date, [holidays])
pub fn networkdays(
    start_date: i32,
    end_date: i32,
    holidays: Option<&[i32]>,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
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
pub fn weeknum(
    serial_number: i32,
    return_type: Option<i32>,
    system: ExcelDateSystem,
) -> ExcelResult<i32> {
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
