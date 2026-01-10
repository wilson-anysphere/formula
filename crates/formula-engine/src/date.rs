use crate::error::{ExcelError, ExcelResult};

/// Excel workbook date system.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelDateSystem {
    /// 1900 date system (Windows). Optionally emulates the Lotus 1-2-3 leap-year bug
    /// where Excel treats 1900 as a leap year and includes a non-existent
    /// `1900-02-29` as serial day 60.
    Excel1900 { lotus_compat: bool },
    /// 1904 date system (Mac). Serial day 0 is `1904-01-01`.
    Excel1904,
}

impl ExcelDateSystem {
    /// Excel's default 1900 date system with the Lotus compatibility bug enabled.
    pub const EXCEL_1900: ExcelDateSystem = ExcelDateSystem::Excel1900 { lotus_compat: true };
}

/// Calendar date representation that can model Excel's fictitious `1900-02-29`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ExcelDate {
    pub year: i32,
    pub month: u8,
    pub day: u8,
}

impl ExcelDate {
    pub const fn new(year: i32, month: u8, day: u8) -> Self {
        Self { year, month, day }
    }
}

pub fn ymd_to_serial(date: ExcelDate, system: ExcelDateSystem) -> ExcelResult<i32> {
    match system {
        ExcelDateSystem::Excel1900 { lotus_compat } => {
            if lotus_compat && date == ExcelDate::new(1900, 2, 29) {
                return Ok(60);
            }

            validate_ymd(date)?;

            let days = days_from_civil(date.year, date.month, date.day);
            let base = days_from_civil(1899, 12, 31);
            let mut serial = days - base;

            if lotus_compat {
                let march_1_1900 = days_from_civil(1900, 3, 1);
                if days >= march_1_1900 {
                    serial += 1;
                }
            }

            Ok(i32::try_from(serial).map_err(|_| ExcelError::Num)?)
        }
        ExcelDateSystem::Excel1904 => {
            validate_ymd(date)?;
            let days = days_from_civil(date.year, date.month, date.day);
            let base = days_from_civil(1904, 1, 1);
            let serial = days - base;
            Ok(i32::try_from(serial).map_err(|_| ExcelError::Num)?)
        }
    }
}

pub fn serial_to_ymd(serial: i32, system: ExcelDateSystem) -> ExcelResult<ExcelDate> {
    match system {
        ExcelDateSystem::Excel1900 { lotus_compat } => {
            if lotus_compat && serial == 60 {
                return Ok(ExcelDate::new(1900, 2, 29));
            }

            let serial = i64::from(serial);
            let serial = if lotus_compat && serial > 60 {
                serial - 1
            } else {
                serial
            };

            let base = days_from_civil(1899, 12, 31);
            let days = base + serial;
            let (y, m, d) = civil_from_days(days);
            Ok(ExcelDate::new(y, m, d))
        }
        ExcelDateSystem::Excel1904 => {
            let base = days_from_civil(1904, 1, 1);
            let days = base + i64::from(serial);
            let (y, m, d) = civil_from_days(days);
            Ok(ExcelDate::new(y, m, d))
        }
    }
}

fn validate_ymd(date: ExcelDate) -> ExcelResult<()> {
    if !(1..=12).contains(&date.month) {
        return Err(ExcelError::Num);
    }
    if date.day == 0 {
        return Err(ExcelError::Num);
    }

    let dim = days_in_month(date.year, date.month);
    if date.day > dim {
        return Err(ExcelError::Num);
    }
    Ok(())
}

fn is_leap_year(year: i32) -> bool {
    (year % 4 == 0 && year % 100 != 0) || (year % 400 == 0)
}

fn days_in_month(year: i32, month: u8) -> u8 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if is_leap_year(year) {
                29
            } else {
                28
            }
        }
        _ => 0,
    }
}

// Howard Hinnant's "civil" algorithms for proleptic Gregorian calendars.
// https://howardhinnant.github.io/date_algorithms.html#days_from_civil
fn days_from_civil(year: i32, month: u8, day: u8) -> i64 {
    let y = i64::from(year) - if month <= 2 { 1 } else { 0 };
    let m = i64::from(month);
    let d = i64::from(day);

    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let mp = m + if m > 2 { -3 } else { 9 };
    let doy = (153 * mp + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146097 + doe - 719468
}

fn civil_from_days(days: i64) -> (i32, u8, u8) {
    let z = days + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = mp + if mp < 10 { 3 } else { -9 };
    let year = y + if m <= 2 { 1 } else { 0 };

    (
        i32::try_from(year).unwrap_or(i32::MAX),
        u8::try_from(m).unwrap_or(u8::MAX),
        u8::try_from(d).unwrap_or(u8::MAX),
    )
}
