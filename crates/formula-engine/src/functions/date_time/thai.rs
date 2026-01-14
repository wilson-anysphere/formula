use crate::date::{serial_to_ymd, ExcelDateSystem};
use crate::error::{ExcelError, ExcelResult};

const THAI_MONTHS: [&str; 12] = [
    "มกราคม",
    "กุมภาพันธ์",
    "มีนาคม",
    "เมษายน",
    "พฤษภาคม",
    "มิถุนายน",
    "กรกฎาคม",
    "สิงหาคม",
    "กันยายน",
    "ตุลาคม",
    "พฤศจิกายน",
    "ธันวาคม",
];

// Excel's THAIDAYOFWEEK returns Thai weekday names. We use the common full forms with the "วัน"
// prefix to match Thai Excel's built-in display strings.
const THAI_WEEKDAYS: [&str; 7] = [
    "วันอาทิตย์",
    "วันจันทร์",
    "วันอังคาร",
    "วันพุธ",
    "วันพฤหัสบดี",
    "วันศุกร์",
    "วันเสาร์",
];

const BUDDHIST_ERA_OFFSET: i32 = 543;

/// THAIDAYOFWEEK(serial_number)
pub fn thaidayofweek(serial_number: i32, system: ExcelDateSystem) -> ExcelResult<&'static str> {
    // Match WEEKDAY(serial_number,1): 1=Sunday..7=Saturday.
    let weekday = super::weekday(serial_number, Some(1), system)?;
    let idx = usize::try_from(weekday - 1).map_err(|_| ExcelError::Num)?;
    THAI_WEEKDAYS.get(idx).copied().ok_or(ExcelError::Num)
}

/// THAIMONTHOFYEAR(serial_number)
pub fn thaimonthofyear(serial_number: i32, system: ExcelDateSystem) -> ExcelResult<&'static str> {
    let date = serial_to_ymd(serial_number, system)?;
    let idx = usize::from(date.month.saturating_sub(1));
    THAI_MONTHS.get(idx).copied().ok_or(ExcelError::Num)
}

/// THAIYEAR(serial_number)
pub fn thaiyear(serial_number: i32, system: ExcelDateSystem) -> ExcelResult<i32> {
    let date = serial_to_ymd(serial_number, system)?;
    date.year
        .checked_add(BUDDHIST_ERA_OFFSET)
        .ok_or(ExcelError::Num)
}
