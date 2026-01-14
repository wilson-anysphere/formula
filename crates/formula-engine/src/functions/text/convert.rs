use chrono::{DateTime, Utc};

use crate::coercion::datetime::parse_value_text;
use crate::coercion::number::parse_number_strict;
use crate::coercion::ValueLocaleConfig;
use crate::date::ExcelDateSystem;
use crate::error::{ExcelError, ExcelResult};

/// VALUE(text)
///
/// Implements a subset of Excel's VALUE function:
/// - Numeric parsing with locale-aware separators
/// - Date/time text parsing (via DATEVALUE/TIMEVALUE rules)
pub fn value(text: &str) -> ExcelResult<f64> {
    value_with_locale(
        text,
        ValueLocaleConfig::en_us(),
        Utc::now(),
        ExcelDateSystem::EXCEL_1900,
    )
}

pub fn value_with_locale(
    text: &str,
    cfg: ValueLocaleConfig,
    now_utc: DateTime<Utc>,
    system: ExcelDateSystem,
) -> ExcelResult<f64> {
    parse_value_text(text, cfg, now_utc, system)
}

/// NUMBERVALUE(number_text, [decimal_separator], [group_separator])
pub fn numbervalue(
    number_text: &str,
    decimal_separator: Option<char>,
    group_separator: Option<char>,
) -> ExcelResult<f64> {
    let decimal_separator = decimal_separator.unwrap_or('.');

    if group_separator == Some(decimal_separator) {
        return Err(ExcelError::Value);
    }

    parse_number_strict(number_text, decimal_separator, group_separator)
}
