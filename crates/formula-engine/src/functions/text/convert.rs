use crate::error::{ExcelError, ExcelResult};
use crate::value::parse_number_with_separators;

/// VALUE(text)
///
/// Implements numeric parsing with the common US-style separators (`,` thousands,
/// `.` decimal). Excel's full VALUE function also parses dates/times based on
/// locale; that will be added when the calculation engine has locale context.
pub fn value(text: &str) -> ExcelResult<f64> {
    numbervalue(text, Some('.'), Some(','))
}

/// NUMBERVALUE(number_text, [decimal_separator], [group_separator])
pub fn numbervalue(
    number_text: &str,
    decimal_separator: Option<char>,
    group_separator: Option<char>,
) -> ExcelResult<f64> {
    let decimal_separator = decimal_separator.unwrap_or('.');
    let group_separator = group_separator.unwrap_or(',');

    if decimal_separator == group_separator {
        return Err(ExcelError::Value);
    }

    parse_number_with_separators(number_text, decimal_separator, Some(group_separator))
}

// NOTE: Parsing logic is shared with implicit coercions in `crate::value`.
