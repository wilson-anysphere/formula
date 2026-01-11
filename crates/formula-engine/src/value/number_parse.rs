use crate::coercion::number::parse_number_coercion;
use crate::error::ExcelResult;

/// Locale configuration for parsing numbers from text.
///
/// This is intentionally separate from [`crate::LocaleConfig`]: formula lexing has to avoid
/// ambiguous thousands separators (e.g. `,` in en-US formulas), while numeric coercion wants
/// to accept those separators.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NumberLocale {
    pub decimal_separator: char,
    pub group_separator: Option<char>,
}

impl NumberLocale {
    pub const fn en_us() -> Self {
        Self {
            decimal_separator: '.',
            group_separator: Some(','),
        }
    }

    pub const fn new(decimal_separator: char, group_separator: Option<char>) -> Self {
        Self {
            decimal_separator,
            group_separator,
        }
    }
}

pub(crate) fn parse_number(text: &str, locale: NumberLocale) -> ExcelResult<f64> {
    parse_number_coercion(text, locale.decimal_separator, locale.group_separator)
}
