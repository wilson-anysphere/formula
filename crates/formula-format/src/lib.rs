//! Excel-compatible number/date formatting.
//!
//! This crate currently provides two layers:
//! - [`locale`] helpers for applying locale-specific decimal/thousands separators.
//! - A higher-level Excel/OOXML number format engine that understands common
//!   format codes (`#,##0.00`, `m/d/yyyy`, `0%`, `0.00E+00`, and multi-section
//!   formats like `positive;negative;zero;text`).

pub mod locale;

mod builtin;
mod datetime;
mod literal;
mod number;
mod parse;

pub use crate::builtin::builtin_format_code;
pub use crate::datetime::DateSystem;
pub use crate::parse::{FormatCode, ParseError};

pub use locale::{format_number, get_locale, NumberLocale, DE_DE, EN_US};

/// A locale definition used for formatting separators.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Locale {
    /// Decimal separator (e.g. `.` in `en-US`, `,` in many EU locales).
    pub decimal_sep: char,
    /// Thousands separator (e.g. `,` in `en-US`, `.` in `de-DE`).
    pub thousands_sep: char,
    /// Date separator used when the format code uses `/`.
    pub date_sep: char,
    /// Time separator used when the format code uses `:`.
    pub time_sep: char,
}

impl Locale {
    pub const fn en_us() -> Self {
        Self {
            decimal_sep: '.',
            thousands_sep: ',',
            date_sep: '/',
            time_sep: ':',
        }
    }

    pub const fn de_de() -> Self {
        Self {
            decimal_sep: ',',
            thousands_sep: '.',
            date_sep: '.',
            time_sep: ':',
        }
    }
}

/// Formatting options that affect how serial dates and numbers are rendered.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormatOptions {
    pub locale: Locale,
    pub date_system: DateSystem,
}

impl Default for FormatOptions {
    fn default() -> Self {
        Self {
            locale: Locale::en_us(),
            date_system: DateSystem::Excel1900,
        }
    }
}

/// A minimal value representation for formatting.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Value<'a> {
    Number(f64),
    Text(&'a str),
    Bool(bool),
    Blank,
    Error(&'a str),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AlignmentHint {
    Left,
    Center,
    Right,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormattedValue {
    pub text: String,
    pub alignment: AlignmentHint,
}

/// Format a value using an Excel number format code.
///
/// If `format_code` is `None` or empty, `"General"` is used.
pub fn format_value(value: Value<'_>, format_code: Option<&str>, options: &FormatOptions) -> FormattedValue {
    let code_str = format_code.unwrap_or("General");
    let code_str = if code_str.trim().is_empty() {
        "General"
    } else {
        code_str
    };

    let code = FormatCode::parse(code_str).unwrap_or_else(|_| FormatCode::general());

    let (text, alignment) = match value {
        Value::Blank => (String::new(), AlignmentHint::Left),
        Value::Error(err) => (err.to_string(), AlignmentHint::Center),
        Value::Text(s) => (format_text(s, &code), AlignmentHint::Left),
        Value::Bool(b) => {
            let s = if b { "TRUE" } else { "FALSE" };
            (format_text(s, &code), AlignmentHint::Center)
        }
        Value::Number(n) => {
            let section = code.select_section_for_number(n);
            if section.pattern.trim().is_empty() {
                (String::new(), AlignmentHint::Right)
            } else if crate::datetime::looks_like_datetime(section.pattern) {
                (
                    crate::datetime::format_datetime(n, section.pattern, options),
                    AlignmentHint::Right,
                )
            } else {
                let text =
                    crate::number::format_number(n, section.pattern, section.auto_negative_sign, options);
                let alignment = if crate::number::pattern_is_text(section.pattern) {
                    AlignmentHint::Left
                } else {
                    AlignmentHint::Right
                };
                (text, alignment)
            }
        }
    };

    FormattedValue { text, alignment }
}

fn format_text(text: &str, code: &FormatCode) -> String {
    if let Some(section) = code.text_section() {
        crate::literal::render_text_section(section, text)
    } else {
        text.to_string()
    }
}
