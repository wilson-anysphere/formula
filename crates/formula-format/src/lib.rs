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
pub use crate::builtin::builtin_format_code_with_locale;
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

    pub const fn fr_fr() -> Self {
        Self {
            decimal_sep: ',',
            thousands_sep: '\u{00A0}',
            date_sep: '/',
            time_sep: ':',
        }
    }

    pub const fn it_it() -> Self {
        Self {
            decimal_sep: ',',
            thousands_sep: '.',
            date_sep: '/',
            time_sep: ':',
        }
    }

    pub const fn es_es() -> Self {
        Self {
            decimal_sep: ',',
            thousands_sep: '.',
            date_sep: '/',
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

/// A color override derived from an Excel format code section token such as `[Red]` or `[Color10]`.
///
/// Format code colors are *not* part of the rendered display text; they are a rendering hint that
/// the UI can apply as an override.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColorOverride {
    /// 8-digit ARGB value (`0xAARRGGBB`).
    Argb(u32),
    /// Excel "indexed" palette entry (typically 1-56).
    Indexed(u8),
    /// Theme color (rare in format codes, but included for future compatibility).
    Theme { index: u8, tint: i16 },
}

/// Layout hints for literal spacing/fill tokens in format codes.
///
/// Excel format codes can contain:
/// - `_X` (underscore): reserve the width of the next character `X` (often used for accounting
///   formats to align parentheses). In a text-only renderer we approximate this as a single space,
///   but report an [`LiteralLayoutOp::Underscore`] so a UI can do better.
/// - `*X` (asterisk fill): repeat `X` to fill the remaining cell width. This cannot be represented
///   in a width-agnostic string, so we omit it from `text` and report an [`LiteralLayoutOp::Fill`]
///   instead.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct LiteralLayoutHint {
    pub ops: Vec<LiteralLayoutOp>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LiteralLayoutOp {
    /// The underscore token was rendered as a single space at `byte_index`.
    Underscore { byte_index: usize, width_of: char },
    /// A fill instruction should begin at `byte_index` by repeating `fill_with`.
    Fill { byte_index: usize, fill_with: char },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FormattedValue {
    pub text: String,
    pub alignment: AlignmentHint,
}

/// Full render result returned by [`render_value`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RenderResult {
    pub text: String,
    pub alignment: AlignmentHint,
    pub color: Option<ColorOverride>,
    pub layout_hint: Option<LiteralLayoutHint>,
}

/// Format a value using an Excel number format code.
///
/// If `format_code` is `None` or empty, `"General"` is used.
pub fn format_value(value: Value<'_>, format_code: Option<&str>, options: &FormatOptions) -> FormattedValue {
    let rendered = render_value(value, format_code, options);
    FormattedValue {
        text: rendered.text,
        alignment: rendered.alignment,
    }
}

/// Render a value using an Excel number format code, returning additional render hints.
///
/// This is the preferred API for UI rendering because it surfaces:
/// - color override tokens (`[Red]`, `[Color10]`, …)
/// - underscore/fill layout hints used by accounting formats.
///
/// [`format_value`] remains available for callers that only need the rendered string.
pub fn render_value(value: Value<'_>, format_code: Option<&str>, options: &FormatOptions) -> RenderResult {
    let code_str = format_code.unwrap_or("General");
    let code_str = if code_str.trim().is_empty() {
        "General"
    } else {
        code_str
    };

    let code = FormatCode::parse(code_str).unwrap_or_else(|_| FormatCode::general());

    match value {
        Value::Blank => RenderResult {
            text: String::new(),
            alignment: AlignmentHint::Left,
            color: None,
            layout_hint: None,
        },
        Value::Error(err) => RenderResult {
            text: err.to_string(),
            alignment: AlignmentHint::Center,
            color: None,
            layout_hint: None,
        },
        Value::Text(s) => {
            let (pattern, color) = code.select_section_for_text();
            let rendered = if let Some(pattern) = pattern {
                crate::literal::render_text_section(pattern, s)
            } else {
                crate::literal::RenderedText::new(s.to_string())
            };
            let layout_hint = rendered.layout_hint();
            RenderResult {
                text: rendered.text,
                alignment: AlignmentHint::Left,
                color,
                layout_hint,
            }
        }
        Value::Bool(b) => {
            let s = if b { "TRUE" } else { "FALSE" };
            let (pattern, color) = code.select_section_for_text();
            let rendered = if let Some(pattern) = pattern {
                crate::literal::render_text_section(pattern, s)
            } else {
                crate::literal::RenderedText::new(s.to_string())
            };
            let layout_hint = rendered.layout_hint();
            RenderResult {
                text: rendered.text,
                alignment: AlignmentHint::Center,
                color,
                layout_hint,
            }
        }
        Value::Number(n) => {
            if !n.is_finite() {
                // Excel does not have NaN/Infinity numeric values; treat them as #NUM!.
                return RenderResult {
                    text: "#NUM!".to_string(),
                    alignment: AlignmentHint::Center,
                    color: None,
                    layout_hint: None,
                };
            }

            let section = code.select_section_for_number(n);
            let mut section_options = *options;
            if let Some(locale) = section.locale_override {
                section_options.locale = locale;
            }
            if crate::datetime::looks_like_datetime(section.pattern) {
                let rendered = crate::datetime::format_datetime(n, section.pattern, &section_options);
                let layout_hint = rendered.layout_hint();
                RenderResult {
                    text: rendered.text,
                    alignment: AlignmentHint::Right,
                    color: section.color,
                    layout_hint,
                }
            } else {
                let rendered = crate::number::format_number(
                    n,
                    section.pattern,
                    section.auto_negative_sign,
                    &section_options,
                );
                let alignment = if crate::number::pattern_is_text(section.pattern) {
                    AlignmentHint::Left
                } else {
                    AlignmentHint::Right
                };
                let layout_hint = rendered.layout_hint();
                RenderResult {
                    text: rendered.text,
                    alignment,
                    color: section.color,
                    layout_hint,
                }
            }
        }
    }
}
