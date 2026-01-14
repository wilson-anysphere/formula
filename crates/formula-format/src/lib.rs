//! Excel-compatible number/date formatting.
//!
//! This crate currently provides two layers:
//! - [`locale`] helpers for applying locale-specific decimal/thousands separators.
//! - A higher-level Excel/OOXML number format engine that understands common
//!   format codes (`#,##0.00`, `m/d/yyyy`, `0%`, `0.00E+00`, and multi-section
//!   formats like `positive;negative;zero;text`).
//!
//! For spreadsheet importers, the crate also exposes helpers for resolving Excel's
//! built-in number format IDs (`numFmtId` / `ifmt`) to format codes:
//! - [`builtin_format_code`] (canonical en-US mapping for ids 0–49)
//! - [`builtin_format_code_with_locale`] (best-effort locale-aware variants)
//! - [`locale_for_lcid`] (map LCIDs in format code tokens like `[$€-407]` to [`Locale`])
//!
//! Some importers also preserve unknown built-ins as placeholder strings like
//! `__builtin_numFmtId:14`. The formatter treats these placeholders as references
//! to the built-in table.

pub mod locale;

mod builtin;
mod cell_format;
mod datetime;
mod literal;
mod number;
mod parse;

pub use crate::builtin::builtin_format_code;
pub use crate::builtin::builtin_format_code_with_locale;
pub use crate::builtin::builtin_format_id;
pub use crate::cell_format::{cell_format_code, cell_parentheses_flag};
pub use crate::datetime::DateSystem;
pub use crate::parse::{locale_for_lcid, FormatCode, ParseError};

/// Format-related flags exposed by Excel's `CELL` function.
///
/// These correspond to:
/// - `CELL("color")`
/// - `CELL("parentheses")`
///
/// Note: this reflects *number format string* semantics only (e.g. `0;[Red]0`),
/// not Excel conditional formatting rules.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellFormatInfo {
    /// `CELL("color")`: `1` if the number format specifies a color for negative numbers.
    pub color: u8,
    /// `CELL("parentheses")`: `1` if the number format specifies parentheses for negative numbers.
    pub parentheses: u8,
}

/// Prefix used for placeholder format codes that indicate an Excel built-in number
/// format ID without embedding a concrete format string.
///
/// Some importers (notably `formula-xlsx` and `formula-xlsb`) use these
/// placeholders to preserve `numFmtId` values that were present in the source
/// file but did not include an explicit format code.
pub const BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX: &str = "__builtin_numFmtId:";

pub use locale::{
    format_number, get_locale, number_locale_from_locale, NumberLocale, DE_CH, DE_DE, EN_GB, EN_US,
    DA_DK, ES_ES, ES_MX, FR_CH, FR_FR, IT_CH, IT_IT, JA_JP, KO_KR, NB_NO, NL_BE, NL_NL, PL_PL,
    PT_BR, PT_PT, RU_RU, SV_SE, TR_TR, ZH_CN, ZH_HK, ZH_MO, ZH_SG, ZH_TW,
};

fn resolve_builtin_placeholder(code: &str) -> Option<&'static str> {
    let id = code
        .strip_prefix(BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX)?
        .trim()
        .parse::<u16>()
        .ok()?;
    builtin_format_code(id)
}

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

fn render_general_value(value: Value<'_>, options: &FormatOptions) -> RenderResult {
    match value {
        Value::Blank => RenderResult {
            text: String::new(),
            alignment: AlignmentHint::Left,
            color: None,
            layout_hint: None,
        },
        Value::Text(s) => RenderResult {
            text: s.to_string(),
            alignment: AlignmentHint::Left,
            color: None,
            layout_hint: None,
        },
        Value::Bool(b) => RenderResult {
            text: if b { "TRUE" } else { "FALSE" }.to_string(),
            alignment: AlignmentHint::Center,
            color: None,
            layout_hint: None,
        },
        Value::Error(err) => RenderResult {
            text: err.to_string(),
            alignment: AlignmentHint::Center,
            color: None,
            layout_hint: None,
        },
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

            // General numeric rendering is on the hot path, so avoid the full format
            // parser for the common None/""/"General" cases.
            let rendered = crate::number::format_number(n, "General", false, options);
            RenderResult {
                text: rendered.text,
                alignment: AlignmentHint::Right,
                color: None,
                layout_hint: None,
            }
        }
    }
}

fn format_code_is_general_without_tokens(format_code: Option<&str>) -> bool {
    let Some(code) = format_code else {
        return true;
    };

    let trimmed = code.trim();
    if trimmed.is_empty() {
        return true;
    }

    // Only treat "General" (case-insensitive) as the fast path when the format code is *exactly*
    // General. Anything tokenized (e.g. locale tags like `[$-409]General`, conditions, section
    // separators, built-in placeholders like `__builtin_numFmtId:*`, etc.) must go through the
    // parser to preserve Excel semantics.
    trimmed.eq_ignore_ascii_case("general")
}

fn resolve_builtin_placeholder_format_code(
    code: &str,
    locale: Locale,
) -> Option<std::borrow::Cow<'static, str>> {
    let Some(rest) = code.strip_prefix(BUILTIN_NUM_FMT_ID_PLACEHOLDER_PREFIX) else {
        return None;
    };

    let id = match rest.trim().parse::<u16>() {
        Ok(id) => id,
        Err(_) => return Some("General".into()),
    };

    // For the canonical en-US built-in mapping, resolve the placeholder to the
    // exact `builtin_format_code(id)` string.
    if locale == Locale::en_us() {
        if let Some(resolved) = resolve_builtin_placeholder(code) {
            return Some(resolved.into());
        }
    }

    // If the ID is one of the standard OOXML built-ins (0–49), we can resolve it directly.
    if let Some(resolved) = crate::builtin_format_code_with_locale(id, locale) {
        return Some(resolved);
    }

    // Excel uses additional reserved built-in ids (not standardized in OOXML) for
    // locale-specific date/time formats. When a file references those ids without
    // providing a formatCode, we fall back to a reasonable default so date serials
    // still render as dates instead of raw numbers.
    //
    // Note: this is best-effort; we preserve the original id separately for
    // round-tripping via the placeholder itself.
    if matches!(id, 50..=58) {
        return crate::builtin_format_code_with_locale(14, locale);
    }

    Some("General".into())
}

/// Render a value using an Excel number format code, returning additional render hints.
///
/// This is the preferred API for UI rendering because it surfaces:
/// - color override tokens (`[Red]`, `[Color10]`, …)
/// - underscore/fill layout hints used by accounting formats.
///
/// [`format_value`] remains available for callers that only need the rendered string.
pub fn render_value(value: Value<'_>, format_code: Option<&str>, options: &FormatOptions) -> RenderResult {
    if format_code_is_general_without_tokens(format_code) {
        return render_general_value(value, options);
    }

    let code_str = format_code.unwrap_or("General");
    let code_str = if code_str.trim().is_empty() {
        "General"
    } else {
        code_str
    };

    let resolved_placeholder = resolve_builtin_placeholder_format_code(code_str, options.locale);
    let code_str = resolved_placeholder.as_deref().unwrap_or(code_str);
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

/// Inspect an Excel/OOXML number format code and return the flags used by Excel's `CELL` function
/// for `CELL("color")` and `CELL("parentheses")`.
///
/// The result depends only on the format string (not on conditional formatting rules).
///
/// This helper follows Excel's section selection rules for negative values, including conditional
/// format sections like `[<0]...;...`.
pub fn cell_format_info(format_code: Option<&str>, options: &FormatOptions) -> CellFormatInfo {
    let code_str = format_code.unwrap_or("General");
    let code_str = if code_str.trim().is_empty() {
        "General"
    } else {
        code_str
    };

    let resolved_placeholder = resolve_builtin_placeholder_format_code(code_str, options.locale);
    let code_str = resolved_placeholder.as_deref().unwrap_or(code_str);
    let code = FormatCode::parse(code_str).unwrap_or_else(|_| FormatCode::general());

    // Excel's `CELL("color")` / `CELL("parentheses")` are concerned with how *negative* values are
    // formatted. Pick a representative negative number and select the format section for it.
    let negative_section = code.select_section_for_number(-1.0);

    // Excel reports 0/0 for one-section formats (where negatives use the first section and Excel
    // automatically prefixes a '-' sign).
    //
    // Even if the first section contains a color token or parentheses literals, there is no
    // explicit negative section in the format code.
    if negative_section.auto_negative_sign {
        return CellFormatInfo {
            color: 0,
            parentheses: 0,
        };
    }

    CellFormatInfo {
        color: u8::from(negative_section.color.is_some()),
        parentheses: u8::from(section_has_unescaped_parentheses(negative_section.pattern)),
    }
}

fn section_has_unescaped_parentheses(pattern: &str) -> bool {
    let mut in_quotes = false;
    let mut escape = false;
    let mut in_brackets = false;
    let mut has_open = false;
    let mut has_close = false;
    let mut chars = pattern.chars();
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

        if in_brackets {
            if ch == ']' {
                in_brackets = false;
            }
            continue;
        }

        match ch {
            '"' => in_quotes = true,
            '\\' => escape = true,
            '[' => in_brackets = true,
            // Layout tokens `_X` and `*X` consume the following character. When scanning for
            // negative-parentheses semantics, ignore the operand character even if it's a
            // parenthesis.
            '_' | '*' => {
                let _ = chars.next();
            }
            '(' => has_open = true,
            ')' => has_close = true,
            _ => {}
        }
    }

    has_open && has_close
}
