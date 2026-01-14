use formula_format::{AlignmentHint, FormatOptions, Value as FmtValue};

use crate::{CellValue, HorizontalAlignment, Style, Workbook};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CellDisplay {
    pub text: String,
    pub alignment: HorizontalAlignment,
}

/// Format a [`CellValue`] using an Excel number format code from a [`Style`].
///
/// This helper is intended for UI rendering: it provides the user-visible
/// string and an alignment hint derived from Excel's "General" alignment rules.
pub fn format_cell_display(
    value: &CellValue,
    style: Option<&Style>,
    options: &FormatOptions,
) -> CellDisplay {
    // Some `CellValue` variants (e.g. Record/Entity) map to a synthesized display
    // string rather than a direct reference into the input value.
    let mut display_buf: Option<String> = None;
    let fmt_value = match value {
        CellValue::Empty => FmtValue::Blank,
        CellValue::Number(n) => FmtValue::Number(*n),
        CellValue::String(s) => FmtValue::Text(s.as_str()),
        CellValue::Boolean(b) => FmtValue::Bool(*b),
        CellValue::Error(e) => FmtValue::Error(e.as_str()),
        CellValue::RichText(r) => FmtValue::Text(r.text.as_str()),
        CellValue::Entity(entity) => FmtValue::Text(entity.display_value.as_str()),
        CellValue::Record(record) => {
            let display = record
                .display_field
                .as_deref()
                .and_then(|field| record.get_field_case_insensitive(field))
                .and_then(|value| match value {
                    CellValue::Empty => Some(""),
                    CellValue::String(s) => Some(s.as_str()),
                    CellValue::Number(n) => {
                        display_buf = Some(
                            formula_format::format_value(FmtValue::Number(*n), None, options).text,
                        );
                        display_buf.as_deref()
                    }
                    CellValue::Boolean(b) => {
                        display_buf = Some(if *b { "TRUE" } else { "FALSE" }.to_string());
                        display_buf.as_deref()
                    }
                    CellValue::Error(e) => Some(e.as_str()),
                    CellValue::RichText(rt) => Some(rt.text.as_str()),
                    CellValue::Entity(entity) => Some(entity.display_value.as_str()),
                    CellValue::Record(record) => {
                        display_buf = Some(record.to_string());
                        display_buf.as_deref()
                    }
                    CellValue::Image(image) => image
                        .alt_text
                        .as_deref()
                        .filter(|s| !s.is_empty())
                        .or_else(|| {
                            display_buf = Some("[Image]".to_string());
                            display_buf.as_deref()
                        }),
                    _ => None,
                })
                .or_else(|| {
                    (!record.display_value.is_empty()).then_some(record.display_value.as_str())
                })
                .unwrap_or("");
            FmtValue::Text(display)
        }
        CellValue::Image(image) => {
            let display = image
                .alt_text
                .as_deref()
                .filter(|s| !s.is_empty())
                .unwrap_or_else(|| {
                    display_buf = Some("[Image]".to_string());
                    display_buf.as_deref().unwrap_or("")
                });
            FmtValue::Text(display)
        }
        // For now arrays/spills are UI-rendered elsewhere.
        CellValue::Array(_) | CellValue::Spill(_) => FmtValue::Blank,
    };

    let number_format = style.and_then(|s| s.number_format.as_deref());
    let formatted = formula_format::format_value(fmt_value, number_format, options);

    let explicit_alignment = style
        .and_then(|s| s.alignment.as_ref())
        .and_then(|a| a.horizontal);

    let alignment = match explicit_alignment {
        Some(HorizontalAlignment::Left) => HorizontalAlignment::Left,
        Some(HorizontalAlignment::Center) => HorizontalAlignment::Center,
        Some(HorizontalAlignment::Right) => HorizontalAlignment::Right,
        Some(HorizontalAlignment::Fill) => HorizontalAlignment::Fill,
        Some(HorizontalAlignment::Justify) => HorizontalAlignment::Justify,
        Some(HorizontalAlignment::General) | None => match formatted.alignment {
            AlignmentHint::Left => HorizontalAlignment::Left,
            AlignmentHint::Center => HorizontalAlignment::Center,
            AlignmentHint::Right => HorizontalAlignment::Right,
        },
    };

    CellDisplay {
        text: formatted.text,
        alignment,
    }
}

/// Format a [`CellValue`] using this workbook's date system for serial dates.
pub fn format_cell_display_in_workbook(
    workbook: &Workbook,
    value: &CellValue,
    style: Option<&Style>,
    locale: formula_format::Locale,
) -> CellDisplay {
    let options = workbook.format_options(locale);
    format_cell_display(value, style, &options)
}
