use formula_format::{AlignmentHint, FormatOptions, Value as FmtValue};

use crate::{CellValue, HorizontalAlignment, Style};

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
    let fmt_value = match value {
        CellValue::Empty => FmtValue::Blank,
        CellValue::Number(n) => FmtValue::Number(*n),
        CellValue::String(s) => FmtValue::Text(s.as_str()),
        CellValue::Boolean(b) => FmtValue::Bool(*b),
        CellValue::Error(e) => FmtValue::Error(e.as_str()),
        CellValue::RichText(r) => FmtValue::Text(r.text.as_str()),
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
