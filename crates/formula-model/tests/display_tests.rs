use formula_format::FormatOptions;
use formula_model::{format_cell_display, CellDisplay, CellValue, HorizontalAlignment, Style};

#[test]
fn model_formats_numbers_using_style_number_format() {
    let options = FormatOptions::default();
    let style = Style {
        number_format: Some("#,##0.00".to_string()),
        ..Style::default()
    };

    let display = format_cell_display(&CellValue::Number(1234.5), Some(&style), &options);
    assert_eq!(
        display,
        CellDisplay {
            text: "1,234.50".to_string(),
            alignment: HorizontalAlignment::Right
        }
    );
}

#[test]
fn model_aligns_bools_and_errors_center_under_general_alignment() {
    let options = FormatOptions::default();

    let display = format_cell_display(&CellValue::Boolean(true), None, &options);
    assert_eq!(display.text, "TRUE");
    assert_eq!(display.alignment, HorizontalAlignment::Center);

    let display = format_cell_display(&CellValue::Error(formula_model::ErrorValue::Div0), None, &options);
    assert_eq!(display.text, "#DIV/0!");
    assert_eq!(display.alignment, HorizontalAlignment::Center);
}

