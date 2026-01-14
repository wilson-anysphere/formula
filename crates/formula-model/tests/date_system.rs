use formula_format::Locale;
use formula_model::{format_cell_display, CellValue, DateSystem, Style, Workbook};

#[test]
fn workbook_date_system_defaults_to_excel_1900() {
    let wb = Workbook::new();
    assert_eq!(wb.date_system, DateSystem::Excel1900);
}

#[test]
fn workbook_deserialize_defaults_date_system_to_excel_1900() {
    let wb: Workbook = serde_json::from_value(serde_json::json!({})).unwrap();
    assert_eq!(wb.date_system, DateSystem::Excel1900);
}

#[test]
fn date_serial_formatting_depends_on_workbook_date_system() {
    let style = Style {
        number_format: Some("m/d/yyyy".to_string()),
        ..Style::default()
    };

    let wb_1900 = Workbook::new();
    let opts_1900 = wb_1900.format_options(Locale::en_us());
    let display_1900 = format_cell_display(&CellValue::Number(0.0), Some(&style), &opts_1900);
    assert_eq!(display_1900.text, "12/31/1899");

    let mut wb_1904 = Workbook::new();
    wb_1904.date_system = DateSystem::Excel1904;
    let opts_1904 = wb_1904.format_options(Locale::en_us());
    let display_1904 = format_cell_display(&CellValue::Number(0.0), Some(&style), &opts_1904);
    assert_eq!(display_1904.text, "1/1/1904");

    assert_ne!(display_1900.text, display_1904.text);
}
