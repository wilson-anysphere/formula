use std::path::PathBuf;

use formula_format::{Locale, Value};
use formula_model::{CellRef, CellValue, DateSystem};

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join(name)
}

fn assert_cell_renders(
    result: &formula_xls::XlsImportResult,
    sheet_name: &str,
    a1: &str,
    expected_serial: f64,
    expected_text: &str,
) {
    let sheet = result
        .workbook
        .sheet_by_name(sheet_name)
        .unwrap_or_else(|| panic!("{sheet_name} missing"));

    let cell_ref = CellRef::from_a1(a1).unwrap();
    let cell = sheet.cell(cell_ref).unwrap_or_else(|| panic!("{a1} missing"));

    let serial = match cell.value {
        CellValue::Number(n) => n,
        _ => panic!("{a1} expected numeric serial"),
    };
    assert_eq!(serial, expected_serial, "{a1} serial mismatch");

    let format_code = result
        .workbook
        .styles
        .get(cell.style_id)
        .and_then(|s| s.number_format.as_deref())
        .unwrap_or_else(|| panic!("{a1} missing number format"));

    let opts = result.workbook.format_options(Locale::en_us());
    let rendered = formula_format::format_value(Value::Number(serial), Some(format_code), &opts).text;
    assert_eq!(rendered, expected_text, "{a1} rendered mismatch");
}

#[test]
fn renders_dates_using_excel_1900_date_system() {
    let result =
        formula_xls::import_xls_path(fixture_path("date_system_1900.xls")).expect("import xls");
    assert_eq!(result.workbook.date_system, DateSystem::Excel1900);

    assert_cell_renders(&result, "Dates", "A1", 1.0, "1/1/00");
    assert_cell_renders(&result, "Dates", "A2", 1.5, "1/1/00 12:00:00");
    assert_cell_renders(&result, "Dates", "A3", 0.5, "12:00:00");
    assert_cell_renders(&result, "Dates", "A4", 1.5, "36:00:00");

    // Guard against silently ignoring the workbook date system.
    let sheet = result.workbook.sheet_by_name("Dates").expect("Dates missing");
    let cell = sheet.cell(CellRef::from_a1("A1").unwrap()).expect("A1 missing");
    let serial = match cell.value {
        CellValue::Number(n) => n,
        _ => panic!("A1 expected numeric serial"),
    };
    let format_code = result
        .workbook
        .styles
        .get(cell.style_id)
        .and_then(|s| s.number_format.as_deref())
        .expect("A1 missing format");
    let mut wrong_opts = result.workbook.format_options(Locale::en_us());
    wrong_opts.date_system = formula_format::DateSystem::Excel1904;
    let wrong_rendered =
        formula_format::format_value(Value::Number(serial), Some(format_code), &wrong_opts).text;
    assert_ne!(wrong_rendered, "1/1/00");
}

#[test]
fn renders_dates_using_excel_1904_date_system() {
    let result =
        formula_xls::import_xls_path(fixture_path("date_system_1904.xls")).expect("import xls");
    assert_eq!(result.workbook.date_system, DateSystem::Excel1904);

    assert_cell_renders(&result, "Dates", "A1", 0.0, "1/1/04");
    assert_cell_renders(&result, "Dates", "A2", 1.5, "1/2/04 12:00:00");
    assert_cell_renders(&result, "Dates", "A3", 0.5, "12:00:00");
    assert_cell_renders(&result, "Dates", "A4", 1.5, "36:00:00");

    // Guard against silently ignoring the workbook date system.
    let sheet = result.workbook.sheet_by_name("Dates").expect("Dates missing");
    let cell = sheet.cell(CellRef::from_a1("A1").unwrap()).expect("A1 missing");
    let serial = match cell.value {
        CellValue::Number(n) => n,
        _ => panic!("A1 expected numeric serial"),
    };
    let format_code = result
        .workbook
        .styles
        .get(cell.style_id)
        .and_then(|s| s.number_format.as_deref())
        .expect("A1 missing format");
    let mut wrong_opts = result.workbook.format_options(Locale::en_us());
    wrong_opts.date_system = formula_format::DateSystem::Excel1900;
    let wrong_rendered =
        formula_format::format_value(Value::Number(serial), Some(format_code), &wrong_opts).text;
    assert_ne!(wrong_rendered, "1/1/04");
}

