use std::path::PathBuf;

use formula_model::{CellRef, CellValue};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../fixtures").join(rel)
}

fn xlsb_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xlsb/tests/fixtures")
        .join(rel)
}

fn xls_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xls/tests/fixtures")
        .join(rel)
}

#[test]
fn open_workbook_model_xlsx() {
    let path = fixture_path("xlsx/basic/basic.xlsx");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::Number(1.0)
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("Hello".to_string())
    );
}

#[test]
fn open_workbook_model_xlsb() {
    let path = xlsb_fixture_path("simple.xlsb");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("Hello".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::Number(42.5)
    );
    assert_eq!(sheet.formula(CellRef::from_a1("C1").unwrap()), Some("B1*2"));
}

#[test]
fn open_workbook_model_xls() {
    let path = xls_fixture_path("basic.xls");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");

    assert_eq!(workbook.sheets.len(), 2);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
    assert_eq!(workbook.sheets[1].name, "Second");

    let sheet1 = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet1.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("Hello".to_string())
    );
    assert_eq!(
        sheet1.value(CellRef::from_a1("B2").unwrap()),
        CellValue::Number(123.0)
    );
    assert_eq!(
        sheet1.formula(CellRef::from_a1("C3").unwrap()),
        Some("B2*2")
    );

    let sheet2 = workbook.sheet_by_name("Second").expect("Second missing");
    assert_eq!(
        sheet2.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("Second sheet".to_string())
    );
}
