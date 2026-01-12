use std::path::PathBuf;

use formula_model::{CellRef, CellValue};
use std::io::Write as _;

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
fn open_workbook_model_xlsx_reads_formulas() {
    let path = fixture_path("xlsx/formulas/formulas.xlsx");
    let workbook = formula_io::open_workbook_model(&path).expect("open workbook model");

    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet.formula(CellRef::from_a1("C1").unwrap()), Some("A1+B1"));
}

#[test]
fn open_workbook_model_sniffs_extensionless_xlsx() {
    let path = fixture_path("xlsx/basic/basic.xlsx");
    let bytes = std::fs::read(&path).expect("read fixture");

    let mut tmp = tempfile::Builder::new()
        .prefix("basic_xlsx_")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(&bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
}

#[test]
fn open_workbook_model_sniffs_xlsx_with_wrong_extension() {
    let path = fixture_path("xlsx/basic/basic.xlsx");
    let bytes = std::fs::read(&path).expect("read fixture");

    let mut tmp = tempfile::Builder::new()
        .prefix("basic_xlsx_wrong_ext_")
        .suffix(".xls")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(&bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
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
fn open_workbook_model_sniffs_extensionless_xlsb() {
    let path = xlsb_fixture_path("simple.xlsb");
    let bytes = std::fs::read(&path).expect("read fixture");

    let mut tmp = tempfile::Builder::new()
        .prefix("simple_xlsb_")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(&bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
}

#[test]
fn open_workbook_model_sniffs_xlsb_with_wrong_extension() {
    let path = xlsb_fixture_path("simple.xlsb");
    let bytes = std::fs::read(&path).expect("read fixture");

    let mut tmp = tempfile::Builder::new()
        .prefix("simple_xlsb_wrong_ext_")
        .suffix(".xlsx")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(&bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
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

#[test]
fn open_workbook_model_sniffs_extensionless_xls() {
    let path = xls_fixture_path("basic.xls");
    let bytes = std::fs::read(&path).expect("read fixture");

    let mut tmp = tempfile::Builder::new()
        .prefix("basic_xls_")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(&bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 2);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
    assert_eq!(workbook.sheets[1].name, "Second");
}

#[test]
fn open_workbook_model_sniffs_xls_with_wrong_extension() {
    let path = xls_fixture_path("basic.xls");
    let bytes = std::fs::read(&path).expect("read fixture");

    let mut tmp = tempfile::Builder::new()
        .prefix("basic_xls_wrong_ext_")
        .suffix(".xlsx")
        .tempfile()
        .expect("tempfile");
    tmp.write_all(&bytes).expect("write tempfile");

    let workbook = formula_io::open_workbook_model(tmp.path()).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 2);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
    assert_eq!(workbook.sheets[1].name, "Second");
}

#[test]
fn open_workbook_model_csv() {
    let dir = tempfile::tempdir().expect("temp dir");
    let csv_path = dir.path().join("data.csv");
    std::fs::write(&csv_path, "col1,col2\n1,hello\n2,world\n").expect("write csv");

    let workbook = formula_io::open_workbook_model(&csv_path).expect("open workbook model");
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "data");

    let sheet = workbook.sheet_by_name("data").expect("data sheet missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("hello".to_string())
    );
    assert_eq!(sheet.value_a1("A2").unwrap(), CellValue::Number(2.0));
    assert_eq!(
        sheet.value_a1("B2").unwrap(),
        CellValue::String("world".to_string())
    );
}
