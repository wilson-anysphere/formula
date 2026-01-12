use std::path::PathBuf;

use formula_io::{open_workbook, save_workbook, Workbook};
use formula_model::{CellRef, CellValue, DateSystem, SheetVisibility};

fn xls_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xls/tests/fixtures")
        .join(rel)
}

#[test]
fn opens_basic_xls_fixture_and_exports_to_xlsx_with_expected_cells() {
    let path = xls_fixture_path("basic.xls");
    let wb = open_workbook(&path).expect("open workbook");

    match &wb {
        Workbook::Xls(_) => {}
        other => panic!("expected Workbook::Xls, got {other:?}"),
    }

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("export workbook");

    let bytes = std::fs::read(&out_path).expect("read exported workbook");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load exported workbook");

    assert_eq!(doc.workbook.sheets.len(), 2);

    let sheet1 = doc.workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet1.value_a1("A1").unwrap(),
        CellValue::String("Hello".to_string())
    );
    assert_eq!(sheet1.value_a1("B2").unwrap(), CellValue::Number(123.0));
    assert_eq!(
        sheet1.formula(CellRef::from_a1("C3").unwrap()),
        Some("B2*2")
    );

    let sheet2 = doc.workbook.sheet_by_name("Second").expect("Second missing");
    assert_eq!(
        sheet2.value_a1("A1").unwrap(),
        CellValue::String("Second sheet".to_string())
    );
}

#[test]
fn xls_export_preserves_date_system_1904() {
    let path = xls_fixture_path("date_system_1904.xls");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("export workbook");

    let bytes = std::fs::read(&out_path).expect("read exported workbook");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load exported workbook");

    assert_eq!(doc.workbook.date_system, DateSystem::Excel1904);
}

#[test]
fn xls_export_preserves_sheet_visibility() {
    let path = xls_fixture_path("hidden.xls");
    let wb = open_workbook(&path).expect("open workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("export workbook");

    let bytes = std::fs::read(&out_path).expect("read exported workbook");
    let doc = formula_xlsx::load_from_bytes(&bytes).expect("load exported workbook");

    let visible = doc.workbook.sheet_by_name("Visible").expect("Visible sheet missing");
    assert_eq!(visible.visibility, SheetVisibility::Visible);

    let hidden = doc.workbook.sheet_by_name("Hidden").expect("Hidden sheet missing");
    assert_eq!(hidden.visibility, SheetVisibility::Hidden);
}

