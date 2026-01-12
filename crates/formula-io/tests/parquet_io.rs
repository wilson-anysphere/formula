#![cfg(feature = "parquet")]

use std::path::PathBuf;

use formula_io::{open_workbook, save_workbook};
use formula_model::CellValue;

fn parquet_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../packages/data-io/test/fixtures")
        .join(rel)
}

#[test]
fn parquet_import_can_export_to_xlsx() {
    let parquet_path = parquet_fixture_path("simple.parquet");
    let wb = open_workbook(&parquet_path).expect("open parquet workbook");

    let dir = tempfile::tempdir().expect("temp dir");
    let out_path = dir.path().join("export.xlsx");
    save_workbook(&wb, &out_path).expect("save workbook as xlsx");

    let file = std::fs::File::open(&out_path).expect("open exported xlsx");
    let exported = formula_xlsx::read_workbook_from_reader(file).expect("read exported workbook");
    let sheet = exported
        .sheet_by_name("simple")
        .expect("expected worksheet name to match file stem");

    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("Alice".to_string())
    );
    assert_eq!(sheet.value_a1("C2").unwrap(), CellValue::Boolean(false));
    assert_eq!(sheet.value_a1("D3").unwrap(), CellValue::Number(3.75));
}

