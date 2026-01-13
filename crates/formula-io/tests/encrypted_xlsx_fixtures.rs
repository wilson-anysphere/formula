use std::path::PathBuf;

use formula_model::{CellRef, CellValue};
use formula_io::{open_workbook_model_with_password, Error};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../fixtures").join(rel)
}

fn assert_basic_values(workbook: &formula_model::Workbook) {
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
fn decrypts_and_parses_agile_encrypted_xlsx_fixture() {
    let path = fixture_path("xlsx/encrypted/agile-password.xlsx");
    let bytes = std::fs::read(&path).expect("read agile fixture");

    let dir = tempfile::tempdir().expect("temp dir");
    let tmp_path = dir.path().join("agile.xlsx");
    std::fs::write(&tmp_path, &bytes).expect("write temp fixture copy");

    let workbook =
        open_workbook_model_with_password(&tmp_path, Some("Password1234_")).expect("decrypt workbook");
    assert_basic_values(&workbook);

    let err =
        open_workbook_model_with_password(&tmp_path, Some("WrongPassword")).expect_err("wrong password");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );
}

#[test]
fn decrypts_and_parses_standard_encrypted_xlsx_fixture() {
    let path = fixture_path("xlsx/encrypted/standard-password.xlsx");
    let bytes = std::fs::read(&path).expect("read standard fixture");

    let dir = tempfile::tempdir().expect("temp dir");
    let tmp_path = dir.path().join("standard.xlsx");
    std::fs::write(&tmp_path, &bytes).expect("write temp fixture copy");

    let workbook =
        open_workbook_model_with_password(&tmp_path, Some("Password1234_")).expect("decrypt workbook");
    assert_basic_values(&workbook);

    let err =
        open_workbook_model_with_password(&tmp_path, Some("WrongPassword")).expect_err("wrong password");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );
}
