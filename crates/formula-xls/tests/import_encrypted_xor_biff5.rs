use std::path::PathBuf;

use formula_model::CellValue;

#[test]
fn imports_xor_encrypted_biff5_with_password() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted_xor_biff5.xls");

    let result =
        formula_xls::import_xls_path_with_password(&fixture_path, "xorpass").expect("import xls");

    let sheet = result
        .workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 missing");

    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));
}

#[test]
fn xor_encrypted_biff5_wrong_password_errors() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted_xor_biff5.xls");

    let err = formula_xls::import_xls_path_with_password(&fixture_path, "wrong")
        .expect_err("expected wrong password");

    assert!(matches!(err, formula_xls::ImportError::InvalidPassword));
}

#[test]
fn xor_encrypted_biff5_missing_password_reports_encrypted_workbook() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted_xor_biff5.xls");

    let err = formula_xls::import_xls_path(&fixture_path).expect_err("expected encrypted workbook");
    assert!(matches!(err, formula_xls::ImportError::EncryptedWorkbook));
}
