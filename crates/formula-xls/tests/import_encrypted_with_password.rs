use std::path::PathBuf;

use formula_model::CellValue;

const PASSWORD: &str = "correct horse battery staple";

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted")
        .join("biff8_rc4_cryptoapi_pw_open.xls")
}

#[test]
fn imports_rc4_cryptoapi_encrypted_xls_bytes_with_password() {
    let bytes = std::fs::read(fixture_path()).expect("read fixture");
    let result = formula_xls::import_xls_bytes_with_password(&bytes, PASSWORD)
        .expect("import encrypted xls bytes");

    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));
}

#[test]
fn encrypted_xls_bytes_wrong_password_errors() {
    let bytes = std::fs::read(fixture_path()).expect("read fixture");
    let err = formula_xls::import_xls_bytes_with_password(&bytes, "wrong password")
        .expect_err("expected invalid password");
    assert!(matches!(err, formula_xls::ImportError::InvalidPassword));
}
