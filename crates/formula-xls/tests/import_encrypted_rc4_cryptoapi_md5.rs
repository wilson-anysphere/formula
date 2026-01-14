use std::path::PathBuf;

use formula_model::CellValue;

const PASSWORD: &str = "password";
const WRONG_PASSWORD: &str = "wrong password";

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted")
        .join("biff8_rc4_cryptoapi_md5_pw_open.xls")
}

#[test]
fn decrypts_rc4_cryptoapi_md5_biff8_xls() {
    let result = formula_xls::import_xls_path_with_password(fixture_path(), Some(PASSWORD))
        .expect("expected decrypt + import to succeed");
    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));
}

#[test]
fn rc4_cryptoapi_md5_wrong_password_errors() {
    let err = formula_xls::import_xls_path_with_password(fixture_path(), Some(WRONG_PASSWORD))
        .expect_err("expected wrong password error");
    assert!(matches!(err, formula_xls::ImportError::InvalidPassword));
}

