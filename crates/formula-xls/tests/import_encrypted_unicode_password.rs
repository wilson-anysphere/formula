use std::path::PathBuf;

use formula_model::CellValue;

#[test]
fn decrypts_encrypted_xls_with_non_ascii_password() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted_rc4cryptoapi_non_ascii_password.xls");

    let result = formula_xls::import_xls_path_with_password(&fixture_path, Some("pässwörd"))
        .expect("import xls");

    let sheet1 = result
        .workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 missing");

    assert_eq!(sheet1.value_a1("A1").unwrap(), CellValue::Number(42.0));
}

#[test]
fn encrypted_xls_password_normalization_is_not_applied() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted_rc4cryptoapi_non_ascii_password.xls");

    // Visually-similar but different Unicode string: `ä`/`ö` decomposed into base letter + combining diaeresis.
    let decomposed = "pa\u{0308}sswo\u{0308}rd";

    let err = formula_xls::import_xls_path_with_password(&fixture_path, Some(decomposed))
        .expect_err("expected invalid password");
    assert!(matches!(err, formula_xls::ImportError::InvalidPassword));
}
