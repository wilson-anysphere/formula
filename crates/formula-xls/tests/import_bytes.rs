use std::path::PathBuf;

use formula_model::CellValue;

#[test]
fn imports_unencrypted_xls_from_bytes() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("basic.xls");
    let bytes = std::fs::read(&fixture_path).expect("read fixture");

    let result = formula_xls::import_xls_bytes(&bytes).expect("import xls bytes");
    assert_eq!(result.source.path.to_string_lossy(), "<memory>");

    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1");
    assert_eq!(
        sheet.value_a1("A1").unwrap(),
        CellValue::String("Hello".to_owned())
    );
}

#[test]
fn encrypted_xls_bytes_without_password_errors() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encryption/biff8_rc4_cryptoapi_pw_open.xls");
    let bytes = std::fs::read(&fixture_path).expect("read encrypted fixture");

    let err = formula_xls::import_xls_bytes(&bytes).expect_err("expected password required");
    assert!(matches!(err, formula_xls::ImportError::EncryptedWorkbook));
}

