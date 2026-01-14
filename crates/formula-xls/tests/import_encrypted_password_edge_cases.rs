use std::path::PathBuf;

use formula_model::CellValue;

#[test]
fn decrypts_rc4_standard_with_long_password_truncation() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted")
        .join("biff8_rc4_standard_pw_open_long_password.xls");

    // Excel legacy `.xls` encryption uses only the first 15 UTF-16 code units of the password.
    // The 16th character is ignored.
    let full = "0123456789abcdef"; // 16 chars
    let truncated = "0123456789abcde"; // first 15 chars

    // Both variants should decrypt successfully (Excel treats them as equivalent).
    for password in [full, truncated] {
        let result = formula_xls::import_xls_path_with_password(&fixture_path, password)
            .expect("decrypt and import");
        let sheet1 = result
            .workbook
            .sheet_by_name("Sheet1")
            .expect("Sheet1 missing");
        assert_eq!(
            sheet1.value_a1("A1").unwrap(),
            CellValue::String("Hello".to_owned())
        );
        assert_eq!(sheet1.value_a1("B2").unwrap(), CellValue::Number(123.0));
    }

    // A password that differs within the first 15 characters must fail.
    let wrong = "1123456789abcdef";
    let err = formula_xls::import_xls_path_with_password(&fixture_path, wrong)
        .expect_err("expected wrong password to fail");
    assert!(
        matches!(err, formula_xls::ImportError::InvalidPassword),
        "expected ImportError::InvalidPassword, got {err:?}"
    );
}

#[test]
fn decrypts_rc4_standard_with_empty_password() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted")
        .join("biff8_rc4_standard_pw_open_empty_password.xls");

    // Empty passwords are permitted by the legacy RC4 key derivation algorithm.
    let result = formula_xls::import_xls_path_with_password(&fixture_path, "")
        .expect("decrypt and import with empty password");
    let sheet2 = result
        .workbook
        .sheet_by_name("Second")
        .expect("Second missing");
    assert_eq!(
        sheet2.value_a1("A1").unwrap(),
        CellValue::String("Second sheet".to_owned())
    );

    let err = formula_xls::import_xls_path_with_password(&fixture_path, "not-empty")
        .expect_err("expected wrong password to fail");
    assert!(
        matches!(err, formula_xls::ImportError::InvalidPassword),
        "expected ImportError::InvalidPassword, got {err:?}"
    );
}

#[test]
fn rc4_cryptoapi_does_not_truncate_password_to_15_chars() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted")
        .join("biff8_rc4_cryptoapi_pw_open.xls");

    let full = "correct horse battery staple";
    let truncated = "correct horse b"; // first 15 chars

    // CryptoAPI encryption uses the full password string (unlike legacy RC4 standard which truncates
    // to 15 UTF-16 code units).
    let result = formula_xls::import_xls_path_with_password(&fixture_path, full)
        .expect("decrypt and import");
    let sheet1 = result.workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet1.value_a1("A1").unwrap(), CellValue::Number(42.0));

    let err = formula_xls::import_xls_path_with_password(&fixture_path, truncated)
        .expect_err("expected truncated password to fail");
    assert!(
        matches!(err, formula_xls::ImportError::InvalidPassword),
        "expected ImportError::InvalidPassword, got {err:?}"
    );
}
