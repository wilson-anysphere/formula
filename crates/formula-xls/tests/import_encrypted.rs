use std::path::PathBuf;

#[test]
fn errors_on_encrypted_xls_fixtures() {
    let fixtures_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("tests")
        .join("fixtures")
        .join("encrypted");

    let fixtures = [
        "biff8_xor_pw_open.xls",
        "biff8_rc4_standard_pw_open.xls",
        "biff8_rc4_cryptoapi_pw_open.xls",
    ];

    for filename in fixtures {
        let path = fixtures_dir.join(filename);
        let err = formula_xls::import_xls_path(&path)
            .expect_err(&format!("expected encrypted workbook error for {path:?}"));
        assert!(
            matches!(err, formula_xls::ImportError::EncryptedWorkbook),
            "expected ImportError::EncryptedWorkbook for {path:?}, got {err:?}"
        );

        let msg = err.to_string().to_lowercase();
        assert!(
            msg.contains("encrypted"),
            "expected error message to mention encryption; got: {msg}"
        );
        assert!(
            msg.contains("password"),
            "expected error message to mention password protection; got: {msg}"
        );
    }
}
