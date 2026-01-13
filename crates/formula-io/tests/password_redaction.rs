use formula_io::open_workbook_with_password;

const WRONG_PASSWORD: &str = "hunter2";

#[test]
fn wrong_password_error_does_not_leak_password() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(
        "../formula-xls/tests/fixtures/encrypted/biff8_rc4_cryptoapi_pw_open.xls",
    );

    let err = open_workbook_with_password(&path, Some(WRONG_PASSWORD))
        .expect_err("expected wrong password");
    assert!(
        matches!(err, formula_io::Error::InvalidPassword { .. }),
        "expected wrong-password path to map to Error::InvalidPassword, got {err:?}"
    );

    let display = err.to_string();
    let debug = format!("{err:?}");

    assert!(!display.contains(WRONG_PASSWORD));
    assert!(!debug.contains(WRONG_PASSWORD));
}
