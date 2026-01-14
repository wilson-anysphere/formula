use std::path::PathBuf;

use formula_io::{open_workbook_model_with_password, open_workbook_with_password, Error, Workbook};

fn cryptoapi_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(
        "../formula-xls/tests/fixtures/encrypted/biff8_rc4_cryptoapi_pw_open.xls",
    )
}

#[test]
fn cryptoapi_encrypted_xls_password_errors_are_mapped_consistently() {
    let path = cryptoapi_fixture_path();

    // No password: password-capable open APIs should surface `PasswordRequired` for encrypted legacy
    // `.xls` workbooks so callers can prompt for a password.
    let err = open_workbook_with_password(&path, None).expect_err("expected password required");
    assert!(
        matches!(err, Error::PasswordRequired { .. }),
        "expected Error::PasswordRequired, got {err:?}"
    );
    let err =
        open_workbook_model_with_password(&path, None).expect_err("expected password required");
    assert!(
        matches!(err, Error::PasswordRequired { .. }),
        "expected Error::PasswordRequired, got {err:?}"
    );

    // Wrong password: map legacy `.xls` decrypt failures to InvalidPassword (matching OOXML).
    let wrong_password = "wrong-password";
    let err = open_workbook_with_password(&path, Some(wrong_password)).expect_err("expected error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );
    let msg = err.to_string();
    assert!(
        !msg.contains(wrong_password),
        "error message should not include the password; got: {msg}"
    );

    let err =
        open_workbook_model_with_password(&path, Some(wrong_password)).expect_err("expected error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );
    let msg = err.to_string();
    assert!(
        !msg.contains(wrong_password),
        "error message should not include the password; got: {msg}"
    );

    // Correct password: open should succeed.
    let wb = open_workbook_with_password(&path, Some("correct horse battery staple"))
        .expect("expected decrypted workbook to open");
    assert!(
        matches!(wb, Workbook::Xls(_)),
        "expected Workbook::Xls, got {wb:?}"
    );
    open_workbook_model_with_password(&path, Some("correct horse battery staple"))
        .expect("expected decrypted workbook to open as model");
}
