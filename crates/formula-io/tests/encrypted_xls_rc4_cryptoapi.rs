use std::path::PathBuf;

use formula_io::{open_workbook_model_with_password, open_workbook_with_password, Error, Workbook};
use formula_model::CellValue;

const PASSWORD: &str = "correct horse battery staple";

fn encrypted_xls_fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xls/tests/fixtures/encrypted")
        .join("biff8_rc4_cryptoapi_pw_open.xls")
}

#[test]
fn opens_encrypted_xls_with_password() {
    let path = encrypted_xls_fixture_path();

    // Model loader path.
    let model =
        open_workbook_model_with_password(&path, Some(PASSWORD)).expect("open encrypted xls model");
    let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));

    // Full workbook loader path.
    let wb = open_workbook_with_password(&path, Some(PASSWORD)).expect("open encrypted xls");
    let Workbook::Xls(result) = wb else {
        panic!("expected Workbook::Xls, got {wb:?}");
    };
    let sheet = result.workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(42.0));
}

#[test]
fn encrypted_xls_requires_password_on_password_api() {
    let path = encrypted_xls_fixture_path();

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
}

#[test]
fn encrypted_xls_wrong_password_returns_invalid_password() {
    let path = encrypted_xls_fixture_path();

    let err = open_workbook_with_password(&path, Some("wrong password"))
        .expect_err("expected invalid password");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );

    let err = open_workbook_model_with_password(&path, Some("wrong password"))
        .expect_err("expected invalid password");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );
}

