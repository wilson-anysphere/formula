//! End-to-end password-open tests for a Standard-encrypted `.xlsx` fixture.
//!
//! These are gated behind the `encrypted-workbooks` feature because decryption is optional.
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::path::PathBuf;

use formula_io::{open_workbook_model_with_password, open_workbook_with_password, Error, Workbook};
use formula_model::{CellRef, CellValue};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures")
        .join(rel)
}

#[test]
fn open_workbook_with_password_opens_standard_encrypted_xlsx() {
    let src = fixture_path("xlsx/encrypted/standard_password.xlsx");
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("standard_password.xlsx");
    std::fs::copy(&src, &path).expect("copy encrypted fixture to temp path");

    let wb = open_workbook_with_password(&path, Some("Password1234_"))
        .expect("open encrypted workbook");
    match wb {
        Workbook::Xlsx(_) => {}
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }

    let model = open_workbook_model_with_password(&path, Some("Password1234_"))
        .expect("open encrypted workbook model");
    let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("hello".to_string())
    );
}

#[test]
fn open_workbook_with_password_wrong_password_is_invalid_password() {
    let src = fixture_path("xlsx/encrypted/standard_password.xlsx");
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("standard_password.xlsx");
    std::fs::copy(&src, &path).expect("copy encrypted fixture to temp path");

    let err =
        open_workbook_with_password(&path, Some("not_the_password")).expect_err("expected error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );

    let err = open_workbook_model_with_password(&path, Some("not_the_password"))
        .expect_err("expected error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );
}
