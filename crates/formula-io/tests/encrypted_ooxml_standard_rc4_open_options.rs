//! Ensure `open_workbook_with_options` can decrypt Standard/CryptoAPI RC4 encrypted OOXML.
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::path::PathBuf;

use formula_io::{
    open_workbook_model_with_password, open_workbook_with_options, Error, OpenOptions, Workbook,
};
use formula_model::{CellRef, CellValue};

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(name)
}

#[test]
fn open_workbook_with_options_decrypts_standard_rc4_fixture() {
    let path = fixture_path("standard-rc4.xlsx");

    let wb = open_workbook_with_options(
        &path,
        OpenOptions {
            password: Some("password".to_string()),
            ..Default::default()
        },
    )
    .expect("decrypt + open standard-rc4.xlsx");

    let Workbook::Xlsx(pkg) = wb else {
        panic!("expected Workbook::Xlsx, got {wb:?}");
    };

    assert!(
        pkg.read_part("xl/workbook.xml")
            .expect("read xl/workbook.xml")
            .is_some(),
        "expected decrypted package to contain xl/workbook.xml"
    );
}

#[test]
fn open_workbook_with_options_standard_rc4_wrong_password_is_invalid_password() {
    let path = fixture_path("standard-rc4.xlsx");

    let err = open_workbook_with_options(
        &path,
        OpenOptions {
            password: Some("wrong-password".to_string()),
            ..Default::default()
        },
    )
    .expect_err("expected wrong password error");

    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );
}

#[test]
fn open_workbook_model_with_password_decrypts_standard_rc4_fixture() {
    let path = fixture_path("standard-rc4.xlsx");

    let workbook = open_workbook_model_with_password(&path, Some("password"))
        .expect("decrypt + open model standard-rc4.xlsx");
    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::Number(1.0)
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("Hello".to_string())
    );
}

#[test]
fn open_workbook_model_with_password_standard_rc4_wrong_password_is_invalid_password() {
    let path = fixture_path("standard-rc4.xlsx");

    let err = open_workbook_model_with_password(&path, Some("wrong-password"))
        .expect_err("expected wrong password error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );
}
