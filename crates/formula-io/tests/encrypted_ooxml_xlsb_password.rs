//! End-to-end decryption tests for Office-encrypted XLSB workbooks (Agile encryption).
//!
//! These are gated behind the `encrypted-workbooks` feature because decryption is optional.
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::path::PathBuf;

use formula_io::{
    open_workbook_model_with_password, open_workbook_with_password, Error, Workbook,
};
use formula_model::{CellRef, CellValue};

fn xlsb_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xlsb/tests/fixtures")
        .join(rel)
}

fn encrypt_ooxml_package_bytes(bytes: &[u8], password: &str) -> Vec<u8> {
    // Keep the KDF work factor low in tests so encryption/decryption stays fast while still
    // exercising the full code path.
    let mut opts = formula_office_crypto::EncryptOptions::default();
    opts.spin_count = 1_000;
    formula_office_crypto::encrypt_package_to_ole(bytes, password, opts)
        .expect("encrypt xlsb payload into OLE EncryptedPackage wrapper")
}

#[test]
fn opens_encrypted_ooxml_xlsb_with_password() {
    let xlsb_bytes = std::fs::read(xlsb_fixture_path("simple.xlsb")).expect("read xlsb fixture");
    let password = "secret";
    let encrypted = encrypt_ooxml_package_bytes(&xlsb_bytes, password);

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&path, &encrypted).expect("write encrypted workbook");

    let workbook =
        open_workbook_with_password(&path, Some(password)).expect("open with password");
    assert!(
        matches!(workbook, Workbook::Xlsb(_)),
        "expected Workbook::Xlsb, got {workbook:?}"
    );

    let model = open_workbook_model_with_password(&path, Some(password))
        .expect("open model with password");
    let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");

    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("Hello".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::Number(42.5)
    );
    assert_eq!(sheet.formula(CellRef::from_a1("C1").unwrap()), Some("B1*2"));
}

#[test]
fn wrong_password_errors_for_encrypted_ooxml_xlsb() {
    let xlsb_bytes = std::fs::read(xlsb_fixture_path("simple.xlsb")).expect("read xlsb fixture");
    let encrypted = encrypt_ooxml_package_bytes(&xlsb_bytes, "correct");

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&path, &encrypted).expect("write encrypted workbook");

    let err =
        open_workbook_with_password(&path, Some("wrong")).expect_err("expected error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected InvalidPassword, got {err:?}"
    );

    let err =
        open_workbook_model_with_password(&path, Some("wrong")).expect_err("expected error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected InvalidPassword, got {err:?}"
    );
}
