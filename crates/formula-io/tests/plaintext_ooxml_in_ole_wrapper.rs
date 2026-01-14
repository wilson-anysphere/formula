//! Regression tests for OOXML-in-OLE containers where `EncryptedPackage` is already plaintext.
//!
//! Some pipelines wrap an OOXML ZIP package in the standard Office encryption OLE container shape
//! (`EncryptionInfo` + `EncryptedPackage`) but place the *plaintext* package bytes in
//! `EncryptedPackage` (often still with the usual 8-byte size prefix). `formula-io` supports
//! opening these via the password-aware APIs without requiring the `encrypted-workbooks` feature.
#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Write as _};

use formula_io::{open_workbook_model_with_password, open_workbook_with_password, Error, Workbook};
use formula_model::{CellRef, CellValue};

fn build_tiny_xlsx() -> Vec<u8> {
    let mut workbook = formula_model::Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");
    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
    sheet.set_value(CellRef::from_a1("A1").unwrap(), CellValue::Number(1.0));
    sheet.set_value(
        CellRef::from_a1("B1").unwrap(),
        CellValue::String("Hello".to_string()),
    );

    let mut cursor = Cursor::new(Vec::new());
    formula_io::xlsx::write_workbook_to_writer(&workbook, &mut cursor)
        .expect("write xlsx to bytes");
    cursor.into_inner()
}

fn wrap_plain_zip_in_encrypted_ooxml_ole(plain_zip: &[u8]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create OLE container");

    {
        // Minimal Agile (4.4) header; the bytes are not interpreted for the plaintext fast path.
        let mut stream = ole
            .create_stream("EncryptionInfo")
            .expect("create EncryptionInfo stream");
        stream
            .write_all(&[4, 0, 4, 0, 0, 0, 0, 0])
            .expect("write EncryptionInfo header");
    }
    {
        let mut stream = ole
            .create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
        let len = plain_zip.len() as u64;
        stream
            .write_all(&len.to_le_bytes())
            .expect("write EncryptedPackage size prefix");
        stream
            .write_all(plain_zip)
            .expect("write EncryptedPackage plaintext bytes");
    }

    ole.into_inner().into_inner()
}

fn assert_expected_contents(workbook: &formula_model::Workbook) {
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
fn opens_plaintext_ooxml_in_encrypted_ole_wrapper_when_password_is_provided() {
    let plain_xlsx = build_tiny_xlsx();
    let wrapped = wrap_plain_zip_in_encrypted_ooxml_ole(&plain_xlsx);

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("wrapped.xlsx");
    std::fs::write(&path, wrapped).expect("write wrapper file");

    // Without a password, still treat it as an encrypted OOXML wrapper.
    let err = open_workbook_with_password(&path, None).expect_err("expected error");
    if cfg!(feature = "encrypted-workbooks") {
        assert!(
            matches!(err, Error::PasswordRequired { .. }),
            "expected Error::PasswordRequired, got {err:?}"
        );
    } else {
        assert!(
            matches!(err, Error::UnsupportedEncryption { .. }),
            "expected Error::UnsupportedEncryption, got {err:?}"
        );
    }

    // With any password, open the plaintext package.
    let wb = open_workbook_with_password(&path, Some("any-password"))
        .expect("open wrapped workbook");
    match wb {
        Workbook::Xlsx(package) => {
            let bytes = package.write_to_bytes().expect("serialize package bytes");
            let model = formula_io::xlsx::read_workbook_from_reader(Cursor::new(bytes))
                .expect("parse package bytes");
            assert_expected_contents(&model);
        }
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }

    let model = open_workbook_model_with_password(&path, Some("any-password"))
        .expect("open wrapped workbook as model");
    assert_expected_contents(&model);
}

