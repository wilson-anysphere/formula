//! End-to-end decryption tests for Office-encrypted OOXML workbooks (Agile encryption).
//!
//! These are gated behind the `encrypted-workbooks` feature because decryption is optional.
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::io::{Cursor, Write as _};
use std::path::{Path, PathBuf};

use ms_offcrypto_writer::Ecma376AgileWriter;
use zip::write::FileOptions;

use formula_io::{
    open_workbook_model, open_workbook_model_with_password, open_workbook_with_password, Error,
    Workbook,
};
use formula_model::{CellRef, CellValue};

fn build_tiny_zip() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(cursor);
    writer
        .start_file(
            "hello.txt",
            FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored),
        )
        .expect("start zip file");
    writer.write_all(b"hello").expect("write zip contents");
    writer.finish().expect("finish zip").into_inner()
}

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

fn encrypt_zip_with_password(plain_zip: &[u8], password: &str) -> Vec<u8> {
    let mut cursor = Cursor::new(Vec::new());
    let mut agile =
        Ecma376AgileWriter::create(&mut rand::rng(), password, &mut cursor).expect("create agile");
    agile
        .write_all(plain_zip)
        .expect("write plaintext zip to agile writer");
    agile.finalize().expect("finalize agile writer");
    cursor.into_inner()
}

#[test]
fn open_workbook_with_password_decrypts_agile_encrypted_package() {
    let password = "correct horse battery staple";
    let plain_zip = build_tiny_zip();
    let encrypted_cfb = encrypt_zip_with_password(&plain_zip, password);

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsx");
    std::fs::write(&path, &encrypted_cfb).expect("write encrypted file");

    // Missing password => prompt.
    let err = open_workbook_with_password(&path, None).expect_err("expected PasswordRequired");
    assert!(
        matches!(err, Error::PasswordRequired { .. }),
        "expected Error::PasswordRequired, got {err:?}"
    );

    // Wrong password => invalid password.
    let err =
        open_workbook_with_password(&path, Some("wrong-password")).expect_err("expected error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );

    // Correct password => decrypted ZIP is passed through to XlsxPackage.
    let wb = open_workbook_with_password(&path, Some(password)).expect("open decrypted workbook");
    match wb {
        Workbook::Xlsx(package) => {
            let contents = package.part("hello.txt").expect("hello.txt missing in zip");
            assert_eq!(contents, b"hello");
        }
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

#[test]
fn open_workbook_model_with_password_decrypts_agile_encrypted_xlsx() {
    let password = "password";
    let plain_xlsx = build_tiny_xlsx();
    let encrypted_cfb = encrypt_zip_with_password(&plain_xlsx, password);

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsx");
    std::fs::write(&path, &encrypted_cfb).expect("write encrypted file");

    let model =
        open_workbook_model_with_password(&path, Some(password)).expect("open decrypted model");
    let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::Number(1.0)
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::String("Hello".to_string())
    );
}

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(rel)
}

fn assert_expected_contents(workbook: &formula_model::Workbook) {
    assert_eq!(workbook.sheets.len(), 1, "expected exactly one sheet");
    assert_eq!(workbook.sheets[0].name, "Sheet1");

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

fn open_model_with_password(path: &Path, password: &str) -> formula_model::Workbook {
    open_workbook_model_with_password(path, Some(password))
        .unwrap_or_else(|err| panic!("open encrypted workbook {path:?} failed: {err:?}"))
}

#[test]
fn decrypts_agile_fixture_with_correct_password() {
    let plaintext_path = fixture_path("plaintext.xlsx");
    let agile_path = fixture_path("agile.xlsx");

    let plaintext = open_workbook_model(&plaintext_path).expect("open plaintext.xlsx");
    assert_expected_contents(&plaintext);

    let agile = open_model_with_password(&agile_path, "password");
    assert_expected_contents(&agile);
}

#[test]
fn errors_on_wrong_password_fixtures() {
    let agile_path = fixture_path("agile.xlsx");
    let standard_path = fixture_path("standard.xlsx");

    for path in [&agile_path, &standard_path] {
        assert!(
            matches!(
                open_workbook_model_with_password(path, Some("wrong-password")),
                Err(Error::InvalidPassword { .. })
            ),
            "expected InvalidPassword error for {path:?}"
        );
    }
}
