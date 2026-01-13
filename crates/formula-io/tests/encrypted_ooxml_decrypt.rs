//! End-to-end decryption tests for Office-encrypted OOXML workbooks (Agile encryption).
//!
//! These are gated behind the `encrypted-workbooks` feature because decryption is optional.
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::io::{Cursor, Write as _};
use std::path::{Path, PathBuf};

use ms_offcrypto_writer::Ecma376AgileWriter;
use zip::write::FileOptions;

use formula_io::{
    detect_workbook_format, open_workbook_model, open_workbook_model_with_password,
    open_workbook_with_password, Error, Workbook, WorkbookFormat,
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

fn xlsb_fixture_bytes() -> Vec<u8> {
    let path = Path::new(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../formula-xlsb/tests/fixtures/simple.xlsb"
    ));
    std::fs::read(path).expect("read xlsb fixture bytes")
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

#[test]
fn open_workbook_with_password_decrypts_agile_encrypted_xlsb() {
    let password = "password";
    let plain_xlsb = xlsb_fixture_bytes();
    let encrypted_cfb = encrypt_zip_with_password(&plain_xlsb, password);

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&path, &encrypted_cfb).expect("write encrypted file");

    let wb = open_workbook_with_password(&path, Some(password)).expect("open decrypted workbook");
    match wb {
        Workbook::Xlsb(wb) => {
            assert_eq!(wb.sheet_metas().len(), 1);
            let sheet = wb.read_sheet(0).expect("read sheet");
            assert!(
                sheet.cells.iter().any(|c| c.row == 0 && c.col == 0),
                "expected to see cell A1"
            );
        }
        other => panic!("expected Workbook::Xlsb, got {other:?}"),
    }
}

#[test]
fn open_workbook_model_with_password_decrypts_agile_encrypted_xlsb() {
    let password = "password";
    let plain_xlsb = xlsb_fixture_bytes();
    let encrypted_cfb = encrypt_zip_with_password(&plain_xlsb, password);

    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("encrypted.xlsb");
    std::fs::write(&path, &encrypted_cfb).expect("write encrypted file");

    let model = open_workbook_model_with_password(&path, Some(password)).expect("open model");
    let sheet = model.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(
        sheet.value(CellRef::from_a1("A1").unwrap()),
        CellValue::String("Hello".to_string())
    );
    assert_eq!(
        sheet.value(CellRef::from_a1("B1").unwrap()),
        CellValue::Number(42.5)
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

fn open_decrypted_package_bytes_with_password(path: &Path, password: &str) -> Vec<u8> {
    let wb = open_workbook_with_password(path, Some(password))
        .unwrap_or_else(|err| panic!("open encrypted workbook {path:?} failed: {err:?}"));
    match wb {
        Workbook::Xlsx(package) => package
            .write_to_bytes()
            .expect("serialize decrypted workbook package to bytes"),
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

fn assert_has_vba_project(decrypted: &[u8]) {
    let archive = zip::ZipArchive::new(Cursor::new(decrypted)).expect("open decrypted ZIP");
    let mut found = false;
    for name in archive.file_names() {
        if name.eq_ignore_ascii_case("xl/vbaProject.bin") {
            found = true;
            break;
        }
    }
    assert!(
        found,
        "expected decrypted package to contain xl/vbaProject.bin"
    );
}

fn assert_detects_xlsm(decrypted: &[u8]) {
    let tmp = tempfile::tempdir().expect("temp dir");
    let path = tmp.path().join("book.xlsm");
    std::fs::write(&path, decrypted).expect("write decrypted workbook bytes");
    assert_eq!(
        detect_workbook_format(&path).expect("detect workbook format"),
        WorkbookFormat::Xlsm
    );
}

#[test]
fn decrypts_agile_fixture_with_correct_password() {
    let plaintext_path = fixture_path("plaintext.xlsx");
    let agile_path = fixture_path("agile.xlsx");
    let agile_empty_password_path = fixture_path("agile-empty-password.xlsx");

    let plaintext = open_workbook_model(&plaintext_path).expect("open plaintext.xlsx");
    assert_expected_contents(&plaintext);

    let agile = open_model_with_password(&agile_path, "password");
    assert_expected_contents(&agile);

    let agile_empty = open_model_with_password(&agile_empty_password_path, "");
    assert_expected_contents(&agile_empty);
}

#[test]
fn decrypts_macro_enabled_xlsm_fixtures_with_correct_password() {
    let plaintext_basic_path = fixture_path("plaintext-basic.xlsm");
    let agile_basic_path = fixture_path("agile-basic.xlsm");
    let standard_basic_path = fixture_path("standard-basic.xlsm");

    assert_eq!(
        detect_workbook_format(&plaintext_basic_path).expect("detect plaintext-basic.xlsm"),
        WorkbookFormat::Xlsm
    );
    let plaintext_basic_bytes =
        std::fs::read(&plaintext_basic_path).expect("read plaintext-basic.xlsm");
    assert_has_vba_project(&plaintext_basic_bytes);

    let agile_basic = open_model_with_password(&agile_basic_path, "password");
    assert!(
        !agile_basic.sheets.is_empty(),
        "expected decrypted macro workbook to have at least one sheet"
    );
    let agile_basic_bytes =
        open_decrypted_package_bytes_with_password(&agile_basic_path, "password");
    assert_has_vba_project(&agile_basic_bytes);
    assert_detects_xlsm(&agile_basic_bytes);

    let standard_basic = open_model_with_password(&standard_basic_path, "password");
    assert!(
        !standard_basic.sheets.is_empty(),
        "expected decrypted macro workbook to have at least one sheet"
    );
    let standard_basic_bytes =
        open_decrypted_package_bytes_with_password(&standard_basic_path, "password");
    assert_has_vba_project(&standard_basic_bytes);
    assert_detects_xlsm(&standard_basic_bytes);
}

#[test]
fn errors_on_missing_password_for_empty_password_fixture() {
    let agile_empty_password_path = fixture_path("agile-empty-password.xlsx");

    let err = open_workbook_model_with_password(&agile_empty_password_path, None)
        .expect_err("expected missing password to error");
    assert!(
        matches!(err, Error::PasswordRequired { .. }),
        "expected Error::PasswordRequired, got {err:?}"
    );
}

#[test]
fn errors_on_wrong_password_fixtures() {
    let agile_path = fixture_path("agile.xlsx");
    let agile_empty_password_path = fixture_path("agile-empty-password.xlsx");
    let standard_path = fixture_path("standard.xlsx");
    let agile_unicode_path = fixture_path("agile-unicode.xlsx");
    let agile_basic_path = fixture_path("agile-basic.xlsm");
    let standard_basic_path = fixture_path("standard-basic.xlsm");

    for path in [
        &agile_path,
        &agile_empty_password_path,
        &standard_path,
        &agile_unicode_path,
        &agile_basic_path,
        &standard_basic_path,
    ] {
        assert!(
            matches!(
                open_workbook_model_with_password(path, Some("wrong-password")),
                Err(Error::InvalidPassword { .. })
            ),
            "expected InvalidPassword error for {path:?}"
        );
    }
}

#[test]
fn decrypts_agile_unicode_password() {
    let path = fixture_path("agile-unicode.xlsx");
    let wb = open_model_with_password(&path, "pässwörd");
    assert_expected_contents(&wb);
}

#[test]
fn agile_unicode_password_different_normalization_fails() {
    // NFC password is "pässwörd" (U+00E4, U+00F6). NFD decomposes those into combining marks.
    let nfd = "pa\u{0308}sswo\u{0308}rd";
    assert_ne!(
        nfd, "pässwörd",
        "strings should differ before UTF-16 encoding"
    );

    let path = fixture_path("agile-unicode.xlsx");
    assert!(
        matches!(
            open_workbook_model_with_password(&path, Some(nfd)),
            Err(Error::InvalidPassword { .. })
        ),
        "expected InvalidPassword error for NFD-normalized password"
    );
}
