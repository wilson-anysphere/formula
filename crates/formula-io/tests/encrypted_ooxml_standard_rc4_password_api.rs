//! Ensure the password-based open APIs can decrypt Standard/CryptoAPI RC4 encrypted OOXML.
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::io::{Read as _, Write as _};
use std::path::PathBuf;

use formula_io::{open_workbook_model_with_password, open_workbook_with_password, Error, Workbook};
use formula_model::CellValue;

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(name)
}

#[test]
fn open_workbook_with_password_decrypts_standard_rc4_fixture() {
    let path = fixture_path("standard-rc4.xlsx");

    let wb = open_workbook_with_password(&path, Some("password"))
        .expect("decrypt + open standard-rc4.xlsx via password API");

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
fn open_workbook_model_with_password_decrypts_standard_rc4_fixture() {
    let path = fixture_path("standard-rc4.xlsx");

    let workbook = open_workbook_model_with_password(&path, Some("password"))
        .expect("decrypt + open standard-rc4.xlsx via model password API");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("Hello".to_string())
    );
}

#[test]
fn standard_rc4_wrong_password_is_invalid_password_for_password_api() {
    let path = fixture_path("standard-rc4.xlsx");

    let err = open_workbook_with_password(&path, Some("wrong-password"))
        .expect_err("expected wrong password to error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );

    let err = open_workbook_model_with_password(&path, Some("wrong-password"))
        .expect_err("expected wrong password to error");
    assert!(
        matches!(err, Error::InvalidPassword { .. }),
        "expected Error::InvalidPassword, got {err:?}"
    );
}

#[test]
fn open_workbook_model_with_password_decrypts_standard_rc4_fixture_when_size_prefix_high_dword_is_reserved(
) {
    // Regression: some producers treat the 8-byte `EncryptedPackage` size prefix as `(u32 size, u32 reserved)`
    // and may write a non-zero reserved high DWORD.
    let standard_path = fixture_path("standard-rc4.xlsx");

    let file = std::fs::File::open(&standard_path).expect("open standard-rc4.xlsx fixture");
    let mut ole = cfb::CompoundFile::open(file).expect("parse OLE");

    let mut encryption_info = Vec::new();
    ole.open_stream("EncryptionInfo")
        .or_else(|_| ole.open_stream("/EncryptionInfo"))
        .expect("open EncryptionInfo")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let mut encrypted_package = Vec::new();
    ole.open_stream("EncryptedPackage")
        .or_else(|_| ole.open_stream("/EncryptedPackage"))
        .expect("open EncryptedPackage")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    assert!(
        encrypted_package.len() >= 8,
        "EncryptedPackage too short (missing size prefix)"
    );

    // Set the high DWORD (reserved) to a non-zero value.
    encrypted_package[4..8].copy_from_slice(&1u32.to_le_bytes());

    // Re-wrap the streams in a fresh OLE container so we exercise the path-based open APIs.
    let cursor = std::io::Cursor::new(Vec::new());
    let mut out_ole = cfb::CompoundFile::create(cursor).expect("create OLE");
    out_ole
        .create_stream("EncryptionInfo")
        .expect("create EncryptionInfo")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo");
    out_ole
        .create_stream("EncryptedPackage")
        .expect("create EncryptedPackage")
        .write_all(&encrypted_package)
        .expect("write EncryptedPackage");

    let bytes = out_ole.into_inner().into_inner();
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("standard_rc4_reserved_high_dword.xlsx");
    std::fs::write(&path, &bytes).expect("write fixture to disk");

    let workbook = open_workbook_model_with_password(&path, Some("password"))
        .expect("decrypt + open standard-rc4.xlsx with reserved high DWORD");

    let sheet = workbook.sheet_by_name("Sheet1").expect("Sheet1 missing");
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(
        sheet.value_a1("B1").unwrap(),
        CellValue::String("Hello".to_string())
    );
}
