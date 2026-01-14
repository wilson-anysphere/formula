#![cfg(not(target_arch = "wasm32"))]

use std::io::{Cursor, Read, Write};
use std::path::Path;

use formula_xlsx::XlsxError;
use zip::write::FileOptions;
use zip::CompressionMethod;
use zip::ZipWriter;

fn fixture_path(rel: &str) -> std::path::PathBuf {
    Path::new(concat!(env!("CARGO_MANIFEST_DIR"), "/../../fixtures/encrypted/ooxml/")).join(rel)
}

fn build_word_ooxml_zip() -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let opts = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

    zip.start_file("[Content_Types].xml", opts)
        .expect("start content types");
    zip.write_all(b"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>")
        .expect("write content types");

    zip.start_file("word/document.xml", opts)
        .expect("start word document");
    zip.write_all(
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
    )
    .expect("write word document");

    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn encrypted_ole_correct_password_opens_workbook() {
    let bytes = std::fs::read(fixture_path("agile.xlsx")).expect("read encrypted fixture");

    let pkg = formula_xlsx::load_from_encrypted_ole_bytes(&bytes, "password")
        .expect("decrypt + open as xlsx package");
    assert!(
        pkg.part("xl/workbook.xml").is_some(),
        "expected decrypted package to contain xl/workbook.xml"
    );

    let workbook =
        formula_xlsx::read_workbook_from_encrypted_reader(Cursor::new(bytes), "password")
            .expect("decrypt + parse workbook");

    assert!(
        workbook.sheets.iter().any(|sheet| sheet.name == "Sheet1"),
        "expected decrypted workbook to contain Sheet1"
    );
}

#[test]
fn encrypted_ole_standard_correct_password_opens_workbook() {
    let bytes = std::fs::read(fixture_path("standard.xlsx")).expect("read encrypted fixture");

    let pkg = formula_xlsx::load_from_encrypted_ole_bytes(&bytes, "password")
        .expect("decrypt + open as xlsx package");
    assert!(
        pkg.part("xl/workbook.xml").is_some(),
        "expected decrypted package to contain xl/workbook.xml"
    );

    let workbook =
        formula_xlsx::read_workbook_from_encrypted_reader(Cursor::new(bytes), "password")
            .expect("decrypt + parse workbook");
    assert!(
        workbook.sheets.iter().any(|sheet| sheet.name == "Sheet1"),
        "expected decrypted workbook to contain Sheet1"
    );
}

#[test]
fn encrypted_ole_refuses_non_excel_decrypted_payloads() {
    let zip_bytes = build_word_ooxml_zip();
    let ole_bytes = formula_office_crypto::encrypt_package_to_ole(
        &zip_bytes,
        "password",
        formula_office_crypto::EncryptOptions::default(),
    )
    .expect("encrypt Word-like OOXML into EncryptedPackage OLE");

    let err = formula_xlsx::load_from_encrypted_ole_bytes(&ole_bytes, "password")
        .expect_err("expected Word OOXML payload to be rejected");
    match err {
        XlsxError::InvalidEncryptedWorkbook(msg) => {
            assert!(
                msg.contains("Word document"),
                "expected Word hint in error message, got: {msg}"
            );
        }
        other => panic!("expected InvalidEncryptedWorkbook, got {other:?}"),
    }
}

#[test]
fn encrypted_ole_wrong_password_returns_invalid_password() {
    let bytes = std::fs::read(fixture_path("agile.xlsx")).expect("read encrypted fixture");

    let err = formula_xlsx::read_workbook_from_encrypted_reader(Cursor::new(bytes), "wrong")
        .expect_err("expected invalid password to error");
    assert!(
        matches!(err, XlsxError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}

#[test]
fn encrypted_ole_integrity_failure_is_reported_as_invalid_password() {
    let bytes = std::fs::read(fixture_path("agile.xlsx")).expect("read encrypted fixture");

    let cursor = Cursor::new(bytes);
    let mut ole_in = cfb::CompoundFile::open(cursor).expect("open OLE container");

    let mut encryption_info = Vec::new();
    ole_in
        .open_stream("EncryptionInfo")
        .expect("open EncryptionInfo")
        .read_to_end(&mut encryption_info)
        .expect("read EncryptionInfo");

    let mut encrypted_package = Vec::new();
    ole_in
        .open_stream("EncryptedPackage")
        .expect("open EncryptedPackage")
        .read_to_end(&mut encrypted_package)
        .expect("read EncryptedPackage");

    // Flip a ciphertext byte so the Agile `dataIntegrity` HMAC check fails (but the password
    // verifier can still succeed). This should be treated as retryable "invalid password" by the
    // higher-level API surface.
    assert!(
        encrypted_package.len() > 8,
        "fixture EncryptedPackage should include ciphertext bytes"
    );
    let idx = encrypted_package.len() - 1;
    encrypted_package[idx] ^= 0x01;

    // Re-wrap into a minimal OLE container with the modified ciphertext.
    let cursor = Cursor::new(Vec::new());
    let mut ole_out = cfb::CompoundFile::create(cursor).expect("create OLE container");
    ole_out
        .create_stream("EncryptionInfo")
        .expect("create EncryptionInfo")
        .write_all(&encryption_info)
        .expect("write EncryptionInfo");
    ole_out
        .create_stream("EncryptedPackage")
        .expect("create EncryptedPackage")
        .write_all(&encrypted_package)
        .expect("write EncryptedPackage");
    let bytes = ole_out.into_inner().into_inner();

    let err = formula_xlsx::read_workbook_from_encrypted_reader(Cursor::new(bytes), "password")
        .expect_err("expected integrity failure to error");
    assert!(
        matches!(err, XlsxError::InvalidPassword),
        "expected InvalidPassword, got {err:?}"
    );
}
