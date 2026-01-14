//! End-to-end decryption test for the Standard (CryptoAPI) OOXML open path.
//!
//! This is gated behind `encrypted-workbooks` because password-based decryption is optional.
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::path::PathBuf;

use formula_io::{
    open_workbook_model, open_workbook_model_with_password, open_workbook_with_options, OpenOptions,
    Workbook,
};
use formula_model::{CellRef, CellValue};
use std::io::{Cursor, Read as _, Write as _};

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

#[test]
fn decrypts_standard_fixture_via_open_workbook_with_options() {
    let plaintext_path = fixture_path("plaintext.xlsx");
    let standard_path = fixture_path("standard.xlsx");

    let plaintext = open_workbook_model(&plaintext_path).expect("open plaintext.xlsx");
    assert_expected_contents(&plaintext);

    let decrypted = open_workbook_with_options(
        &standard_path,
        OpenOptions {
            password: Some("password".to_string()),
            ..Default::default()
        },
    )
    .expect("open standard.xlsx with password");

    let decrypted_model = match decrypted {
        Workbook::Model(workbook) => workbook,
        Workbook::Xlsx(package) => {
            // Materialize the decrypted ZIP bytes and parse them into a model workbook so we can
            // validate cell contents.
            let decrypted_bytes = package
                .write_to_bytes()
                .expect("serialize decrypted Standard workbook package");
            formula_io::xlsx::read_workbook_from_reader(Cursor::new(decrypted_bytes))
                .expect("read decrypted Standard workbook package into model")
        }
        other => panic!(
            "expected Workbook::Model or Workbook::Xlsx for decrypted Standard workbook, got {other:?}"
        ),
    };
    assert_expected_contents(&decrypted_model);

    // Sanity: compare some key cell values with the known-good plaintext fixture.
    let plain_sheet = plaintext.sheet_by_name("Sheet1").expect("Sheet1 missing");
    let dec_sheet = decrypted_model.sheet_by_name("Sheet1").expect("Sheet1 missing");
    for addr in ["A1", "B1"] {
        let cref = CellRef::from_a1(addr).unwrap();
        assert_eq!(dec_sheet.value(cref), plain_sheet.value(cref), "addr={addr}");
    }
}

#[test]
fn decrypts_standard_fixture_via_streaming_reader_when_size_prefix_high_dword_is_reserved() {
    // Some producers treat the 8-byte `EncryptedPackage` length prefix as `(u32 size, u32 reserved)`
    // and may write a non-zero reserved high DWORD. Ensure the streaming Standard AES open path
    // tolerates this (it reads the prefix directly from the OLE stream, not from an in-memory
    // buffer).
    let standard_path = fixture_path("standard.xlsx");

    let file = std::fs::File::open(&standard_path).expect("open standard.xlsx fixture");
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

    // Re-wrap the streams in a fresh OLE container so we exercise the path-based streaming open.
    let cursor = Cursor::new(Vec::new());
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
    let path = tmp.path().join("standard_reserved_high_dword.xlsx");
    std::fs::write(&path, &bytes).expect("write fixture to disk");

    let decrypted = open_workbook_model_with_password(&path, Some("password"))
        .expect("open standard.xlsx via password decrypt");
    assert_expected_contents(&decrypted);
}
