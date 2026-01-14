//! Round-trip preservation tests for Office-encrypted OOXML OLE containers.
//!
//! Specifically: ensure we preserve non-package OLE metadata streams (e.g.
//! `\u{0005}SummaryInformation`) when decrypting and then re-encrypting.
#![cfg(all(feature = "encrypted-workbooks", not(target_arch = "wasm32")))]

use std::io::{Cursor, Read as _, Write as _};
use std::path::PathBuf;

use ms_offcrypto_writer::Ecma376AgileWriter;
use rand::{rngs::StdRng, SeedableRng as _};
use zip::write::FileOptions;

use formula_io::{
    open_workbook_with_password, open_workbook_with_password_and_preserved_ole, Error, Workbook,
};

const SUMMARY_INFORMATION: &str = "\u{0005}SummaryInformation";

fn ooxml_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/encrypted/ooxml")
        .join(rel)
}

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

fn encrypt_zip_with_password(plain_zip: &[u8], password: &str) -> Vec<u8> {
    let mut cursor = Cursor::new(Vec::new());
    let mut rng = StdRng::from_seed([0u8; 32]);
    let mut agile =
        Ecma376AgileWriter::create(&mut rng, password, &mut cursor).expect("create agile");
    agile
        .write_all(plain_zip)
        .expect("write plaintext zip to agile writer");
    agile.finalize().expect("finalize agile writer");
    cursor.into_inner()
}

#[test]
fn preserves_extra_ole_metadata_streams_on_encrypt_roundtrip() {
    let password = "correct horse battery staple";
    let plain_zip = build_tiny_zip();
    let encrypted_cfb = encrypt_zip_with_password(&plain_zip, password);

    // Add a dummy metadata stream to the OLE wrapper. Excel commonly stores SummaryInformation /
    // DocumentSummaryInformation streams alongside the encryption streams.
    let dummy_bytes = b"dummy summary information bytes";
    let cursor = Cursor::new(encrypted_cfb);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open encrypted cfb");
    ole.create_stream(SUMMARY_INFORMATION)
        .expect("create SummaryInformation stream")
        .write_all(dummy_bytes)
        .expect("write SummaryInformation bytes");
    let encrypted_with_extra = ole.into_inner().into_inner();

    let tmp = tempfile::tempdir().expect("tempdir");
    let in_path = tmp.path().join("input.xlsx");
    std::fs::write(&in_path, &encrypted_with_extra).expect("write encrypted input");

    let opened = open_workbook_with_password_and_preserved_ole(&in_path, Some(password))
        .expect("open encrypted workbook with preservation");
    assert!(
        opened.preserved_ole.is_some(),
        "expected preserved OLE entries to be captured"
    );

    let out_path = tmp.path().join("output.xlsx");
    opened
        .save_preserving_encryption(&out_path, password)
        .expect("save with encryption preserved");

    // Verify the extra stream is preserved byte-for-byte in the output OLE container.
    let out_bytes = std::fs::read(&out_path).expect("read output bytes");
    let mut out_ole = cfb::CompoundFile::open(Cursor::new(out_bytes)).expect("open output cfb");
    let mut stream = out_ole
        .open_stream(SUMMARY_INFORMATION)
        .expect("open SummaryInformation stream");
    let mut got = Vec::new();
    stream.read_to_end(&mut got).expect("read SummaryInformation");
    assert_eq!(got, dummy_bytes);

    // Sanity check: the output encrypted container should still decrypt/open.
    let reopened =
        open_workbook_with_password(&out_path, Some(password)).expect("reopen encrypted output");
    match reopened {
        Workbook::Xlsx(package) => {
            let contents = package
                .read_part("hello.txt")
                .expect("read hello.txt")
                .expect("hello.txt missing");
            assert_eq!(contents, b"hello");
        }
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

#[test]
fn preserves_extra_ole_metadata_streams_for_standard_fixture_roundtrip() {
    let password = "password";
    let encrypted_bytes =
        std::fs::read(ooxml_fixture_path("standard.xlsx")).expect("read standard.xlsx fixture");

    // Add a dummy metadata stream to the Standard-encrypted OLE wrapper. We want to ensure the
    // `open_workbook_with_password_and_preserved_ole` path (used for encrypted round-trip) supports
    // Standard/CryptoAPI decryption and preserves non-encryption OLE streams.
    let dummy_bytes = b"dummy summary information bytes (standard)";
    let cursor = Cursor::new(encrypted_bytes);
    let mut ole = cfb::CompoundFile::open(cursor).expect("open standard encrypted cfb");
    ole.create_stream(SUMMARY_INFORMATION)
        .expect("create SummaryInformation stream")
        .write_all(dummy_bytes)
        .expect("write SummaryInformation bytes");
    let encrypted_with_extra = ole.into_inner().into_inner();

    let tmp = tempfile::tempdir().expect("tempdir");
    let in_path = tmp.path().join("input-standard.xlsx");
    std::fs::write(&in_path, &encrypted_with_extra).expect("write encrypted input");

    let opened = open_workbook_with_password_and_preserved_ole(&in_path, Some(password))
        .expect("open standard encrypted workbook with preservation");
    assert!(
        opened.preserved_ole.is_some(),
        "expected preserved OLE entries to be captured"
    );

    let out_path = tmp.path().join("output.xlsx");
    opened
        .save_preserving_encryption(&out_path, password)
        .expect("save with encryption preserved");

    // Verify the extra stream is preserved byte-for-byte in the output OLE container.
    let out_bytes = std::fs::read(&out_path).expect("read output bytes");
    let mut out_ole = cfb::CompoundFile::open(Cursor::new(out_bytes)).expect("open output cfb");
    let mut stream = out_ole
        .open_stream(SUMMARY_INFORMATION)
        .expect("open SummaryInformation stream");
    let mut got = Vec::new();
    stream.read_to_end(&mut got).expect("read SummaryInformation");
    assert_eq!(got, dummy_bytes);

    // Sanity: ensure the re-encrypted workbook can be reopened and contains a valid XLSX package.
    let reopened =
        open_workbook_with_password(&out_path, Some(password)).expect("reopen encrypted output");
    match reopened {
        Workbook::Xlsx(package) => {
            let workbook_xml = package
                .read_part("xl/workbook.xml")
                .expect("read xl/workbook.xml")
                .expect("xl/workbook.xml missing");
            let s = std::str::from_utf8(&workbook_xml).expect("workbook.xml must be UTF-8");
            assert!(
                s.contains("Sheet1"),
                "expected workbook.xml to contain Sheet1, got:\n{s}"
            );
        }
        other => panic!("expected Workbook::Xlsx, got {other:?}"),
    }
}

#[test]
fn preserved_ole_open_reports_malformed_encryptedpackage_as_unsupported() {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole
            .create_stream("EncryptionInfo")
            .expect("create EncryptionInfo stream");
        // Minimal Agile (4.4) EncryptionInfo header.
        stream
            .write_all(&[4, 0, 4, 0, 0, 0, 0, 0])
            .expect("write EncryptionInfo header");
    }
    {
        // Intentionally create an empty/truncated EncryptedPackage stream.
        ole.create_stream("EncryptedPackage")
            .expect("create EncryptedPackage stream");
    }

    let bytes = ole.into_inner().into_inner();
    let tmp = tempfile::tempdir().expect("tempdir");
    let path = tmp.path().join("malformed.xlsx");
    std::fs::write(&path, bytes).expect("write fixture");

    let err = open_workbook_with_password_and_preserved_ole(&path, Some("wrong"))
        .expect_err("expected malformed encrypted container to error");
    assert!(
        matches!(err, Error::UnsupportedOoxmlEncryption { .. }),
        "expected Error::UnsupportedOoxmlEncryption, got {err:?}"
    );
}
