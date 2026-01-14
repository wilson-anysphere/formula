use std::io::{Cursor, Write};
use std::path::PathBuf;

use formula_io::{detect_workbook_format, Error, WorkbookFormat};

fn fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../fixtures").join(rel)
}

fn xlsb_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xlsb/tests/fixtures")
        .join(rel)
}

fn xls_fixture_path(rel: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../formula-xls/tests/fixtures")
        .join(rel)
}

fn encrypted_ooxml_bytes_with_stream_names(encryption_info: &str, encrypted_package: &str) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut ole = cfb::CompoundFile::create(cursor).expect("create cfb");
    {
        let mut stream = ole
            .create_stream(encryption_info)
            .unwrap_or_else(|_| panic!("create {encryption_info} stream"));
        // Minimal EncryptionInfo header for Agile encryption (4.4).
        stream
            .write_all(&[4, 0, 4, 0, 0, 0, 0, 0])
            .expect("write EncryptionInfo header");
    }
    ole.create_stream(encrypted_package)
        .unwrap_or_else(|_| panic!("create {encrypted_package} stream"));
    ole.into_inner().into_inner()
}

#[test]
fn detects_xlsx() {
    let path = fixture_path("xlsx/basic/basic.xlsx");
    assert_eq!(
        detect_workbook_format(&path).expect("detect"),
        WorkbookFormat::Xlsx
    );
}

#[test]
fn detects_xlsm() {
    let path = fixture_path("xlsx/macros/basic.xlsm");
    assert_eq!(
        detect_workbook_format(&path).expect("detect"),
        WorkbookFormat::Xlsm
    );
}

#[test]
fn detects_xltx_as_xlsx() {
    let src = fixture_path("xlsx/basic/basic.xlsx");
    let tmp = tempfile::tempdir().expect("temp dir");
    let dst = tmp.path().join("basic.xltx");
    std::fs::copy(&src, &dst).expect("copy fixture to .xltx");

    // Template files are still OOXML spreadsheets; format detection should classify by content,
    // not extension.
    assert_eq!(
        detect_workbook_format(&dst).expect("detect"),
        WorkbookFormat::Xlsx
    );
}

#[test]
fn detects_xltm_and_xlam_as_xlsm() {
    let src = fixture_path("xlsx/macros/basic.xlsm");
    let tmp = tempfile::tempdir().expect("temp dir");

    for ext in ["xltm", "xlam"] {
        let dst = tmp.path().join(format!("basic.{ext}"));
        std::fs::copy(&src, &dst).expect("copy macro fixture");

        // Macro-enabled templates/add-ins are still OOXML ZIP packages; the presence of
        // `xl/vbaProject.bin` should classify them as "macro-enabled" regardless of extension.
        assert_eq!(
            detect_workbook_format(&dst).expect("detect"),
            WorkbookFormat::Xlsm
        );
    }
}

#[test]
fn detects_xlsb() {
    let path = xlsb_fixture_path("simple.xlsb");
    assert_eq!(
        detect_workbook_format(&path).expect("detect"),
        WorkbookFormat::Xlsb
    );
}

#[test]
fn detects_xls() {
    let path = xls_fixture_path("basic.xls");
    assert_eq!(
        detect_workbook_format(&path).expect("detect"),
        WorkbookFormat::Xls
    );
}

#[test]
fn detects_xlt_and_xla_as_xls() {
    let src = xls_fixture_path("basic.xls");
    let tmp = tempfile::tempdir().expect("temp dir");

    for ext in ["xlt", "xla"] {
        let dst = tmp.path().join(format!("basic.{ext}"));
        std::fs::copy(&src, &dst).expect("copy .xls fixture to legacy template/add-in extension");

        // `.xlt` and `.xla` are legacy BIFF/OLE containers; format detection should classify by
        // content, not extension.
        assert_eq!(
            detect_workbook_format(&dst).expect("detect"),
            WorkbookFormat::Xls
        );
    }
}

#[test]
fn detects_csv_by_extension() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.csv");
    std::fs::write(&path, "col1,col2\n1,hello\n").expect("write csv");
    assert_eq!(
        detect_workbook_format(&path).expect("detect"),
        WorkbookFormat::Csv
    );
}

#[test]
fn detects_parquet_by_magic() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.parquet");
    std::fs::write(&path, b"PAR1\x00\x00\x00\x00").expect("write parquet header");
    assert_eq!(
        detect_workbook_format(&path).expect("detect"),
        WorkbookFormat::Parquet
    );
}

#[test]
fn detects_unknown_format() {
    let dir = tempfile::tempdir().expect("temp dir");
    let path = dir.path().join("data.bin");
    std::fs::write(&path, b"not a workbook").expect("write file");
    assert_eq!(
        detect_workbook_format(&path).expect("detect"),
        WorkbookFormat::Unknown
    );
}

#[test]
fn detects_encrypted_ooxml_container() {
    let tmp = tempfile::tempdir().expect("tempdir");
    for (info, package) in [
        ("EncryptionInfo", "EncryptedPackage"),
        ("encryptioninfo", "encryptedpackage"),
        ("/encryptioninfo", "/encryptedpackage"),
    ] {
        let path = tmp.path().join("encrypted.xlsx");
        std::fs::write(&path, encrypted_ooxml_bytes_with_stream_names(info, package))
            .expect("write encrypted fixture");

        let err = detect_workbook_format(&path).expect_err("expected encrypted workbook to error");
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
    }
}

#[test]
fn detects_real_encrypted_ooxml_fixture() {
    let path = fixture_path("encryption/encrypted_agile.xlsx");
    let err = detect_workbook_format(&path).expect_err("expected encrypted workbook to error");
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
}

