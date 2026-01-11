use std::io::{Cursor, Write};
use std::process::Command;

use xlsx_diff::{diff_archives_with_options, DiffOptions, Severity, WorkbookArchive};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn zip_bytes(parts: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

    for (name, bytes) in parts {
        writer.start_file(*name, options).unwrap();
        writer.write_all(bytes).unwrap();
    }

    writer.finish().unwrap().into_inner()
}

#[test]
fn missing_calcchain_xml_is_warning() {
    let expected_zip = zip_bytes(&[("xl/calcChain.xml", br#"<calcChain/>"#)]);
    let actual_zip = zip_bytes(&[]);
    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert_eq!(report.differences.len(), 1);
    assert_eq!(report.differences[0].kind, "missing_part");
    assert_eq!(report.differences[0].part, "xl/calcChain.xml");
    assert_eq!(report.differences[0].severity, Severity::Warning);
}

#[test]
fn missing_calcchain_bin_is_warning() {
    let expected_zip = zip_bytes(&[("xl/calcChain.bin", &[0x01, 0x02, 0x03])]);
    let actual_zip = zip_bytes(&[]);
    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert_eq!(report.differences.len(), 1);
    assert_eq!(report.differences[0].kind, "missing_part");
    assert_eq!(report.differences[0].part, "xl/calcChain.bin");
    assert_eq!(report.differences[0].severity, Severity::Warning);
}

#[test]
fn ignore_glob_suppresses_calcchain_diffs() {
    let expected_zip = zip_bytes(&[("xl/calcChain.xml", br#"<calcChain/>"#)]);
    let actual_zip = zip_bytes(&[]);
    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let options = DiffOptions {
        ignore_parts: Default::default(),
        ignore_globs: vec!["xl/calcChain.*".to_string()],
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(report.is_empty(), "expected no diffs, got {:#?}", report.differences);
}

#[test]
fn cli_exit_code_honors_fail_on_with_only_warning_diffs() {
    let expected_zip = zip_bytes(&[("xl/calcChain.xml", br#"<calcChain/>"#)]);
    let actual_zip = zip_bytes(&[]);

    let tempdir = tempfile::tempdir().unwrap();
    let original_path = tempdir.path().join("original.xlsx");
    let modified_path = tempdir.path().join("modified.xlsx");
    std::fs::write(&original_path, expected_zip).unwrap();
    std::fs::write(&modified_path, actual_zip).unwrap();

    let status = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&original_path)
        .arg(&modified_path)
        .arg("--fail-on")
        .arg("critical")
        .status()
        .unwrap();
    assert!(status.success(), "expected exit 0, got {status:?}");

    let status = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&original_path)
        .arg(&modified_path)
        .arg("--fail-on")
        .arg("warning")
        .status()
        .unwrap();
    assert!(!status.success(), "expected non-zero exit, got {status:?}");
}

