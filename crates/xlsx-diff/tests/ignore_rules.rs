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
fn calcchain_related_rels_and_content_types_downgrade_to_warning() {
    let expected_zip = zip_bytes(&[
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>"#,
        ),
        (
            "xl/calcChain.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
        ),
    ]);

    let actual_zip = zip_bytes(&[
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
        ),
    ]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert_eq!(report.count(Severity::Critical), 0, "{:#?}", report.differences);
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.kind == "missing_part" && d.part == "xl/calcChain.xml" && d.severity == Severity::Warning),
        "expected calcChain missing_part to be a warning, got {:#?}",
        report.differences
    );
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "[Content_Types].xml" && d.severity == Severity::Warning),
        "expected content types calcChain diff to be warning, got {:#?}",
        report.differences
    );
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part.ends_with(".rels") && d.severity == Severity::Warning),
        "expected rels calcChain diff to be warning, got {:#?}",
        report.differences
    );
}

#[test]
fn non_calcchain_rels_diffs_remain_critical() {
    let expected_zip = zip_bytes(&[(
        "xl/_rels/workbook.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
    )]);
    let actual_zip = zip_bytes(&[(
        "xl/_rels/workbook.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>"#,
    )]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert!(
        report.has_at_least(Severity::Critical),
        "expected critical diffs, got {:#?}",
        report.differences
    );
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
    let expected_zip = zip_bytes(&[
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
</Types>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>"#,
        ),
        (
            "xl/calcChain.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#,
        ),
    ]);

    let actual_zip = zip_bytes(&[
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>"#,
        ),
        (
            "xl/_rels/workbook.xml.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
        ),
    ]);

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
