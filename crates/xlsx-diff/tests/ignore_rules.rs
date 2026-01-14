use std::collections::BTreeSet;
use std::io::{Cursor, Write};
use std::process::Command;

use xlsx_diff::{
    diff_archives_with_options, DiffOptions, IgnorePathRule, Severity, WorkbookArchive,
};
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
fn extra_calcchain_xml_is_warning_by_default_and_critical_in_strict_mode() {
    let expected_zip = zip_bytes(&[]);
    let actual_zip = zip_bytes(&[("xl/calcChain.xml", br#"<calcChain/>"#)]);
    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert_eq!(report.differences.len(), 1);
    assert_eq!(report.differences[0].kind, "extra_part");
    assert_eq!(report.differences[0].part, "xl/calcChain.xml");
    assert_eq!(report.differences[0].severity, Severity::Warning);

    let options = DiffOptions {
        strict_calc_chain: true,
        ..Default::default()
    };
    let report = diff_archives_with_options(&expected, &actual, &options);
    assert_eq!(report.differences.len(), 1);
    assert_eq!(report.differences[0].kind, "extra_part");
    assert_eq!(report.differences[0].part, "xl/calcChain.xml");
    assert_eq!(report.differences[0].severity, Severity::Critical);
}

#[test]
fn calcchain_xml_content_diffs_are_warning_by_default_and_critical_in_strict_mode() {
    let expected_zip = zip_bytes(&[(
        "xl/calcChain.xml",
        br#"<calcChain><c r="A1"/></calcChain>"#,
    )]);
    let actual_zip = zip_bytes(&[(
        "xl/calcChain.xml",
        br#"<calcChain><c r="B2"/></calcChain>"#,
    )]);
    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert!(
        report.differences.iter().any(|d| {
            d.part == "xl/calcChain.xml"
                && d.kind != "missing_part"
                && d.kind != "extra_part"
                && d.severity == Severity::Warning
        }),
        "expected calcChain.xml content diffs to be WARNING by default, got {:#?}",
        report.differences
    );
    assert_eq!(
        report.count(Severity::Critical),
        0,
        "expected no CRITICAL diffs by default, got {:#?}",
        report.differences
    );

    let options = DiffOptions {
        strict_calc_chain: true,
        ..Default::default()
    };
    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report.differences.iter().any(|d| {
            d.part == "xl/calcChain.xml"
                && d.kind != "missing_part"
                && d.kind != "extra_part"
                && d.severity == Severity::Critical
        }),
        "expected calcChain.xml content diffs to be CRITICAL in strict mode, got {:#?}",
        report.differences
    );
}

#[test]
fn calcchain_bin_binary_diffs_are_warning_by_default_and_critical_in_strict_mode() {
    let expected_zip = zip_bytes(&[("xl/calcChain.bin", &[0x01, 0x02, 0x03])]);
    let actual_zip = zip_bytes(&[("xl/calcChain.bin", &[0x01, 0x02, 0x04])]);
    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert_eq!(report.differences.len(), 1);
    assert_eq!(report.differences[0].kind, "binary_diff");
    assert_eq!(report.differences[0].part, "xl/calcChain.bin");
    assert_eq!(report.differences[0].severity, Severity::Warning);

    let options = DiffOptions {
        strict_calc_chain: true,
        ..Default::default()
    };
    let report = diff_archives_with_options(&expected, &actual, &options);
    assert_eq!(report.differences.len(), 1);
    assert_eq!(report.differences[0].kind, "binary_diff");
    assert_eq!(report.differences[0].part, "xl/calcChain.bin");
    assert_eq!(report.differences[0].severity, Severity::Critical);
}

#[test]
fn missing_calcchain_xml_is_critical_when_strict_enabled() {
    let expected_zip = zip_bytes(&[("xl/calcChain.xml", br#"<calcChain/>"#)]);
    let actual_zip = zip_bytes(&[]);
    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let options = DiffOptions {
        strict_calc_chain: true,
        ..Default::default()
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert_eq!(report.differences.len(), 1);
    assert_eq!(report.differences[0].kind, "missing_part");
    assert_eq!(report.differences[0].part, "xl/calcChain.xml");
    assert_eq!(report.differences[0].severity, Severity::Critical);
}

#[test]
fn missing_calcchain_bin_is_critical_when_strict_enabled() {
    let expected_zip = zip_bytes(&[("xl/calcChain.bin", &[0x01, 0x02, 0x03])]);
    let actual_zip = zip_bytes(&[]);
    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let options = DiffOptions {
        strict_calc_chain: true,
        ..Default::default()
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert_eq!(report.differences.len(), 1);
    assert_eq!(report.differences[0].kind, "missing_part");
    assert_eq!(report.differences[0].part, "xl/calcChain.bin");
    assert_eq!(report.differences[0].severity, Severity::Critical);
}

#[test]
fn calcchain_related_rels_and_content_types_remain_critical_when_strict_enabled() {
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
        ("xl/calcChain.xml", br#"<calcChain/>"#),
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

    let options = DiffOptions {
        strict_calc_chain: true,
        ..Default::default()
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "xl/calcChain.xml"
                && d.kind == "missing_part"
                && d.severity == Severity::Critical),
        "expected missing xl/calcChain.xml to be CRITICAL, got {:#?}",
        report.differences
    );
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "[Content_Types].xml" && d.severity == Severity::Critical),
        "expected [Content_Types].xml calcChain diffs to remain CRITICAL, got {:#?}",
        report.differences
    );
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "xl/_rels/workbook.xml.rels" && d.severity == Severity::Critical),
        "expected workbook.xml.rels calcChain diffs to remain CRITICAL, got {:#?}",
        report.differences
    );
}

#[test]
fn calcchain_relationship_id_renumbering_is_warning_by_default_and_critical_in_strict_mode() {
    let expected_zip = zip_bytes(&[(
        "xl/_rels/workbook.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>"#,
    )]);
    let actual_zip = zip_bytes(&[(
        "xl/_rels/workbook.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>"#,
    )]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    // Default behavior: calcChain-related relationship churn is warning-level noise.
    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert_eq!(
        report.count(Severity::Critical),
        0,
        "expected no CRITICAL diffs, got {:#?}",
        report.differences
    );
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "xl/_rels/workbook.xml.rels"
                && d.kind == "relationship_id_changed"
                && d.severity == Severity::Warning),
        "expected calcChain relationship_id_changed diff to be WARNING by default, got {:#?}",
        report.differences
    );

    // Strict mode: keep calcChain diffs CRITICAL.
    let options = DiffOptions {
        strict_calc_chain: true,
        ..Default::default()
    };
    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "xl/_rels/workbook.xml.rels"
                && d.kind == "relationship_id_changed"
                && d.severity == Severity::Critical),
        "expected calcChain relationship_id_changed diff to be CRITICAL in strict mode, got {:#?}",
        report.differences
    );
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
    assert_eq!(
        report.count(Severity::Critical),
        0,
        "{:#?}",
        report.differences
    );
    assert!(
        report.differences.iter().any(|d| d.kind == "missing_part"
            && d.part == "xl/calcChain.xml"
            && d.severity == Severity::Warning),
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
fn calcchain_bin_related_rels_and_content_types_downgrade_to_warning() {
    let expected_zip = zip_bytes(&[
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="bin" ContentType="application/octet-stream"/>
  <Override PartName="/xl/calcChain.bin" ContentType="application/vnd.ms-excel.calcChain"/>
</Types>"#,
        ),
        (
            "xl/_rels/workbook.bin.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.bin"/>
</Relationships>"#,
        ),
        ("xl/calcChain.bin", b"dummy"),
    ]);

    let actual_zip = zip_bytes(&[
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="bin" ContentType="application/octet-stream"/>
</Types>"#,
        ),
        (
            "xl/_rels/workbook.bin.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
        ),
    ]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert_eq!(
        report.count(Severity::Critical),
        0,
        "{:#?}",
        report.differences
    );
    assert!(
        report.differences.iter().any(|d| d.kind == "missing_part"
            && d.part == "xl/calcChain.bin"
            && d.severity == Severity::Warning),
        "expected calcChain.bin missing_part to be a warning, got {:#?}",
        report.differences
    );
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "[Content_Types].xml" && d.severity == Severity::Warning),
        "expected content types calcChain.bin diff to be warning, got {:#?}",
        report.differences
    );
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part.ends_with(".rels") && d.severity == Severity::Warning),
        "expected rels calcChain.bin diff to be warning, got {:#?}",
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
fn missing_and_extra_part_severity_is_part_aware() {
    let content_types = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#;
    let rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
    let doc_props = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>"#;

    // Missing parts: severity should match part importance.
    let expected_zip = zip_bytes(&[
        ("[Content_Types].xml", content_types),
        ("xl/_rels/workbook.xml.rels", rels),
        ("docProps/app.xml", doc_props),
    ]);
    let actual_zip = zip_bytes(&[]);
    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert!(
        report.differences.iter().any(|d| d.kind == "missing_part"
            && d.part == "[Content_Types].xml"
            && d.severity == Severity::Critical),
        "expected missing [Content_Types].xml to be CRITICAL, got {:#?}",
        report.differences
    );
    assert!(
        report.differences.iter().any(|d| d.kind == "missing_part"
            && d.part.ends_with(".rels")
            && d.severity == Severity::Critical),
        "expected missing *.rels to be CRITICAL, got {:#?}",
        report.differences
    );
    assert!(
        report.differences.iter().any(|d| d.kind == "missing_part"
            && d.part == "docProps/app.xml"
            && d.severity == Severity::Info),
        "expected missing docProps/* to be INFO, got {:#?}",
        report.differences
    );

    // Extra parts: `[Content_Types].xml` and `.rels` should remain CRITICAL, `docProps/*` INFO.
    let expected_zip = zip_bytes(&[]);
    let actual_zip = zip_bytes(&[
        ("[Content_Types].xml", content_types),
        ("xl/_rels/workbook.xml.rels", rels),
        ("docProps/app.xml", doc_props),
    ]);
    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert!(
        report.differences.iter().any(|d| d.kind == "extra_part"
            && d.part == "[Content_Types].xml"
            && d.severity == Severity::Critical),
        "expected extra [Content_Types].xml to be CRITICAL, got {:#?}",
        report.differences
    );
    assert!(
        report.differences.iter().any(|d| d.kind == "extra_part"
            && d.part.ends_with(".rels")
            && d.severity == Severity::Critical),
        "expected extra *.rels to be CRITICAL, got {:#?}",
        report.differences
    );
    assert!(
        report.differences.iter().any(|d| d.kind == "extra_part"
            && d.part == "docProps/app.xml"
            && d.severity == Severity::Info),
        "expected extra docProps/* to be INFO, got {:#?}",
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
        ignore_paths: Vec::new(),
        strict_calc_chain: false,
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report.is_empty(),
        "expected no diffs, got {:#?}",
        report.differences
    );
}

#[test]
fn ignore_rules_normalize_leading_slashes_and_backslashes() {
    let expected_zip = zip_bytes(&[("xl/calcChain.xml", br#"<calcChain/>"#)]);
    let actual_zip = zip_bytes(&[]);
    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let mut ignore_parts = BTreeSet::new();
    ignore_parts.insert(r"/xl\calcChain.xml".to_string());

    let options = DiffOptions {
        ignore_parts,
        ignore_globs: vec![r"\xl\calcChain.*".to_string()],
        ignore_paths: Vec::new(),
        strict_calc_chain: false,
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report.is_empty(),
        "expected no diffs, got {:#?}",
        report.differences
    );
}

#[test]
fn zip_entry_names_are_normalized_for_diffing() {
    let expected_zip = zip_bytes(&[(r"xl\\calcChain.xml", br#"<calcChain/>"#)]);
    let actual_zip = zip_bytes(&[("xl/calcChain.xml", br#"<calcChain/>"#)]);
    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert!(
        report.is_empty(),
        "expected no diffs, got {:#?}",
        report.differences
    );
}

#[test]
fn zip_entry_names_with_leading_slash_are_normalized_for_diffing() {
    let payload = b"workbook-bytes";
    let expected_zip = zip_bytes(&[("xl/workbook.bin", payload)]);
    let actual_zip = zip_bytes(&[("/xl/workbook.bin", payload)]);
    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    assert_eq!(actual.get("xl/workbook.bin"), Some(payload.as_slice()));

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert!(
        report.is_empty(),
        "expected no diffs when entry names only differ by leading '/'; got {:#?}",
        report.differences
    );
}

#[test]
fn workbook_archive_errors_on_duplicate_parts_after_normalization() {
    let zip = zip_bytes(&[
        ("xl/calcChain.xml", br#"<calcChain/>"#),
        (r"xl\\calcChain.xml", br#"<calcChain/>"#),
    ]);

    match WorkbookArchive::from_bytes(&zip) {
        Ok(_) => panic!("expected WorkbookArchive::from_bytes to fail due to duplicate parts"),
        Err(err) => assert!(
            err.to_string()
                .contains("duplicate part name after normalization"),
            "unexpected error: {err}"
        ),
    }
}

#[test]
fn ignore_glob_suppresses_calcchain_related_plumbing_diffs() {
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
        ("xl/calcChain.xml", br#"<calcChain/>"#),
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

    let options = DiffOptions {
        ignore_parts: Default::default(),
        ignore_globs: vec!["xl/calcChain.*".to_string()],
        ignore_paths: Vec::new(),
        strict_calc_chain: false,
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report.is_empty(),
        "expected no diffs, got {:#?}",
        report.differences
    );
}

#[test]
fn ignore_glob_suppresses_docprops_plumbing_diffs() {
    let expected_zip = zip_bytes(&[
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"#,
        ),
        (
            "_rels/.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"#,
        ),
        (
            "docProps/app.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"/>"#,
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
            "_rels/.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
        ),
    ]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let options = DiffOptions {
        ignore_parts: Default::default(),
        ignore_globs: vec!["docProps/*".to_string()],
        ignore_paths: Vec::new(),
        strict_calc_chain: false,
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report.is_empty(),
        "expected no diffs, got {:#?}",
        report.differences
    );
}

#[test]
fn ignore_rules_do_not_hide_relationship_target_changes_to_non_ignored_parts() {
    let expected_zip = zip_bytes(&[
        (
            "[Content_Types].xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"#,
        ),
        (
            "_rels/.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"#,
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
            "_rels/.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="customXml/item1.xml"/>
</Relationships>"#,
        ),
    ]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let options = DiffOptions {
        ignore_parts: Default::default(),
        ignore_globs: vec!["docProps/*".to_string()],
        ignore_paths: Vec::new(),
        strict_calc_chain: false,
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "_rels/.rels" && d.kind == "attribute_changed"),
        "expected _rels/.rels diff to remain (target changed to non-ignored part), got {:#?}",
        report.differences
    );
}

#[test]
fn ignore_glob_suppresses_rels_targets_with_parent_dir_segments() {
    let expected_zip = zip_bytes(&[
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#,
        ),
        ("xl/media/image1.png", b"pngbytes"),
    ]);

    let actual_zip = zip_bytes(&[(
        "xl/drawings/_rels/drawing1.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
    )]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let options = DiffOptions {
        ignore_parts: Default::default(),
        ignore_globs: vec!["xl/media/*".to_string()],
        ignore_paths: Vec::new(),
        strict_calc_chain: false,
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report.is_empty(),
        "expected no diffs, got {:#?}",
        report.differences
    );
}

#[test]
fn ignore_path_suppresses_specific_xml_attribute_diffs() {
    let expected_zip = zip_bytes(&[(
        "xl/worksheets/sheet1.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
    )]);
    let actual_zip = zip_bytes(&[(
        "xl/worksheets/sheet1.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30"/>
</worksheet>"#,
    )]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    // Sanity: without the ignore-path rule, we should see an attribute diff.
    let base_report = xlsx_diff::diff_archives(&expected, &actual);
    assert!(
        base_report
            .differences
            .iter()
            .any(|d| d.part == "xl/worksheets/sheet1.xml"
                && d.kind == "attribute_changed"
                && d.path.contains("dyDescent")),
        "expected a dyDescent attribute diff, got {:#?}",
        base_report.differences
    );

    let options = DiffOptions {
        ignore_parts: Default::default(),
        ignore_globs: Vec::new(),
        ignore_paths: vec![IgnorePathRule {
            part: None,
            path_substring: "dyDescent".to_string(),
            kind: None,
        }],
        strict_calc_chain: false,
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report.is_empty(),
        "expected no diffs, got {:#?}",
        report.differences
    );
}

#[test]
fn ignore_path_scoped_to_part_only_suppresses_matching_part() {
    let expected_zip = zip_bytes(&[
        (
            "xl/worksheets/sheet1.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
        ),
        (
            "xl/worksheets/sheet2.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
        ),
    ]);
    let actual_zip = zip_bytes(&[
        (
            "xl/worksheets/sheet1.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30"/>
</worksheet>"#,
        ),
        (
            "xl/worksheets/sheet2.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30"/>
</worksheet>"#,
        ),
    ]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    // Sanity: without any ignore rules, we should see diffs for both parts.
    let base_report = xlsx_diff::diff_archives(&expected, &actual);
    assert!(
        base_report
            .differences
            .iter()
            .any(|d| d.part == "xl/worksheets/sheet1.xml"),
        "expected sheet1.xml diff, got {:#?}",
        base_report.differences
    );
    assert!(
        base_report
            .differences
            .iter()
            .any(|d| d.part == "xl/worksheets/sheet2.xml"),
        "expected sheet2.xml diff, got {:#?}",
        base_report.differences
    );

    let options = DiffOptions {
        ignore_parts: Default::default(),
        ignore_globs: Vec::new(),
        ignore_paths: vec![IgnorePathRule {
            part: Some("xl/worksheets/sheet1.xml".to_string()),
            path_substring: "dyDescent".to_string(),
            kind: None,
        }],
        strict_calc_chain: false,
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report
            .differences
            .iter()
            .all(|d| d.part == "xl/worksheets/sheet2.xml"),
        "expected only sheet2.xml diffs to remain, got {:#?}",
        report.differences
    );
}

#[test]
fn ignore_path_kind_filter_only_suppresses_matching_kind() {
    let expected_zip = zip_bytes(&[(
        "xl/worksheets/sheet1.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
    )]);
    let actual_zip = zip_bytes(&[(
        "xl/worksheets/sheet1.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30" foo="bar"/>
</worksheet>"#,
    )]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    // Only suppress the dyDescent attribute_changed diff; leave other diff kinds intact.
    let options = DiffOptions {
        ignore_parts: Default::default(),
        ignore_globs: Vec::new(),
        ignore_paths: vec![IgnorePathRule {
            part: None,
            path_substring: "dyDescent".to_string(),
            kind: Some("attribute_changed".to_string()),
        }],
        strict_calc_chain: false,
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert_eq!(
        report.differences.len(),
        1,
        "expected only the foo attribute_added diff to remain, got {:#?}",
        report.differences
    );
    let diff = &report.differences[0];
    assert_eq!(diff.part, "xl/worksheets/sheet1.xml");
    assert_eq!(diff.kind, "attribute_added");
    assert!(
        diff.path.contains("@foo"),
        "expected diff path to include '@foo', got {}",
        diff.path
    );
}

#[test]
fn ignore_path_part_glob_suppresses_all_matching_parts() {
    let expected_zip = zip_bytes(&[
        (
            "xl/worksheets/sheet1.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
        ),
        (
            "xl/worksheets/sheet2.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
        ),
    ]);
    let actual_zip = zip_bytes(&[
        (
            "xl/worksheets/sheet1.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30"/>
</worksheet>"#,
        ),
        (
            "xl/worksheets/sheet2.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30"/>
</worksheet>"#,
        ),
    ]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let options = DiffOptions {
        ignore_parts: Default::default(),
        ignore_globs: Vec::new(),
        ignore_paths: vec![IgnorePathRule {
            part: Some("xl/worksheets/*.xml".to_string()),
            path_substring: "dyDescent".to_string(),
            kind: None,
        }],
        strict_calc_chain: false,
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report.is_empty(),
        "expected no diffs, got {:#?}",
        report.differences
    );
}

#[test]
fn ignore_path_with_invalid_part_glob_does_not_match_all_parts() {
    let expected_zip = zip_bytes(&[(
        "xl/worksheets/sheet1.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
    )]);
    let actual_zip = zip_bytes(&[(
        "xl/worksheets/sheet1.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30"/>
</worksheet>"#,
    )]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    // The part selector is an invalid glob (`[` is not closed). Invalid part globs should be
    // ignored (rule dropped), not treated as "match any part".
    let options = DiffOptions {
        ignore_parts: Default::default(),
        ignore_globs: Vec::new(),
        ignore_paths: vec![IgnorePathRule {
            part: Some("xl/worksheets/*[".to_string()),
            path_substring: "dyDescent".to_string(),
            kind: None,
        }],
        strict_calc_chain: false,
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.part == "xl/worksheets/sheet1.xml"
                && d.kind == "attribute_changed"
                && d.path.contains("dyDescent")),
        "expected dyDescent diff to remain (invalid part glob ignored), got {:#?}",
        report.differences
    );
}

#[test]
fn cli_ignore_path_flag_suppresses_specific_xml_attribute_diffs() {
    let expected_zip = zip_bytes(&[(
        "xl/worksheets/sheet1.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
    )]);
    let actual_zip = zip_bytes(&[(
        "xl/worksheets/sheet1.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30"/>
</worksheet>"#,
    )]);

    let tempdir = tempfile::tempdir().unwrap();
    let original_path = tempdir.path().join("original.xlsx");
    let modified_path = tempdir.path().join("modified.xlsx");
    std::fs::write(&original_path, expected_zip).unwrap();
    std::fs::write(&modified_path, actual_zip).unwrap();

    // Without ignores, the CLI should report a critical diff and exit non-zero.
    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&original_path)
        .arg(&modified_path)
        .output()
        .unwrap();
    assert!(
        !output.status.success(),
        "expected non-zero exit, got {:?}\nstdout:\n{}\nstderr:\n{}",
        output.status,
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&original_path)
        .arg(&modified_path)
        .arg("--ignore-path")
        .arg("dyDescent")
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "expected exit 0, got {:?}\nstdout:\n{}\nstderr:\n{}",
        output.status,
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(
        String::from_utf8_lossy(&output.stdout).contains("No differences."),
        "expected output to say 'No differences.', got:\n{}",
        String::from_utf8_lossy(&output.stdout)
    );
}

#[test]
fn cli_ignore_path_in_flag_scopes_to_matching_part_only() {
    let expected_zip = zip_bytes(&[
        (
            "xl/worksheets/sheet1.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
        ),
        (
            "xl/worksheets/sheet2.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
        ),
    ]);
    let actual_zip = zip_bytes(&[
        (
            "xl/worksheets/sheet1.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30"/>
</worksheet>"#,
        ),
        (
            "xl/worksheets/sheet2.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30"/>
</worksheet>"#,
        ),
    ]);

    let tempdir = tempfile::tempdir().unwrap();
    let original_path = tempdir.path().join("original.xlsx");
    let modified_path = tempdir.path().join("modified.xlsx");
    std::fs::write(&original_path, expected_zip).unwrap();
    std::fs::write(&modified_path, actual_zip).unwrap();

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&original_path)
        .arg(&modified_path)
        .arg("--ignore-path-in")
        .arg("xl/worksheets/sheet1.xml:dyDescent")
        .output()
        .unwrap();

    assert!(
        !output.status.success(),
        "expected non-zero exit (sheet2.xml diff remains), got {:?}\nstdout:\n{}\nstderr:\n{}",
        output.status,
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        stdout.contains("xl/worksheets/sheet2.xml"),
        "expected sheet2.xml diff to remain, got:\n{stdout}"
    );
    assert!(
        !stdout.contains("] xl/worksheets/sheet1.xml:"),
        "expected sheet1.xml diff to be suppressed (ignores may still be listed in header), got:\n{stdout}"
    );
}

#[test]
fn cli_ignore_path_in_glob_suppresses_all_matching_parts() {
    let expected_zip = zip_bytes(&[
        (
            "xl/worksheets/sheet1.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
        ),
        (
            "xl/worksheets/sheet2.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
</worksheet>"#,
        ),
    ]);
    let actual_zip = zip_bytes(&[
        (
            "xl/worksheets/sheet1.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30"/>
</worksheet>"#,
        ),
        (
            "xl/worksheets/sheet2.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.30"/>
</worksheet>"#,
        ),
    ]);

    let tempdir = tempfile::tempdir().unwrap();
    let original_path = tempdir.path().join("original.xlsx");
    let modified_path = tempdir.path().join("modified.xlsx");
    std::fs::write(&original_path, expected_zip).unwrap();
    std::fs::write(&modified_path, actual_zip).unwrap();

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&original_path)
        .arg(&modified_path)
        .arg("--ignore-path-in")
        .arg("xl/worksheets/*.xml:dyDescent")
        .output()
        .unwrap();

    assert!(
        output.status.success(),
        "expected exit 0, got {:?}\nstdout:\n{}\nstderr:\n{}",
        output.status,
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(
        String::from_utf8_lossy(&output.stdout).contains("No differences."),
        "expected output to say 'No differences.', got:\n{}",
        String::from_utf8_lossy(&output.stdout)
    );
}

#[test]
fn ignore_glob_suppresses_rels_targets_with_xl_prefix_without_leading_slash() {
    let expected_zip = zip_bytes(&[
        (
            "xl/_rels/workbook.xml.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="xl/media/image1.png"/>
</Relationships>"#,
        ),
        ("xl/media/image1.png", b"pngbytes"),
    ]);

    let actual_zip = zip_bytes(&[
        (
            "xl/_rels/workbook.xml.rels",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#,
        ),
        ("xl/media/image1.png", b"pngbytes"),
    ]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let options = DiffOptions {
        ignore_globs: vec!["xl/media/*".to_string()],
        ..Default::default()
    };

    let report = diff_archives_with_options(&expected, &actual, &options);
    assert!(
        report.is_empty(),
        "expected no diffs, got {:#?}",
        report.differences
    );
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

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&original_path)
        .arg(&modified_path)
        .arg("--fail-on")
        .arg("critical")
        .output()
        .unwrap();
    assert!(
        output.status.success(),
        "expected exit 0, got {:?}\nstdout:\n{}\nstderr:\n{}",
        output.status,
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&original_path)
        .arg(&modified_path)
        .arg("--strict-calc-chain")
        .arg("--fail-on")
        .arg("critical")
        .output()
        .unwrap();
    assert!(
        !output.status.success(),
        "expected non-zero exit with --strict-calc-chain, got {:?}\nstdout:\n{}\nstderr:\n{}",
        output.status,
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&original_path)
        .arg(&modified_path)
        .arg("--fail-on")
        .arg("warning")
        .output()
        .unwrap();
    assert!(
        !output.status.success(),
        "expected non-zero exit, got {:?}\nstdout:\n{}\nstderr:\n{}",
        output.status,
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
}

#[test]
fn cli_rejects_invalid_ignore_glob_pattern() {
    let empty_zip = zip_bytes(&[]);
    let tempdir = tempfile::tempdir().unwrap();
    let original_path = tempdir.path().join("original.xlsx");
    let modified_path = tempdir.path().join("modified.xlsx");
    std::fs::write(&original_path, &empty_zip).unwrap();
    std::fs::write(&modified_path, &empty_zip).unwrap();

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&original_path)
        .arg(&modified_path)
        .arg("--ignore-glob")
        .arg("[")
        .output()
        .unwrap();

    assert!(
        !output.status.success(),
        "expected non-zero exit, got {:?}\nstdout:\n{}\nstderr:\n{}",
        output.status,
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
}
