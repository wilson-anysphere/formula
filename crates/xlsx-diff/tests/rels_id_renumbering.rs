use std::io::{Cursor, Write};
use std::process::Command;

use xlsx_diff::{Severity, WorkbookArchive};
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

fn utf16le_with_bom(text: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(2 + text.len() * 2);
    out.extend_from_slice(&[0xFF, 0xFE]);
    for unit in text.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn utf16be_with_bom(text: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(2 + text.len() * 2);
    out.extend_from_slice(&[0xFE, 0xFF]);
    for unit in text.encode_utf16() {
        out.extend_from_slice(&unit.to_be_bytes());
    }
    out
}

#[test]
fn rels_id_renumbering_is_reported_as_single_actionable_diff() {
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
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
    )]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert_eq!(
        report.differences.len(),
        1,
        "expected exactly one synthesized diff, got {:#?}",
        report.differences
    );

    let diff = &report.differences[0];
    assert_eq!(diff.part, "xl/_rels/workbook.xml.rels");
    assert_eq!(diff.kind, "relationship_id_changed");
    assert_eq!(diff.severity, Severity::Critical);
    assert_eq!(diff.expected.as_deref(), Some("rId1"));
    assert_eq!(diff.actual.as_deref(), Some("rId5"));
    assert!(
        diff.path.contains("relationships/worksheet"),
        "expected path to include relationship type, got {}",
        diff.path
    );
    assert!(
        diff.path.contains("xl/worksheets/sheet1.xml"),
        "expected path to include resolved target, got {}",
        diff.path
    );
}

#[test]
fn rels_id_renumbering_works_for_utf16le_rels() {
    let expected_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;
    let actual_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let expected_rels = utf16le_with_bom(expected_xml);
    let actual_rels = utf16le_with_bom(actual_xml);

    let expected_zip = zip_bytes(&[("xl/_rels/workbook.xml.rels", expected_rels.as_slice())]);
    let actual_zip = zip_bytes(&[("xl/_rels/workbook.xml.rels", actual_rels.as_slice())]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert_eq!(
        report.differences.len(),
        1,
        "expected exactly one synthesized diff, got {:#?}",
        report.differences
    );

    let diff = &report.differences[0];
    assert_eq!(diff.kind, "relationship_id_changed");
    assert_eq!(diff.expected.as_deref(), Some("rId1"));
    assert_eq!(diff.actual.as_deref(), Some("rId5"));
}

#[test]
fn rels_id_renumbering_works_for_utf16be_rels() {
    let expected_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;
    let actual_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let expected_rels = utf16be_with_bom(expected_xml);
    let actual_rels = utf16be_with_bom(actual_xml);

    let expected_zip = zip_bytes(&[("xl/_rels/workbook.xml.rels", expected_rels.as_slice())]);
    let actual_zip = zip_bytes(&[("xl/_rels/workbook.xml.rels", actual_rels.as_slice())]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert_eq!(
        report.differences.len(),
        1,
        "expected exactly one synthesized diff, got {:#?}",
        report.differences
    );

    let diff = &report.differences[0];
    assert_eq!(diff.kind, "relationship_id_changed");
    assert_eq!(diff.expected.as_deref(), Some("rId1"));
    assert_eq!(diff.actual.as_deref(), Some("rId5"));
}

#[test]
fn rels_target_changes_continue_to_surface_as_attribute_diffs() {
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
        report
            .differences
            .iter()
            .any(|d| d.part == "xl/_rels/workbook.xml.rels"
                && d.kind == "attribute_changed"
                && d.path.contains("Relationship[@Id=\"rId1\"]@Target")),
        "expected Target attribute_changed diff, got {:#?}",
        report.differences
    );
    assert!(
        !report
            .differences
            .iter()
            .any(|d| d.kind == "relationship_id_changed"),
        "unexpected relationship_id_changed diff, got {:#?}",
        report.differences
    );
}

#[test]
fn cli_output_for_rels_id_renumbering_is_concise() {
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
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
    )]);

    let tempdir = tempfile::tempdir().unwrap();
    let original_path = tempdir.path().join("original.xlsx");
    let modified_path = tempdir.path().join("modified.xlsx");
    std::fs::write(&original_path, expected_zip).unwrap();
    std::fs::write(&modified_path, actual_zip).unwrap();

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&original_path)
        .arg(&modified_path)
        .output()
        .unwrap();
    assert!(
        !output.status.success(),
        "expected non-zero exit due to CRITICAL diffs\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        stdout.contains("kind: relationship_id_changed"),
        "expected CLI output to include relationship_id_changed, got:\n{stdout}"
    );
    assert!(
        !stdout.contains("child_missing"),
        "expected no child_missing noise, got:\n{stdout}"
    );
    assert!(
        !stdout.contains("child_added"),
        "expected no child_added noise, got:\n{stdout}"
    );
    assert!(
        stdout.contains("expected: rId1") && stdout.contains("actual:   rId5"),
        "expected CLI output to include old/new Id values, got:\n{stdout}"
    );
}

#[test]
fn rels_id_permutation_is_reported_as_id_changes_not_attribute_diffs() {
    let expected_zip = zip_bytes(&[(
        "xl/_rels/workbook.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#,
    )]);

    let actual_zip = zip_bytes(&[(
        "xl/_rels/workbook.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
    )]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);

    assert_eq!(
        report.differences.len(),
        2,
        "expected two synthesized relationship_id_changed diffs, got {:#?}",
        report.differences
    );
    assert!(
        report
            .differences
            .iter()
            .all(|d| d.kind == "relationship_id_changed"),
        "expected only relationship_id_changed diffs, got {:#?}",
        report.differences
    );
    assert!(
        !report
            .differences
            .iter()
            .any(|d| d.kind == "attribute_changed"),
        "did not expect attribute diffs for pure Id permutation, got {:#?}",
        report.differences
    );
}

#[test]
fn rels_id_renumbering_is_detected_even_when_other_relationships_have_duplicate_semantics() {
    // Some producers emit multiple relationships with the same (Type, Target) semantics (e.g. images).
    // Ensure that ambiguity for those relationships doesn't disable Id-renumbering detection for
    // unrelated, uniquely-identifiable relationships.

    let expected_zip = zip_bytes(&[(
        "xl/_rels/workbook.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#,
    )]);

    let actual_zip = zip_bytes(&[(
        "xl/_rels/workbook.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#,
    )]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert_eq!(
        report.differences.len(),
        1,
        "expected exactly one relationship_id_changed diff, got {:#?}",
        report.differences
    );
    assert_eq!(report.differences[0].kind, "relationship_id_changed");
    assert_eq!(report.differences[0].expected.as_deref(), Some("rId1"));
    assert_eq!(report.differences[0].actual.as_deref(), Some("rId5"));
}

#[test]
fn rels_id_renumbering_with_removed_relationship_does_not_emit_attribute_noise_for_reused_ids() {
    // When a relationship is removed, some producers renumber the remaining relationships such
    // that an existing Id is reused. The raw XML diff (keyed by Id) would report this as
    // attribute diffs for the reused Id, even though the meaningful change is:
    // - one relationship's Id changed
    // - another relationship was removed
    //
    // Ensure we keep the synthesized `relationship_id_changed` diff, but do not surface noisy
    // `attribute_*` diffs for the reused Id.
    let expected_zip = zip_bytes(&[(
        "xl/_rels/workbook.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
    )]);

    let actual_zip = zip_bytes(&[(
        "xl/_rels/workbook.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
    )]);

    let expected = WorkbookArchive::from_bytes(&expected_zip).unwrap();
    let actual = WorkbookArchive::from_bytes(&actual_zip).unwrap();

    let report = xlsx_diff::diff_archives(&expected, &actual);
    assert!(
        report
            .differences
            .iter()
            .any(|d| d.kind == "relationship_id_changed"
                && d.expected.as_deref() == Some("rId2")
                && d.actual.as_deref() == Some("rId1")),
        "expected a relationship_id_changed diff for the worksheet relationship, got {:#?}",
        report.differences
    );
    assert!(
        report.differences.iter().any(|d| {
            d.kind == "relationship_missing"
                && d.path.contains("relationships/calcChain")
                && d.path.contains("xl/calcChain.xml")
        }),
        "expected a relationship_missing diff for the calcChain relationship, got {:#?}",
        report.differences
    );
    assert!(
        !report.differences.iter().any(|d| d.kind.starts_with("attribute_")),
        "did not expect attribute diffs for reused relationship ids, got {:#?}",
        report.differences
    );
}

#[test]
fn cli_output_for_rels_id_reuse_does_not_emit_attribute_noise() {
    let expected_zip = zip_bytes(&[(
        "xl/_rels/workbook.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
    )]);

    let actual_zip = zip_bytes(&[(
        "xl/_rels/workbook.xml.rels",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#,
    )]);

    let tempdir = tempfile::tempdir().unwrap();
    let original_path = tempdir.path().join("original.xlsx");
    let modified_path = tempdir.path().join("modified.xlsx");
    std::fs::write(&original_path, expected_zip).unwrap();
    std::fs::write(&modified_path, actual_zip).unwrap();

    let output = Command::new(env!("CARGO_BIN_EXE_xlsx_diff"))
        .arg(&original_path)
        .arg(&modified_path)
        .output()
        .unwrap();
    assert!(
        !output.status.success(),
        "expected non-zero exit due to CRITICAL diffs\nstdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );

    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(
        stdout.contains("kind: relationship_id_changed"),
        "expected CLI output to include relationship_id_changed, got:\n{stdout}"
    );
    assert!(
        !stdout.contains("child_missing") && !stdout.contains("child_added"),
        "expected no child_* relationship noise, got:\n{stdout}"
    );
    assert!(
        !stdout.contains("attribute_changed")
            && !stdout.contains("attribute_missing")
            && !stdout.contains("attribute_added"),
        "expected no attribute_* relationship noise, got:\n{stdout}"
    );
}
