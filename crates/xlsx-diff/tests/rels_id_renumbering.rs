use std::io::{Cursor, Write};

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
