use std::collections::BTreeMap;
use std::io::{Cursor, Write};

use formula_xlsx::{WorkbookKind, XlsxLazyPackage};
use roxmltree::Document;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn build_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let mut cursor = Cursor::new(Vec::new());
    {
        let mut zip = ZipWriter::new(&mut cursor);
        let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);
        for (name, bytes) in entries {
            zip.start_file(*name, options).unwrap();
            zip.write_all(bytes).unwrap();
        }
        zip.finish().unwrap();
    }
    cursor.into_inner()
}

fn content_types_override_map(xml: &str) -> BTreeMap<String, String> {
    let doc = Document::parse(xml).expect("parse [Content_Types].xml");
    doc.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "Override")
        .filter_map(|n| {
            let part = n.attribute("PartName")?.to_string();
            let ct = n.attribute("ContentType")?.to_string();
            Some((part, ct))
        })
        .collect()
}

fn fixture_prefixed_content_types() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <ct:Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <ct:Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</ct:Types>"#
        .to_string()
}

#[test]
fn lazy_package_enforce_workbook_kind_template_rewrites_only_workbook_override() {
    let content_types = fixture_prefixed_content_types();
    let input = build_zip(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("xl/workbook.xml", b"<workbook/>"),
        ("xl/styles.xml", b"<styleSheet/>"),
        ("xl/worksheets/sheet1.xml", b"<worksheet/>"),
    ]);

    let mut pkg = XlsxLazyPackage::from_bytes(&input).expect("parse pkg");
    pkg.enforce_workbook_kind(WorkbookKind::Template)
        .expect("enforce kind");

    let updated_bytes = pkg
        .read_part("[Content_Types].xml")
        .expect("read part")
        .expect("content types present");
    let updated = std::str::from_utf8(&updated_bytes).unwrap();

    let original_overrides = content_types_override_map(&content_types);
    let mut expected_overrides = original_overrides.clone();
    expected_overrides.insert(
        "/xl/workbook.xml".to_string(),
        WorkbookKind::Template.workbook_content_type().to_string(),
    );

    let actual_overrides = content_types_override_map(updated);
    assert_eq!(actual_overrides, expected_overrides);

    // Prefix behavior: preserve the `ct:` prefix from the root for the workbook override.
    assert!(
        updated.contains("<ct:Override"),
        "expected output to preserve `ct:` prefix, got:\n{updated}"
    );
    assert!(
        !updated.contains("<Override"),
        "should not introduce unprefixed Override tags, got:\n{updated}"
    );
}

#[test]
fn lazy_package_enforce_workbook_kind_addin_rewrites_only_workbook_override() {
    let content_types = fixture_prefixed_content_types();
    let input = build_zip(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("xl/workbook.xml", b"<workbook/>"),
        ("xl/styles.xml", b"<styleSheet/>"),
        ("xl/worksheets/sheet1.xml", b"<worksheet/>"),
    ]);

    let mut pkg = XlsxLazyPackage::from_bytes(&input).expect("parse pkg");
    pkg.enforce_workbook_kind(WorkbookKind::MacroEnabledAddIn)
        .expect("enforce kind");

    let updated_bytes = pkg
        .read_part("[Content_Types].xml")
        .expect("read part")
        .expect("content types present");
    let updated = std::str::from_utf8(&updated_bytes).unwrap();

    let original_overrides = content_types_override_map(&content_types);
    let mut expected_overrides = original_overrides.clone();
    expected_overrides.insert(
        "/xl/workbook.xml".to_string(),
        WorkbookKind::MacroEnabledAddIn
            .workbook_content_type()
            .to_string(),
    );

    let actual_overrides = content_types_override_map(updated);
    assert_eq!(actual_overrides, expected_overrides);
}

#[test]
fn lazy_package_enforce_workbook_kind_is_noop_when_already_correct() {
    let workbook_ct = WorkbookKind::Template.workbook_content_type();
    let content_types = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override PartName="/xl/workbook.xml" ContentType="{workbook_ct}"/>
  <ct:Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</ct:Types>"#
    );
    let input = build_zip(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("xl/workbook.xml", b"<workbook/>"),
        ("xl/styles.xml", b"<styleSheet/>"),
    ]);

    let mut pkg = XlsxLazyPackage::from_bytes(&input).expect("parse pkg");
    let before = pkg
        .read_part("[Content_Types].xml")
        .expect("read part")
        .expect("content types present");

    pkg.enforce_workbook_kind(WorkbookKind::Template)
        .expect("enforce kind");

    let after = pkg
        .read_part("[Content_Types].xml")
        .expect("read part")
        .expect("content types present");
    assert_eq!(after, before, "expected no rewrite when already correct");
}

