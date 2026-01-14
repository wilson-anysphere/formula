use std::collections::BTreeMap;
use std::io::{Cursor, Read, Write};

use formula_xlsx::{WorkbookKind, XlsxLazyPackage};
use roxmltree::Document;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

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

fn read_zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let mut zip = ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let mut file = zip.by_name(name).unwrap();
    let mut out = Vec::new();
    file.read_to_end(&mut out).unwrap();
    out
}

fn zip_has_part(zip_bytes: &[u8], name: &str) -> bool {
    let mut zip = ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let has = zip.by_name(name).is_ok();
    has
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

fn fixture_prefixed_content_types_macro_enabled() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <ct:Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
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

#[test]
fn lazy_package_strip_macros_then_enforce_kind_updates_content_types() {
    let content_types = fixture_prefixed_content_types_macro_enabled();
    let input = build_zip(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("xl/workbook.xml", b"<workbook/>"),
        ("xl/vbaProject.bin", b"fake-vba"),
    ]);

    let mut pkg = XlsxLazyPackage::from_bytes(&input).expect("parse pkg");
    pkg.remove_vba_project().expect("strip macros");
    pkg.enforce_workbook_kind(WorkbookKind::Template)
        .expect("enforce kind");

    let out = pkg.write_to_bytes().expect("write");

    assert!(
        !zip_has_part(&out, "xl/vbaProject.bin"),
        "expected macro part to be stripped"
    );

    let ct_bytes = read_zip_part(&out, "[Content_Types].xml");
    let ct_xml = std::str::from_utf8(&ct_bytes).unwrap();
    let overrides = content_types_override_map(ct_xml);
    assert!(
        !ct_xml.contains("vbaProject.bin"),
        "expected macro-strip to remove the vbaProject content type override, got:\n{ct_xml}"
    );
    assert_eq!(
        overrides.get("/xl/workbook.xml").map(String::as_str),
        Some(WorkbookKind::Template.workbook_content_type()),
        "expected workbook override to be rewritten after macro strip"
    );
}

#[test]
fn lazy_package_remove_vba_project_patches_existing_content_types_override() {
    let content_types = fixture_prefixed_content_types_macro_enabled();
    let input = build_zip(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("xl/workbook.xml", b"<workbook/>"),
        ("xl/vbaProject.bin", b"fake-vba"),
    ]);

    let mut pkg = XlsxLazyPackage::from_bytes(&input).expect("parse pkg");
    // Create a stale workbook override by enforcing a different kind first.
    pkg.enforce_workbook_kind(WorkbookKind::MacroEnabledAddIn)
        .expect("enforce kind");
    pkg.remove_vba_project().expect("strip macros");

    let out = pkg.write_to_bytes().expect("write");

    let ct_bytes = read_zip_part(&out, "[Content_Types].xml");
    let ct_xml = std::str::from_utf8(&ct_bytes).unwrap();
    let overrides = content_types_override_map(ct_xml);
    assert!(
        !ct_xml.contains("vbaProject.bin"),
        "expected macro-strip to remove the vbaProject content type override even when a [Content_Types].xml override existed, got:\n{ct_xml}"
    );
    assert_eq!(
        overrides.get("/xl/workbook.xml").map(String::as_str),
        Some(WorkbookKind::Workbook.workbook_content_type()),
        "expected remove_vba_project to keep [Content_Types].xml override consistent with macro-free target"
    );
}

#[test]
fn lazy_package_enforce_workbook_kind_updates_existing_content_types_override_while_stripping_macros(
) {
    let content_types = fixture_prefixed_content_types_macro_enabled();
    let input = build_zip(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("xl/workbook.xml", b"<workbook/>"),
        ("xl/vbaProject.bin", b"fake-vba"),
    ]);

    let mut pkg = XlsxLazyPackage::from_bytes(&input).expect("parse pkg");
    // Seed an explicit [Content_Types].xml override.
    pkg.enforce_workbook_kind(WorkbookKind::MacroEnabledAddIn)
        .expect("enforce kind");
    pkg.remove_vba_project().expect("strip macros");
    // With macro stripping enabled, enforce_workbook_kind should still keep any existing content
    // types override consistent so it doesn't overwrite the macro-strip pass output.
    pkg.enforce_workbook_kind(WorkbookKind::Template)
        .expect("enforce kind");

    let out = pkg.write_to_bytes().expect("write");

    let ct_bytes = read_zip_part(&out, "[Content_Types].xml");
    let ct_xml = std::str::from_utf8(&ct_bytes).unwrap();
    let overrides = content_types_override_map(ct_xml);
    assert!(
        !ct_xml.contains("vbaProject.bin"),
        "expected macro-strip to remove the vbaProject content type override even when updating a [Content_Types].xml override, got:\n{ct_xml}"
    );
    assert_eq!(
        overrides.get("/xl/workbook.xml").map(String::as_str),
        Some(WorkbookKind::Template.workbook_content_type()),
        "expected enforce_workbook_kind to update existing content types override while macro stripping"
    );
}
