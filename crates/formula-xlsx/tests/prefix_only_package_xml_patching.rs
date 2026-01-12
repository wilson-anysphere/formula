use std::io::{Cursor, Read, Write};

use formula_xlsx::{strip_vba_project_streaming, DateSystem, XlsxPackage};
use roxmltree::Document;

fn build_package(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in entries {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn read_zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let mut zip = zip::ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let mut file = zip.by_name(name).unwrap();
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).unwrap();
    buf
}

#[test]
fn prefix_only_workbook_set_date_system_inserts_prefixed_workbookpr() {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </x:sheets>
</x:workbook>"#;

    let bytes = build_package(&[("xl/workbook.xml", workbook_xml.as_bytes())]);
    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    pkg.set_workbook_date_system(DateSystem::V1904)
        .expect("set date system");

    let updated = std::str::from_utf8(pkg.part("xl/workbook.xml").unwrap()).unwrap();
    let doc = Document::parse(updated).expect("updated workbook.xml parses");

    // Ensure we inserted a prefixed workbookPr (and did not introduce a namespace-less element).
    assert!(updated.contains("<x:workbookPr"));
    assert!(!updated.contains("<workbookPr"));

    let spreadsheetml = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let workbook_pr: Vec<_> = doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "workbookPr")
        .collect();
    assert_eq!(workbook_pr.len(), 1);
    assert_eq!(workbook_pr[0].tag_name().namespace(), Some(spreadsheetml));
    assert_eq!(workbook_pr[0].attribute("date1904"), Some("1"));
}

#[test]
fn prefix_only_workbook_rels_write_inserts_vba_relationship_with_correct_namespace() {
    let content_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</ct:Types>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
  <pr:Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</pr:Relationships>"#;

    let bytes = build_package(&[
        ("[Content_Types].xml", content_types_xml.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/vbaProject.bin", b"fake-vba-project"),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    let written = pkg.write_to_bytes().expect("write package");

    let updated_rels = read_zip_part(&written, "xl/_rels/workbook.xml.rels");
    let updated_rels = std::str::from_utf8(&updated_rels).unwrap();
    let doc = Document::parse(updated_rels).expect("updated workbook.xml.rels parses");

    // Ensure we inserted a namespaced (prefixed) Relationship element.
    assert!(!updated_rels.contains("<Relationship"));

    let rel_type = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
    let rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships";
    let vba_rel = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Type") == Some(rel_type)
        })
        .expect("vbaProject relationship exists");
    assert_eq!(vba_rel.tag_name().namespace(), Some(rels_ns));
    assert_eq!(vba_rel.attribute("Target"), Some("vbaProject.bin"));
}

#[test]
fn prefix_only_content_types_adds_vba_override_and_macro_enabled_workbook_type() {
    let content_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</ct:Types>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
  <pr:Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</pr:Relationships>"#;

    let bytes = build_package(&[
        ("[Content_Types].xml", content_types_xml.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/vbaProject.bin", b"fake-vba-project"),
    ]);

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    let written = pkg.write_to_bytes().expect("write package");

    let updated_ct = read_zip_part(&written, "[Content_Types].xml");
    let updated_ct = std::str::from_utf8(&updated_ct).unwrap();
    let doc = Document::parse(updated_ct).expect("updated [Content_Types].xml parses");

    assert!(!updated_ct.contains("<Override"));

    let ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types";
    let vba_override = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Override"
                && n.attribute("PartName") == Some("/xl/vbaProject.bin")
        })
        .expect("vbaProject override exists");
    assert_eq!(vba_override.tag_name().namespace(), Some(ct_ns));
    assert_eq!(
        vba_override.attribute("ContentType"),
        Some("application/vnd.ms-office.vbaProject")
    );

    let workbook_override = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Override"
                && n.attribute("PartName") == Some("/xl/workbook.xml")
        })
        .expect("workbook override exists");
    assert_eq!(
        workbook_override.attribute("ContentType"),
        Some("application/vnd.ms-excel.sheet.macroEnabled.main+xml")
    );
}

#[test]
fn prefix_only_content_types_macro_strip_preserves_override_prefix() {
    let content_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <ct:Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
</ct:Types>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
  <pr:Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</pr:Relationships>"#;

    let bytes = build_package(&[
        ("[Content_Types].xml", content_types_xml.as_bytes()),
        ("xl/workbook.xml", workbook_xml.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/vbaProject.bin", b"fake-vba-project"),
    ]);

    let mut pkg = XlsxPackage::from_bytes(&bytes).expect("read package");
    pkg.remove_vba_project().expect("strip macros");

    assert!(pkg.part("xl/vbaProject.bin").is_none());

    let updated_ct = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap()).unwrap();
    let doc = Document::parse(updated_ct).expect("updated [Content_Types].xml parses");

    // Ensure we didn't rewrite the patched workbook Override as an unprefixed element (which would
    // be namespace-less in prefix-only documents).
    assert!(!updated_ct.contains("<Override"));

    let ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types";
    let workbook_override = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Override"
                && n.attribute("PartName") == Some("/xl/workbook.xml")
        })
        .expect("workbook override exists");
    assert_eq!(workbook_override.tag_name().namespace(), Some(ct_ns));
    assert_eq!(
        workbook_override.attribute("ContentType"),
        Some("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    );

    assert!(
        doc.descendants().all(|n| {
            !(n.is_element()
                && n.tag_name().name() == "Override"
                && n.attribute("PartName") == Some("/xl/vbaProject.bin"))
        }),
        "expected vbaProject override to be removed"
    );
}

#[test]
fn prefix_only_content_types_macro_strip_streaming_preserves_override_prefix() {
    let content_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<ct:Types xmlns:ct="http://schemas.openxmlformats.org/package/2006/content-types">
  <ct:Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <ct:Default Extension="xml" ContentType="application/xml"/>
  <ct:Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
  <ct:Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>
</ct:Types>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
  <pr:Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>
</pr:Relationships>"#;

    let bytes = build_package(&[
        ("[Content_Types].xml", content_types_xml.as_bytes()),
        ("xl/workbook.xml", workbook_xml.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/vbaProject.bin", b"fake-vba-project"),
    ]);

    let mut out = Cursor::new(Vec::new());
    strip_vba_project_streaming(Cursor::new(bytes), &mut out).expect("streaming macro strip");
    let written = out.into_inner();

    let mut zip = zip::ZipArchive::new(Cursor::new(written.as_slice())).unwrap();
    assert!(
        zip.by_name("xl/vbaProject.bin").is_err(),
        "expected vbaProject.bin to be removed"
    );

    let updated_ct = read_zip_part(&written, "[Content_Types].xml");
    let updated_ct = std::str::from_utf8(&updated_ct).unwrap();
    let doc = Document::parse(updated_ct).expect("updated [Content_Types].xml parses");

    // Ensure we didn't rewrite the patched workbook Override as an unprefixed element (which would
    // be namespace-less in prefix-only documents).
    assert!(!updated_ct.contains("<Override"));

    let ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types";
    let workbook_override = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Override"
                && n.attribute("PartName") == Some("/xl/workbook.xml")
        })
        .expect("workbook override exists");
    assert_eq!(workbook_override.tag_name().namespace(), Some(ct_ns));
    assert_eq!(
        workbook_override.attribute("ContentType"),
        Some("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    );

    assert!(
        doc.descendants().all(|n| {
            !(n.is_element()
                && n.tag_name().name() == "Override"
                && n.attribute("PartName") == Some("/xl/vbaProject.bin"))
        }),
        "expected vbaProject override to be removed"
    );
}
