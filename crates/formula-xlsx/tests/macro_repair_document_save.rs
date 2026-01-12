use std::io::{Cursor, Write};

use formula_xlsx::{load_from_bytes, XlsxPackage};
use roxmltree::Document;
use zip::write::FileOptions;
use zip::ZipWriter;

fn build_package(files: &[(&str, &[u8])]) -> Vec<u8> {
    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    for (name, bytes) in files {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

fn assert_content_type(doc: &Document<'_>, part_name: &str, expected: &str) {
    let node = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Override"
                && n.attribute("PartName") == Some(part_name)
        })
        .unwrap_or_else(|| panic!("missing [Content_Types].xml Override for {part_name}"));
    assert_eq!(
        node.attribute("ContentType"),
        Some(expected),
        "unexpected ContentType for {part_name}"
    );
}

fn assert_relationship(doc: &Document<'_>, rel_type: &str, target: &str) {
    let node = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "Relationship"
                && n.attribute("Type") == Some(rel_type)
                && n.attribute("Target") == Some(target)
        })
        .unwrap_or_else(|| panic!("missing Relationship type={rel_type} target={target}"));
    let _ = node;
}

#[test]
fn xlsxdocument_save_repairs_macro_content_types_and_relationships() {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    // Intentionally omit the vbaProject relationship to ensure we add it.
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></worksheet>"#;

    // Intentionally omit `[Content_Types].xml` overrides for VBA-related parts and omit
    // `xl/_rels/vbaProject.bin.rels` so the save path needs to repair the structure.
    let bytes = build_package(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("xl/workbook.xml", workbook_xml.as_bytes()),
        ("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes()),
        ("xl/worksheets/sheet1.xml", worksheet_xml.as_bytes()),
        ("xl/vbaProject.bin", b"fake-vba-project"),
        ("xl/vbaProjectSignature.bin", b"fake-signature"),
        ("xl/vbaData.xml", b"<vbaData/>"),
    ]);

    let doc = load_from_bytes(&bytes).expect("load xlsx document");
    let saved = doc.save_to_vec().expect("save xlsx document");

    let pkg = XlsxPackage::from_bytes(&saved).expect("read saved package");

    let ct_xml = std::str::from_utf8(pkg.part("[Content_Types].xml").unwrap()).unwrap();
    let ct = Document::parse(ct_xml).expect("parse [Content_Types].xml");
    assert_content_type(
        &ct,
        "/xl/workbook.xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    );
    assert_content_type(&ct, "/xl/vbaProject.bin", "application/vnd.ms-office.vbaProject");
    assert_content_type(
        &ct,
        "/xl/vbaProjectSignature.bin",
        "application/vnd.ms-office.vbaProjectSignature",
    );
    assert_content_type(&ct, "/xl/vbaData.xml", "application/vnd.ms-office.vbaData+xml");

    let workbook_rels_xml =
        std::str::from_utf8(pkg.part("xl/_rels/workbook.xml.rels").unwrap()).unwrap();
    let workbook_rels = Document::parse(workbook_rels_xml).expect("parse workbook.xml.rels");
    assert_relationship(
        &workbook_rels,
        "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
        "vbaProject.bin",
    );

    let vba_rels_xml =
        std::str::from_utf8(pkg.part("xl/_rels/vbaProject.bin.rels").unwrap()).unwrap();
    let vba_rels = Document::parse(vba_rels_xml).expect("parse vbaProject.bin.rels");
    assert_relationship(
        &vba_rels,
        "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature",
        "vbaProjectSignature.bin",
    );
}

