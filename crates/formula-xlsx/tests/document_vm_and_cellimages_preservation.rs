use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn build_fixture_xlsx_with_vm_and_cellimages() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml"/>
</Types>
"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"#;

    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"#;

    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" vm="9" customAttr="x"><v>1</v></c></row>
  </sheetData>
</worksheet>
"#;

    // This is not a fully-specified cell images payload; the test is about preservation.
    let cellimages_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:embed="rId1"/>
</cellImages>
"#;

    let cellimages_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>
"#;

    // 1x1 transparent PNG.
    // Generated once and embedded as raw bytes so the round-trip can be validated byte-for-byte.
    const IMAGE1_PNG: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
        0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
        0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78,
        0x9C, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options)
        .expect("zip file");
    zip.write_all(content_types.as_bytes()).expect("zip write");

    zip.start_file("_rels/.rels", options).expect("zip file");
    zip.write_all(root_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/workbook.xml", options).expect("zip file");
    zip.write_all(workbook_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("zip file");
    zip.write_all(workbook_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("zip file");
    zip.write_all(sheet1_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/cellimages.xml", options)
        .expect("zip file");
    zip.write_all(cellimages_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/_rels/cellimages.xml.rels", options)
        .expect("zip file");
    zip.write_all(cellimages_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/media/image1.png", options)
        .expect("zip file");
    zip.write_all(IMAGE1_PNG).expect("zip write");

    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn preserves_cellimages_parts_through_document_roundtrip_even_when_vm_is_dropped() {
    let fixture = build_fixture_xlsx_with_vm_and_cellimages();

    // Capture original bytes for the parts that must survive byte-for-byte.
    let original_cellimages = zip_part(&fixture, "xl/cellimages.xml");
    let original_cellimages_rels = zip_part(&fixture, "xl/_rels/cellimages.xml.rels");
    let original_image_png = zip_part(&fixture, "xl/media/image1.png");

    let mut doc = load_from_bytes(&fixture).expect("load fixture");

    let sheet_id = doc.workbook.sheet_by_name("Sheet1").expect("Sheet1").id;
    assert!(
        doc.set_cell_value(
            sheet_id,
            CellRef::from_a1("A1").unwrap(),
            CellValue::Number(2.0)
        ),
        "expected set_cell_value to succeed"
    );

    let saved = doc.save_to_vec().expect("save");

    let sheet_xml = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let sheet_xml_str = std::str::from_utf8(&sheet_xml).expect("sheet1.xml utf-8");
    let parsed = roxmltree::Document::parse(sheet_xml_str).expect("parse sheet1.xml");

    let cell = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");

    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .expect("expected <v> in A1");
    assert_eq!(
        v.text(),
        Some("2"),
        "expected edited cell value to be written, got: {sheet_xml_str}"
    );

    assert_eq!(
        cell.attribute("vm"),
        Some("9"),
        "expected vm attribute to be preserved, got: {sheet_xml_str}"
    );
    assert_eq!(
        cell.attribute("customAttr"),
        Some("x"),
        "expected unrelated customAttr attribute to be preserved, got: {sheet_xml_str}"
    );

    assert_eq!(
        zip_part(&saved, "xl/cellimages.xml"),
        original_cellimages,
        "expected xl/cellimages.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&saved, "xl/_rels/cellimages.xml.rels"),
        original_cellimages_rels,
        "expected xl/_rels/cellimages.xml.rels to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(&saved, "xl/media/image1.png"),
        original_image_png,
        "expected xl/media/image1.png to be preserved byte-for-byte"
    );
}
