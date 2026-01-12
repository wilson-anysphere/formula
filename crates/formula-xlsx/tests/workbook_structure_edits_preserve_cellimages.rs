use std::io::{Cursor, Read, Write};

use formula_xlsx::load_from_bytes;
use quick_xml::events::Event;
use quick_xml::Reader;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|b| *b == b':') {
        Some(idx) => &name[idx + 1..],
        None => name,
    }
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn workbook_rels_has_target_suffix(rels_xml: &[u8], suffix: &str) -> bool {
    let mut reader = Reader::from_reader(rels_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).expect("read rels xml") {
            Event::Start(e) | Event::Empty(e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Relationship") =>
            {
                for attr in e.attributes().flatten() {
                    if local_name(attr.key.as_ref()).eq_ignore_ascii_case(b"Target") {
                        let target = attr.unescape_value().expect("Target attr").into_owned();
                        if target.ends_with(suffix) {
                            return true;
                        }
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    false
}

fn content_types_override_content_type(content_types_xml: &[u8], part_name: &str) -> Option<String> {
    let mut reader = Reader::from_reader(content_types_xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf).expect("read content types xml") {
            Event::Start(e) | Event::Empty(e)
                if local_name(e.name().as_ref()).eq_ignore_ascii_case(b"Override") =>
            {
                let mut seen_part_name = None;
                let mut seen_content_type = None;
                for attr in e.attributes().flatten() {
                    let key = local_name(attr.key.as_ref());
                    let val = attr.unescape_value().expect("attr").into_owned();
                    if key.eq_ignore_ascii_case(b"PartName") {
                        seen_part_name = Some(val);
                    } else if key.eq_ignore_ascii_case(b"ContentType") {
                        seen_content_type = Some(val);
                    }
                }
                if seen_part_name.as_deref() == Some(part_name) {
                    return seen_content_type;
                }
            }
            Event::Eof => break,
            _ => {}
        }
        buf.clear();
    }
    None
}

struct CellImagesFixture {
    xlsx: Vec<u8>,
    cellimages_xml: Vec<u8>,
    cellimages_rels: Vec<u8>,
    image1_png: Vec<u8>,
    cellimages_content_type: &'static str,
}

fn build_minimal_xlsx_with_cellimages_relationship() -> CellImagesFixture {
    // 1x1 transparent PNG.
    // Generated once and embedded as raw bytes so the round-trip can be validated byte-for-byte.
    const IMAGE1_PNG: &[u8] = &[
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
        0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
        0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78,
        0x9C, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    let cellimages_content_type = "application/vnd.ms-excel.cellimages+xml";

    let content_types = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="{cellimages_content_type}"/>
</Types>
"#
    );

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

    // Include a non-standard relationship type for cellimages so the test doesn't depend on MS URIs.
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://example.com/relationships/cellimages" Target="cellimages.xml"/>
</Relationships>
"#;

    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"#;

    // Minimal `xl/cellimages.xml` that references an image relationship ID.
    let cellimages_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<etc:cellImages xmlns:etc="http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages"
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <etc:cellImage>
    <xdr:pic>
      <xdr:blipFill>
        <a:blip r:embed="rId1"/>
      </xdr:blipFill>
    </xdr:pic>
  </etc:cellImage>
</etc:cellImages>
"#;

    let cellimages_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>
"#;

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
    zip.write_all(cellimages_xml).expect("zip write");

    zip.start_file("xl/_rels/cellimages.xml.rels", options)
        .expect("zip file");
    zip.write_all(cellimages_rels).expect("zip write");

    zip.start_file("xl/media/image1.png", options)
        .expect("zip file");
    zip.write_all(IMAGE1_PNG).expect("zip write");

    let xlsx = zip.finish().expect("finish zip").into_inner();

    CellImagesFixture {
        xlsx,
        cellimages_xml: cellimages_xml.to_vec(),
        cellimages_rels: cellimages_rels.to_vec(),
        image1_png: IMAGE1_PNG.to_vec(),
        cellimages_content_type,
    }
}

fn assert_preserves_cellimages_parts_and_package_links(
    saved: &[u8],
    expected: &CellImagesFixture,
) {
    assert_eq!(
        zip_part(saved, "xl/cellimages.xml"),
        expected.cellimages_xml,
        "expected xl/cellimages.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(saved, "xl/_rels/cellimages.xml.rels"),
        expected.cellimages_rels,
        "expected xl/_rels/cellimages.xml.rels to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(saved, "xl/media/image1.png"),
        expected.image1_png,
        "expected xl/media/image1.png to be preserved byte-for-byte"
    );

    let workbook_rels_xml = zip_part(saved, "xl/_rels/workbook.xml.rels");
    assert!(
        workbook_rels_has_target_suffix(&workbook_rels_xml, "cellimages.xml"),
        "expected xl/_rels/workbook.xml.rels to retain a relationship targeting cellimages.xml"
    );

    let content_types_xml = zip_part(saved, "[Content_Types].xml");
    let content_type = content_types_override_content_type(&content_types_xml, "/xl/cellimages.xml")
        .expect("expected [Content_Types].xml to contain an Override for /xl/cellimages.xml");
    assert_eq!(
        content_type, expected.cellimages_content_type,
        "expected [Content_Types].xml cellimages override ContentType to be preserved"
    );
}

#[test]
fn workbook_structure_edits_preserve_cellimages_parts_and_relationships() {
    let fixture = build_minimal_xlsx_with_cellimages_relationship();

    let mut doc = load_from_bytes(&fixture.xlsx).expect("load fixture");

    // Add a new sheet, forcing the writer to rewrite workbook.xml, workbook.xml.rels, and
    // [Content_Types].xml.
    doc.workbook.add_sheet("Added").expect("add sheet");

    let saved = doc.save_to_vec().expect("save after add sheet");
    assert_preserves_cellimages_parts_and_package_links(&saved, &fixture);

    // Delete the original sheet to force a second workbook structure edit path.
    let sheet1_id = doc
        .workbook
        .sheet_by_name("Sheet1")
        .expect("Sheet1 exists")
        .id;
    doc.workbook.delete_sheet(sheet1_id).expect("delete sheet");

    let saved_after_delete = doc.save_to_vec().expect("save after delete sheet");
    assert_preserves_cellimages_parts_and_package_links(&saved_after_delete, &fixture);
}

