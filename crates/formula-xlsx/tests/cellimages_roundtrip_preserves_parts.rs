use std::io::{Cursor, Read, Write};

use formula_xlsx::load_from_bytes;
use pretty_assertions::assert_eq;
use zip::write::FileOptions;
use zip::{ZipArchive, ZipWriter};

fn build_cellimages_fixture_xlsx() -> (Vec<u8>, Vec<u8>, Vec<u8>, Vec<u8>) {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/cellImages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
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

    // Include a stub relationship to the cellImages part to reflect how real files are wired.
    // The relationship type is intentionally a placeholder; the test validates preservation.
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2020/07/relationships/cellImages" Target="cellImages.xml"/>
</Relationships>
"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
"#;

    let cell_images_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2020/07/main"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage>
    <a:blip r:embed="rId1"/>
  </cellImage>
  <cellImage>
    <a:blip r:embed="rId1"/>
  </cellImage>
</cellImages>
"#
    .to_vec();

    let cell_images_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>
"#
    .to_vec();

    // Binary payload with non-UTF8 bytes to ensure we preserve raw blobs byte-for-byte.
    let image1_png = vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
        0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
        0x00, 0x1F, 0x15, 0xC4, 0x89,
    ];

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options)
        .expect("zip file");
    zip.write_all(content_types.as_bytes()).expect("zip write");

    zip.start_file("_rels/.rels", options).expect("zip file");
    zip.write_all(root_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/workbook.xml", options)
        .expect("zip file");
    zip.write_all(workbook_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("zip file");
    zip.write_all(workbook_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("zip file");
    zip.write_all(worksheet_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/cellImages.xml", options)
        .expect("zip file");
    zip.write_all(&cell_images_xml).expect("zip write");

    zip.start_file("xl/_rels/cellImages.xml.rels", options)
        .expect("zip file");
    zip.write_all(&cell_images_rels).expect("zip write");

    zip.start_file("xl/media/image1.png", options)
        .expect("zip file");
    zip.write_all(&image1_png).expect("zip write");

    let bytes = zip.finish().expect("finish zip").into_inner();
    (bytes, cell_images_xml, cell_images_rels, image1_png)
}

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

#[test]
fn cellimages_roundtrip_preserves_parts() {
    let (fixture, cell_images_xml, cell_images_rels, image1_png) = build_cellimages_fixture_xlsx();

    let doc = load_from_bytes(&fixture).expect("load fixture");
    let saved = doc.save_to_vec().expect("save");

    assert_eq!(zip_part(&saved, "xl/cellImages.xml"), cell_images_xml);
    assert_eq!(zip_part(&saved, "xl/_rels/cellImages.xml.rels"), cell_images_rels);
    assert_eq!(zip_part(&saved, "xl/media/image1.png"), image1_png);
}

