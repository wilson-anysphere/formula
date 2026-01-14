use std::io::{Cursor, Read as _, Write as _};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{PackageCellPatch, XlsxPackage};

fn build_minimal_xlsx_with_cellimages() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2020/relationships/cellImages" Target="cellimages.xml"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    let worksheet_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;

    // We don't currently parse this schema, but we must preserve it verbatim.
    let cellimages_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/11/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage>
    <a:blip r:embed="rId1"/>
  </cellImage>
</cellImages>
"#;

    let cellimages_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.ms-excel.cellimages+xml"/>
</Types>"#;

    let image_bytes: &[u8] = b"\x89PNG\r\n\x1a\n\x00fake-png-bytes";

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options =
        zip::write::FileOptions::<()>::default().compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/_rels/sheet1.xml.rels", options)
        .unwrap();
    zip.write_all(worksheet_rels.as_bytes()).unwrap();

    zip.start_file("xl/cellimages.xml", options).unwrap();
    zip.write_all(cellimages_xml).unwrap();

    zip.start_file("xl/_rels/cellimages.xml.rels", options).unwrap();
    zip.write_all(cellimages_rels.as_bytes()).unwrap();

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(image_bytes).unwrap();

    zip.finish().unwrap().into_inner()
}

fn read_zip_entry(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(zip_bytes)).unwrap();
    let mut file = archive.by_name(name).unwrap();
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).unwrap();
    buf
}

#[test]
fn preserve_cellimages_parts_byte_for_byte_on_streaming_patch(
) -> Result<(), Box<dyn std::error::Error>> {
    let original_zip = build_minimal_xlsx_with_cellimages();

    let before_cellimages = read_zip_entry(&original_zip, "xl/cellimages.xml");
    let before_cellimages_rels = read_zip_entry(&original_zip, "xl/_rels/cellimages.xml.rels");
    let before_image = read_zip_entry(&original_zip, "xl/media/image1.png");

    let pkg = XlsxPackage::from_bytes(&original_zip)?;
    let patch = PackageCellPatch::for_sheet_name(
        "Sheet1",
        CellRef::from_a1("A1")?,
        CellValue::Number(2.0),
        Some("=1+1".to_string()),
    );
    let out_bytes = pkg.apply_cell_patches_to_bytes(&[patch])?;

    // Ensure the parts still exist, and that they are preserved byte-for-byte.
    let after_cellimages = read_zip_entry(&out_bytes, "xl/cellimages.xml");
    let after_cellimages_rels = read_zip_entry(&out_bytes, "xl/_rels/cellimages.xml.rels");
    let after_image = read_zip_entry(&out_bytes, "xl/media/image1.png");

    assert_eq!(
        before_cellimages, after_cellimages,
        "expected xl/cellimages.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        before_cellimages_rels, after_cellimages_rels,
        "expected xl/_rels/cellimages.xml.rels to be preserved byte-for-byte"
    );
    assert_eq!(
        before_image, after_image,
        "expected xl/media/image1.png to be preserved byte-for-byte"
    );

    let content_types = String::from_utf8(read_zip_entry(&out_bytes, "[Content_Types].xml"))?;
    assert!(
        content_types.contains(r#"PartName="/xl/cellimages.xml""#),
        "expected [Content_Types].xml to retain the /xl/cellimages.xml override, got: {content_types}"
    );

    Ok(())
}
