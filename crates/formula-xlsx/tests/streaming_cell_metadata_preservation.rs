use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{patch_xlsx_streaming, WorksheetCellPatch};
use zip::ZipArchive;

fn build_minimal_xlsx_with_cell_images(sheet_xml: &str) -> (Vec<u8>, Vec<u8>, Vec<u8>, Vec<u8>) {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

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
</Relationships>"#;

    // Minimal cell images parts. These are expected to be raw-copied (byte-for-byte) by the
    // streaming patcher when only editing a worksheet XML part.
    let cellimages_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/cellimages"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:id="rId1"/>
</cellImages>"#
        .to_vec();

    let cellimages_rels = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"#
        .to_vec();

    // A tiny valid PNG (1x1 px). Any bytes would work, but keeping it valid makes the fixture
    // easier to reason about.
    let image1_png: Vec<u8> = vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
        0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
        0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08,
        0xD7, 0x63, 0xF8, 0xFF, 0xFF, 0x3F, 0x00, 0x05, 0xFE, 0x02, 0xFE, 0x41, 0xD3, 0x8D,
        0x90, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
    ];

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/cellimages.xml", options).unwrap();
    zip.write_all(&cellimages_xml).unwrap();

    zip.start_file("xl/_rels/cellimages.xml.rels", options)
        .unwrap();
    zip.write_all(&cellimages_rels).unwrap();

    zip.start_file("xl/media/image1.png", options).unwrap();
    zip.write_all(&image1_png).unwrap();

    (
        zip.finish().unwrap().into_inner(),
        cellimages_xml,
        cellimages_rels,
        image1_png,
    )
}

fn zip_part(bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

#[test]
fn streaming_patch_preserves_cell_vm_cm_and_extlst_and_raw_copies_cellimages_parts(
) -> Result<(), Box<dyn std::error::Error>> {
    let extlst =
        r#"<extLst><ext uri="{123}"><test xmlns="http://example.com">ok</test></ext></extLst>"#;
    let worksheet_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" cm="7" vm="9" customAttr="x"><v>1</v>{extlst}</c></row>
  </sheetData>
</worksheet>"#
    );

    let (in_bytes, cellimages_xml, cellimages_rels, image1_png) =
        build_minimal_xlsx_with_cell_images(&worksheet_xml);

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::Number(2.0),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(in_bytes), &mut out, &[patch])?;

    let out_bytes = out.get_ref();

    let out_sheet_xml = String::from_utf8(zip_part(out_bytes, "xl/worksheets/sheet1.xml"))?;
    let doc = roxmltree::Document::parse(&out_sheet_xml)?;
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");

    assert_eq!(cell.attribute("cm"), Some("7"));
    assert_eq!(cell.attribute("vm"), Some("9"));
    assert_eq!(cell.attribute("customAttr"), Some("x"));

    assert!(
        out_sheet_xml.contains(extlst),
        "expected extLst subtree to be preserved byte-for-byte, got: {out_sheet_xml}"
    );

    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v, "2", "expected cached value to update, got: {out_sheet_xml}");

    // These parts are unrelated to the patched worksheet and must be preserved.
    assert_eq!(
        zip_part(out_bytes, "xl/cellimages.xml"),
        cellimages_xml,
        "expected xl/cellimages.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(out_bytes, "xl/_rels/cellimages.xml.rels"),
        cellimages_rels,
        "expected xl/_rels/cellimages.xml.rels to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(out_bytes, "xl/media/image1.png"),
        image1_png,
        "expected xl/media/image1.png to be preserved byte-for-byte"
    );

    Ok(())
}

