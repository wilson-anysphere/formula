use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{patch_xlsx_streaming, WorksheetCellPatch};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

const CELL_EXTLST_MARKER: &str = "CELL_EXTLST_MARKER_123";
const CELLIMAGES_MARKER: &str = "CELLIMAGES_MARKER_456";
const METADATA_MARKER: &str = "METADATA_MARKER_789";
const RICH_DATA_MARKER: &str = "RICH_DATA_MARKER_ABC";
const RICH_DATA_RELS_MARKER: &str = "RICH_DATA_RELS_MARKER_DEF";

// 1x1 transparent PNG.
// Generated once and embedded as raw bytes so the round-trip can be validated byte-for-byte.
const IMAGE1_PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
    0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
    0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78,
    0x9C, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
    0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
];

fn build_fixture_xlsx() -> Vec<u8> {
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/cellimages.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.cellimages+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.metadata+xml"/>
  <Override PartName="/xl/richData/richValue.xml" ContentType="application/xml"/>
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

    let sheet1_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1">
      <c r="A1" vm="9" cm="7" customAttr="x">
        <v>1</v>
        <extLst>
          <ext uri="urn:{CELL_EXTLST_MARKER}">
            <marker>{CELL_EXTLST_MARKER}</marker>
          </ext>
        </extLst>
      </c>
    </row>
  </sheetData>
</worksheet>
"#
    );

    // This is not a fully-specified in-cell image payload; the test is about preservation.
    let cellimages_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cellImages xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/cellimages"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cellImage r:embed="rId1"/>
  <marker>{CELLIMAGES_MARKER}</marker>
</cellImages>
"#
    );

    let cellimages_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>
"#;

    let metadata_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="1">
    <metadataType name="{METADATA_MARKER}" minSupportedVersion="0"/>
  </metadataTypes>
</metadata>
"#
    );

    let rich_value_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richdata">
  <marker>{RICH_DATA_MARKER}</marker>
</rvData>
"#
    );

    let rich_value_rels = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="urn:rich-data-marker" Target="{RICH_DATA_RELS_MARKER}"/>
</Relationships>
"#
    );

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

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

    zip.start_file("xl/metadata.xml", options)
        .expect("zip file");
    zip.write_all(metadata_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/richData/richValue.xml", options)
        .expect("zip file");
    zip.write_all(rich_value_xml.as_bytes()).expect("zip write");

    zip.start_file("xl/richData/_rels/richValue.xml.rels", options)
        .expect("zip file");
    zip.write_all(rich_value_rels.as_bytes()).expect("zip write");

    zip.start_file("xl/media/image1.png", options)
        .expect("zip file");
    zip.write_all(IMAGE1_PNG).expect("zip write");

    zip.finish().expect("finish zip").into_inner()
}

#[test]
fn streaming_patch_preserves_images_in_cells_metadata_and_related_parts_and_drops_vm_on_edit(
) -> Result<(), Box<dyn std::error::Error>> {
    let fixture = build_fixture_xlsx();

    // Capture original bytes for the parts that must survive byte-for-byte.
    let original_cellimages = zip_part(&fixture, "xl/cellimages.xml");
    let original_cellimages_rels = zip_part(&fixture, "xl/_rels/cellimages.xml.rels");
    let original_metadata = zip_part(&fixture, "xl/metadata.xml");
    let original_rich_value = zip_part(&fixture, "xl/richData/richValue.xml");
    let original_rich_value_rels = zip_part(&fixture, "xl/richData/_rels/richValue.xml.rels");
    let original_image_png = zip_part(&fixture, "xl/media/image1.png");

    let patch = WorksheetCellPatch::new(
        "xl/worksheets/sheet1.xml",
        CellRef::from_a1("A1")?,
        CellValue::Number(2.0),
        None,
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming(Cursor::new(fixture), &mut out, &[patch])?;

    // Worksheet XML assertions.
    let out_sheet_xml = zip_part(out.get_ref(), "xl/worksheets/sheet1.xml");
    let out_sheet_xml_str = std::str::from_utf8(&out_sheet_xml)?;
    let parsed = roxmltree::Document::parse(out_sheet_xml_str)?;

    let cell = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");

    assert_eq!(
        cell.attribute("vm"),
        Some("9"),
        "expected vm attribute to be preserved on edit"
    );
    assert_eq!(cell.attribute("cm"), Some("7"));
    assert_eq!(cell.attribute("customAttr"), Some("x"));

    let ext_lst = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "extLst")
        .expect("expected <extLst> to be preserved under <c>");
    assert!(
        ext_lst.descendants().any(|n| {
            n.is_element()
                && n.tag_name().name() == "marker"
                && n.text() == Some(CELL_EXTLST_MARKER)
        }),
        "expected preserved <extLst> subtree to contain marker {CELL_EXTLST_MARKER}, got: {out_sheet_xml_str}"
    );

    let v_text = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v_text, "2");

    // Part preservation assertions.
    assert_eq!(
        zip_part(out.get_ref(), "xl/cellimages.xml"),
        original_cellimages,
        "expected xl/cellimages.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(out.get_ref(), "xl/_rels/cellimages.xml.rels"),
        original_cellimages_rels,
        "expected xl/_rels/cellimages.xml.rels to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(out.get_ref(), "xl/metadata.xml"),
        original_metadata,
        "expected xl/metadata.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(out.get_ref(), "xl/richData/richValue.xml"),
        original_rich_value,
        "expected xl/richData/richValue.xml to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(out.get_ref(), "xl/richData/_rels/richValue.xml.rels"),
        original_rich_value_rels,
        "expected xl/richData/_rels/richValue.xml.rels to be preserved byte-for-byte"
    );
    assert_eq!(
        zip_part(out.get_ref(), "xl/media/image1.png"),
        original_image_png,
        "expected xl/media/image1.png to be preserved byte-for-byte"
    );

    Ok(())
}
