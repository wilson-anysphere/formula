use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{CellPatch, WorkbookCellPatches, XlsxPackage};

fn build_rich_data_package(metadata_xml: &[u8], rich_data_parts: &[(&str, &[u8])]) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    // Include a metadata relationship that must be preserved verbatim when applying cell patches.
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rIdMetadata" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata" Target="metadata.xml"/>
</Relationships>"#;

    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    let content_types_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
  <Override PartName="/xl/richData/richValue1.xml" ContentType="application/vnd.ms-excel.richvalue+xml"/>
  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvalue+xml"/>
</Types>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types_xml.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options).unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/metadata.xml", options).unwrap();
    zip.write_all(metadata_xml).unwrap();

    for (name, bytes) in rich_data_parts {
        zip.start_file(*name, options).unwrap();
        zip.write_all(bytes).unwrap();
    }

    zip.finish().unwrap().into_inner()
}

#[test]
fn apply_cell_patches_preserves_rich_data_parts_and_workbook_rels() {
    let metadata_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <metadataTypes count="0"/>
</metadata>"#;

    let rich_value_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rv:richValue xmlns:rv="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv:value>hello</rv:value>
</rv:richValue>"#;

    let rich_value_rel_xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rv:richValueRel xmlns:rv="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <rv:rel>1</rv:rel>
</rv:richValueRel>"#;

    let rich_data_parts = vec![
        ("xl/richData/richValue1.xml", rich_value_xml.as_slice()),
        ("xl/richData/richValueRel.xml", rich_value_rel_xml.as_slice()),
    ];

    let input_bytes = build_rich_data_package(metadata_xml, &rich_data_parts);
    let mut pkg = XlsxPackage::from_bytes(&input_bytes).expect("read pkg");

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        "Sheet1",
        CellRef::from_a1("A1").unwrap(),
        CellPatch::set_value(CellValue::Number(99.0)),
    );
    pkg.apply_cell_patches(&patches).expect("apply patches");

    let output_bytes = pkg.write_to_bytes().expect("write updated pkg");
    let mut archive = zip::ZipArchive::new(Cursor::new(output_bytes)).expect("read output zip");

    let mut out_metadata = Vec::new();
    archive
        .by_name("xl/metadata.xml")
        .expect("xl/metadata.xml should exist")
        .read_to_end(&mut out_metadata)
        .unwrap();
    assert_eq!(
        out_metadata,
        metadata_xml,
        "expected xl/metadata.xml to be preserved byte-for-byte"
    );

    for (name, expected) in rich_data_parts {
        let mut out_part = Vec::new();
        archive
            .by_name(name)
            .unwrap_or_else(|_| panic!("{name} should exist"))
            .read_to_end(&mut out_part)
            .unwrap();
        assert_eq!(
            out_part,
            expected,
            "expected {name} to be preserved byte-for-byte"
        );
    }

    let mut workbook_rels = String::new();
    archive
        .by_name("xl/_rels/workbook.xml.rels")
        .expect("workbook.xml.rels should exist")
        .read_to_string(&mut workbook_rels)
        .unwrap();

    let doc = roxmltree::Document::parse(&workbook_rels).expect("parse workbook rels");
    let has_metadata_rel = doc.descendants().any(|n| {
        n.is_element()
            && n.tag_name().name() == "Relationship"
            && n.attribute("Target") == Some("metadata.xml")
    });
    assert!(
        has_metadata_rel,
        "expected workbook.xml.rels to keep metadata relationship, got: {workbook_rels}"
    );

    let mut worksheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet1.xml should exist")
        .read_to_string(&mut worksheet_xml)
        .unwrap();
    let doc = roxmltree::Document::parse(&worksheet_xml).expect("parse sheet1.xml");
    let cell = doc
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");
    let v = cell
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .expect("expected <v> in A1");
    assert_eq!(
        v.text().unwrap_or_default(),
        "99",
        "expected A1 value to be updated, got: {worksheet_xml}"
    );
}
