use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{load_from_bytes, XlsxDocument};
use zip::ZipArchive;

fn build_minimal_xlsx(sheet_xml: &str) -> Vec<u8> {
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

    let cursor = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let options = zip::write::FileOptions::<()>::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn set_number(doc: &mut XlsxDocument, a1: &str, value: f64) -> Result<(), Box<dyn std::error::Error>> {
    let sheet_id = doc.workbook.sheets[0].id;
    let cell = CellRef::from_a1(a1)?;
    assert!(doc.set_cell_value(sheet_id, cell, CellValue::Number(value)));
    Ok(())
}

fn read_sheet_xml(bytes: &[u8]) -> Result<String, Box<dyn std::error::Error>> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")?
        .read_to_string(&mut sheet_xml)?;
    Ok(sheet_xml)
}

#[test]
fn writer_expands_prefixed_sheetdata_and_writes_prefixed_cells(
) -> Result<(), Box<dyn std::error::Error>> {
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetData/>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml);
    let mut doc = load_from_bytes(&bytes)?;
    set_number(&mut doc, "A1", 9.0)?;

    let out = doc.save_to_vec()?;
    let sheet_xml = read_sheet_xml(&out)?;

    // Ensure the output is well-formed XML (mismatched prefixes used to break this).
    roxmltree::Document::parse(&sheet_xml)?;

    assert!(
        sheet_xml.contains("<x:sheetData>") && sheet_xml.contains("</x:sheetData>"),
        "expected <x:sheetData> expansion in output XML"
    );
    assert!(
        sheet_xml.contains("<x:row") && sheet_xml.contains("</x:row>"),
        "expected prefixed rows in output XML"
    );
    assert!(
        sheet_xml.contains("<x:c r=\"A1\""),
        "expected prefixed cell tag in output XML"
    );
    assert!(
        sheet_xml.contains("<x:v>9</x:v>"),
        "expected prefixed value element in output XML"
    );
    assert!(
        sheet_xml.contains("<x:dimension ref=\"A1\""),
        "expected prefixed dimension element in output XML"
    );

    Ok(())
}

#[test]
fn writer_inserts_missing_sheetdata_with_prefix() -> Result<(), Box<dyn std::error::Error>> {
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetViews/>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml);
    let mut doc = load_from_bytes(&bytes)?;
    set_number(&mut doc, "A1", 1.0)?;

    let out = doc.save_to_vec()?;
    let sheet_xml = read_sheet_xml(&out)?;

    roxmltree::Document::parse(&sheet_xml)?;
    assert!(
        sheet_xml.contains("<x:sheetData>"),
        "expected writer to insert prefixed <x:sheetData> when missing"
    );
    assert!(
        sheet_xml.contains("<x:c r=\"A1\""),
        "expected inserted sheetData to contain prefixed cells"
    );

    Ok(())
}

#[test]
fn writer_updates_prefixed_dimension_element_in_place() -> Result<(), Box<dyn std::error::Error>> {
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:dimension ref="A1"/>
  <x:sheetData/>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml);
    let mut doc = load_from_bytes(&bytes)?;
    set_number(&mut doc, "A1", 1.0)?;
    set_number(&mut doc, "C3", 2.0)?;

    let out = doc.save_to_vec()?;
    let sheet_xml = read_sheet_xml(&out)?;
    let xml_doc = roxmltree::Document::parse(&sheet_xml)?;

    let dimension_count = xml_doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "dimension")
        .count();
    assert_eq!(
        dimension_count, 1,
        "expected worksheet to contain exactly one dimension element"
    );
    assert!(
        sheet_xml.contains("<x:dimension ref=\"A1:C3\""),
        "expected prefixed <x:dimension> ref to update to A1:C3"
    );
    assert!(
        !sheet_xml.contains("<dimension ref=\""),
        "should not introduce an unprefixed <dimension> element"
    );

    Ok(())
}

