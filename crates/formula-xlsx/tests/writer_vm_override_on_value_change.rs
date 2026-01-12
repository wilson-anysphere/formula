use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
use zip::{write::FileOptions, CompressionMethod, ZipArchive, ZipWriter};

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

fn build_minimal_xlsx_with_vm_cell() -> Vec<u8> {
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

    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" vm="1" cm="2"><v>1</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Stored);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet1_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn writer_honors_vm_override_when_value_changes() -> Result<(), Box<dyn std::error::Error>> {
    let input = build_minimal_xlsx_with_vm_cell();
    let mut doc = load_from_bytes(&input)?;

    let sheet_id = doc
        .workbook
        .sheets
        .first()
        .map(|s| s.id)
        .ok_or("expected Sheet1")?;
    let cell = CellRef::from_a1("A1")?;

    // Change the value to something non-placeholder so the writer would normally drop the
    // original vm pointer.
    doc.set_cell_value(sheet_id, cell, CellValue::Number(2.0));

    // Explicitly override `vm` via metadata; this should win over the drop semantics.
    let meta = doc.xlsx_meta_mut().cell_meta.entry((sheet_id, cell)).or_default();
    meta.vm = Some("99".to_string());

    let saved = doc.save_to_vec()?;
    let sheet_xml = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let sheet_xml_str = std::str::from_utf8(&sheet_xml)?;
    let parsed = roxmltree::Document::parse(sheet_xml_str)?;

    let cell_node = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .ok_or("expected A1 cell")?;

    assert_eq!(
        cell_node.attribute("vm"),
        Some("99"),
        "expected meta override to win, got: {sheet_xml_str}"
    );
    assert_eq!(
        cell_node.attribute("cm"),
        Some("2"),
        "expected cm attribute to be preserved, got: {sheet_xml_str}"
    );

    let v = cell_node
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(
        v, "2",
        "expected patch to update cached value (sanity check), got: {sheet_xml_str}"
    );

    Ok(())
}

