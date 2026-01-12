use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
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

fn zip_part(zip_bytes: &[u8], name: &str) -> Vec<u8> {
    let cursor = Cursor::new(zip_bytes);
    let mut archive = ZipArchive::new(cursor).expect("open zip");
    let mut file = archive.by_name(name).expect("part exists");
    let mut buf = Vec::new();
    file.read_to_end(&mut buf).expect("read part");
    buf
}

#[test]
fn vm_cm_are_recorded_in_cell_meta_and_survive_roundtrip(
) -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="e" vm="3" cm="4"><v>#VALUE!</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);
    let mut doc = load_from_bytes(&bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;

    let a1 = CellRef::from_a1("A1")?;
    let meta = doc.cell_meta(sheet_id, a1).expect("expected CellMeta for A1");
    assert_eq!(meta.vm.as_deref(), Some("3"));
    assert_eq!(meta.cm.as_deref(), Some("4"));

    // Ensure metadata survives clearing the cell record.
    assert!(doc.clear_cell(sheet_id, a1));
    let meta_after_clear = doc
        .cell_meta(sheet_id, a1)
        .expect("expected CellMeta for A1 after clear_cell");
    assert_eq!(meta_after_clear.vm.as_deref(), Some("3"));
    assert_eq!(meta_after_clear.cm.as_deref(), Some("4"));

    assert!(doc.set_cell_value(sheet_id, a1, CellValue::Number(99.0)));

    let saved = doc.save_to_vec()?;
    let out_xml_bytes = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let out_xml = std::str::from_utf8(&out_xml_bytes)?;
    let parsed = roxmltree::Document::parse(out_xml)?;

    let cell_a1 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("A1 cell exists");
    // Editing the cell away from rich-value placeholder semantics should drop `vm` to avoid
    // leaving a dangling value-metadata pointer.
    assert_eq!(cell_a1.attribute("vm"), None);
    assert_eq!(cell_a1.attribute("cm"), Some("4"));

    let v_a1 = cell_a1
        .children()
        .find(|n| n.is_element() && n.tag_name().name() == "v")
        .and_then(|n| n.text())
        .unwrap_or_default();
    assert_eq!(v_a1, "99");

    Ok(())
}
