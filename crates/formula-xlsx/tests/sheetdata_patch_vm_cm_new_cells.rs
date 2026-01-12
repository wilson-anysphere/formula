use std::io::{Cursor, Read, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
use zip::{ZipArchive, ZipWriter};

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
    let mut zip = ZipWriter::new(cursor);
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
fn sheetdata_patch_emits_vm_cm_for_new_cells_from_meta() -> Result<(), Box<dyn std::error::Error>>
{
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let bytes = build_minimal_xlsx(worksheet_xml);
    let mut doc = load_from_bytes(&bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let cell_ref = CellRef::from_a1("A1")?;

    doc.set_cell_value(
        sheet_id,
        cell_ref,
        CellValue::String("MSFT".to_string()),
    );
    let meta = doc.meta.cell_meta.entry((sheet_id, cell_ref)).or_default();
    meta.vm = Some("1".to_string());
    meta.cm = Some("2".to_string());

    let saved = doc.save_to_vec()?;
    let sheet_xml_bytes = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let sheet_xml = std::str::from_utf8(&sheet_xml_bytes)?;
    let parsed = roxmltree::Document::parse(sheet_xml)?;
    let cell = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "c" && n.attribute("r") == Some("A1"))
        .expect("expected A1 cell");
    assert_eq!(cell.attribute("vm"), Some("1"));
    assert_eq!(cell.attribute("cm"), Some("2"));

    Ok(())
}
