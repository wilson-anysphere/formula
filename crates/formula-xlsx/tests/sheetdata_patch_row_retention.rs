use std::io::{Cursor, Read, Write};

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
fn clearing_row_hidden_drops_empty_row_element() -> Result<(), Box<dyn std::error::Error>> {
    // Sheet XML contains an otherwise-empty row that only exists to carry a managed attribute
    // (`hidden="1"`). After clearing the row property in the model, the patch writer should
    // drop the now-truly-empty `<row r="2"/>` element entirely.
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2" hidden="1"/>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml);
    let mut doc = load_from_bytes(&bytes)?;

    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).expect("sheet exists");

    assert!(
        sheet.row_properties(1).is_some(),
        "expected row 2 to be hidden on load"
    );
    sheet.set_row_hidden(1, false);
    assert!(
        sheet.row_properties(1).is_none(),
        "expected clearing row hidden to remove the row_properties entry"
    );

    let saved = doc.save_to_vec()?;
    let xml_bytes = zip_part(&saved, "xl/worksheets/sheet1.xml");
    let xml = std::str::from_utf8(&xml_bytes)?;
    let parsed = roxmltree::Document::parse(xml)?;

    assert!(
        parsed
            .descendants()
            .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("2"))
            .is_none(),
        "expected row 2 to be removed after clearing hidden, got: {xml}"
    );

    Ok(())
}

