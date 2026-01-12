use std::io::{Cursor, Write};

use formula_model::{CellRef, CellValue};
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn build_minimal_xlsx() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    // SpreadsheetML/OPC namespace prefixes are arbitrary. Ensure we can still resolve workbook
    // relationships when `workbook.xml.rels` uses prefixed `<Relationship>` elements.
    //
    // Use a non-default shared strings target so relationship resolution is required for correct
    // value decoding.
    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pr:Relationships xmlns:pr="http://schemas.openxmlformats.org/package/2006/relationships">
  <pr:Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <pr:Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="customSharedStrings.xml"/>
</pr:Relationships>"#;

    let shared_strings_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Hello</t></si>
</sst>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
    </row>
  </sheetData>
</worksheet>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/customSharedStrings.xml", options).unwrap();
    zip.write_all(shared_strings_xml.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn reader_resolves_prefixed_relationship_elements_in_workbook_rels(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_xlsx();

    let workbook = formula_xlsx::read_workbook_model_from_bytes(&bytes)?;
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
    assert_eq!(
        workbook.sheets[0].value(CellRef::from_a1("A1")?),
        CellValue::String("Hello".to_string())
    );

    // Ensure the full reader path (`load_from_bytes`) also resolves workbook relationships.
    let doc = formula_xlsx::load_from_bytes(&bytes)?;
    assert_eq!(doc.workbook.sheets.len(), 1);
    assert_eq!(doc.workbook.sheets[0].name, "Sheet1");
    assert_eq!(
        doc.workbook.sheets[0].value(CellRef::from_a1("A1")?),
        CellValue::String("Hello".to_string())
    );

    Ok(())
}
