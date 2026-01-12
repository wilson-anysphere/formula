use std::io::{Cursor, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::{read_workbook_model_from_bytes, XlsxPackage};

fn build_minimal_xlsx() -> Vec<u8> {
    // SpreadsheetML element/attribute namespace prefixes are arbitrary. Ensure we can discover
    // sheets even when workbook.xml uses non-default prefixes (e.g. `x:` for SpreadsheetML and
    // `rel:` for the officeDocument relationships namespace).
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheets>
    <x:sheet name="Sheet1" sheetId="1" rel:id="rId1"/>
  </x:sheets>
</x:workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>7</v></c>
    </row>
  </sheetData>
</worksheet>"#;

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

#[test]
fn reader_discovers_sheets_with_prefixed_workbook_elements_and_relationship_ids(
) -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_minimal_xlsx();

    let workbook = read_workbook_model_from_bytes(&bytes)?;
    assert_eq!(workbook.sheets.len(), 1);
    assert_eq!(workbook.sheets[0].name, "Sheet1");
    assert_eq!(
        workbook.sheets[0].value(CellRef::from_a1("A1")?),
        CellValue::Number(7.0)
    );

    let pkg = XlsxPackage::from_bytes(&bytes)?;
    let sheets = pkg.worksheet_parts()?;
    assert_eq!(sheets.len(), 1);
    assert_eq!(sheets[0].name, "Sheet1");
    assert_eq!(sheets[0].worksheet_part, "xl/worksheets/sheet1.xml");

    Ok(())
}

