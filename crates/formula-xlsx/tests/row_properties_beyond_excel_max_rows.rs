use std::io::{Cursor, Write};

use formula_xlsx::read_workbook_model_from_bytes;

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

#[test]
fn reader_applies_row_properties_beyond_excel_max_rows() -> Result<(), Box<dyn std::error::Error>> {
    // Excel's UI caps rows at 1,048,576, but the underlying OOXML schema uses unsigned integers.
    // We should preserve row-level properties on read even when r= exceeds Excel's max.
    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="2000000" ht="15" customHeight="1" hidden="1"/>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml);
    let workbook = read_workbook_model_from_bytes(&bytes)?;
    assert_eq!(workbook.sheets.len(), 1);
    let sheet = &workbook.sheets[0];

    let row_0based = 2_000_000u32 - 1;
    let props = sheet
        .row_properties(row_0based)
        .expect("row properties should be preserved beyond Excel max rows");
    assert_eq!(props.height, Some(15.0));
    assert!(sheet.is_row_hidden_effective(2_000_000));
    assert!(props.hidden);
    assert!(sheet.row_count >= 2_000_000);

    Ok(())
}

