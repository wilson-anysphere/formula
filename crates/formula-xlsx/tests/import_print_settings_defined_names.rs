use std::io::{Cursor, Write};

use formula_model::{ColRange, PrintTitles, Range, RowRange};

fn build_minimal_xlsx_with_print_defined_names() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$B$2</definedName>
    <definedName name="_xlnm.Print_Titles" localSheetId="0">Sheet1!$1:$1,Sheet1!$A:$B</definedName>
  </definedNames>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
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

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn imports_print_area_and_titles_into_workbook_model() {
    let bytes = build_minimal_xlsx_with_print_defined_names();

    // Streaming workbook model reader.
    let workbook = formula_xlsx::read_workbook_model_from_bytes(&bytes).unwrap();
    assert_eq!(workbook.print_settings.sheets.len(), 1);
    let settings = &workbook.print_settings.sheets[0];

    assert_eq!(settings.sheet_name, "Sheet1");
    assert_eq!(
        settings.print_area.as_deref(),
        Some(&[Range::from_a1("A1:B2").unwrap()][..])
    );
    assert_eq!(
        settings.print_titles,
        Some(PrintTitles {
            repeat_rows: Some(RowRange { start: 0, end: 0 }),
            repeat_cols: Some(ColRange { start: 0, end: 1 }),
        })
    );

    // Full-fidelity reader should match.
    let doc = formula_xlsx::load_from_bytes(&bytes).unwrap();
    assert_eq!(doc.workbook.print_settings.sheets.len(), 1);
    assert_eq!(
        doc.workbook.print_settings.sheets[0].print_area.as_deref(),
        Some(&[Range::from_a1("A1:B2").unwrap()][..])
    );
    assert_eq!(
        doc.workbook.print_settings.sheets[0].print_titles,
        settings.print_titles
    );
}

