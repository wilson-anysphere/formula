use std::io::{Cursor, Write};

use formula_model::{PrintTitles, Range};

fn build_minimal_xlsx_with_multi_sheet_print_defined_names() -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    <sheet name="Sheet2" sheetId="2" r:id="rId2"/>
  </sheets>
  <definedNames>
    <!-- Intentionally reversed order (Sheet2 first) to assert deterministic sorting by sheet order. -->
    <definedName name="_xlnm.Print_Titles" localSheetId="1">Sheet2!$1:$1</definedName>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$A$1:$A$2,Sheet1!$C$1:$C$2</definedName>
  </definedNames>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
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

    zip.start_file("xl/worksheets/sheet2.xml", options)
        .unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn imports_print_settings_multi_sheet_and_sorts_by_sheet_order() {
    let bytes = build_minimal_xlsx_with_multi_sheet_print_defined_names();
    let workbook = formula_xlsx::read_workbook_model_from_bytes(&bytes).unwrap();

    assert_eq!(workbook.sheets.len(), 2);
    assert_eq!(workbook.print_settings.sheets.len(), 2);

    // Deterministic ordering should match workbook sheet order, not defined-name order.
    assert_eq!(workbook.print_settings.sheets[0].sheet_name, "Sheet1");
    assert_eq!(workbook.print_settings.sheets[1].sheet_name, "Sheet2");

    assert_eq!(
        workbook.print_settings.sheets[0].print_area,
        Some(vec![
            Range::from_a1("A1:A2").unwrap(),
            Range::from_a1("C1:C2").unwrap(),
        ])
    );
    assert_eq!(workbook.print_settings.sheets[0].print_titles, None);

    assert_eq!(
        workbook.print_settings.sheets[1].print_titles,
        Some(PrintTitles {
            repeat_rows: Some(formula_model::RowRange { start: 0, end: 0 }),
            repeat_cols: None,
        })
    );
    assert_eq!(workbook.print_settings.sheets[1].print_area, None);

    // Full-fidelity reader should match.
    let doc = formula_xlsx::load_from_bytes(&bytes).unwrap();
    assert_eq!(doc.workbook.print_settings, workbook.print_settings);
}
