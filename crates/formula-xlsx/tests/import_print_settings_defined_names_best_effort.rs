use std::io::{Cursor, Write};

fn build_minimal_xlsx_with_invalid_print_defined_name() -> Vec<u8> {
    // Deliberately malformed Print_Area reference (missing `!` and invalid A1).
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">NOT_A_REFERENCE</definedName>
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
fn ignores_malformed_print_defined_names_in_workbook_load() {
    let bytes = build_minimal_xlsx_with_invalid_print_defined_name();

    // Best-effort: invalid print defined names should not fail workbook load.
    let workbook = formula_xlsx::read_workbook_model_from_bytes(&bytes).unwrap();
    assert!(workbook.print_settings.sheets.is_empty());

    // Full-fidelity reader should also succeed and ignore invalid refs.
    let doc = formula_xlsx::load_from_bytes(&bytes).unwrap();
    assert!(doc.workbook.print_settings.sheets.is_empty());
}

