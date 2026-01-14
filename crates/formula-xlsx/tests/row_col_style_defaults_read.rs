use std::io::{Cursor, Write};

use formula_model::Workbook;
use formula_xlsx::{load_from_bytes, read_workbook_model_from_bytes};

fn build_minimal_xlsx(sheet_xml: &str, styles_xml: &str) -> Vec<u8> {
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
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

    zip.start_file("xl/styles.xml", options).unwrap();
    zip.write_all(styles_xml.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn style_id_for_number_format(workbook: &Workbook, format_code: &str) -> u32 {
    let style_id = workbook
        .styles
        .styles
        .iter()
        .position(|style| style.number_format.as_deref() == Some(format_code))
        .expect("expected custom number format style to exist") as u32;
    assert_ne!(style_id, 0, "expected style id to be non-zero");
    style_id
}

fn assert_row_col_style_defaults(workbook: &Workbook, style_id: u32) {
    let sheet = workbook.sheet_by_name("Sheet1").expect("sheet exists");
    assert_eq!(sheet.col_properties(1).unwrap().style_id, Some(style_id));
    assert_eq!(sheet.row_properties(1).unwrap().style_id, Some(style_id));
}

fn fixture_bytes() -> Vec<u8> {
    // Custom number format to force a non-default style entry.
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="0.00"/>
  </numFmts>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="2" max="2" style="1" customFormat="1"/>
  </cols>
  <sheetData>
    <row r="2" s="1" customFormat="1"/>
  </sheetData>
</worksheet>"#;

    build_minimal_xlsx(sheet_xml, styles_xml)
}

#[test]
fn reads_row_and_col_default_styles_load_from_bytes() {
    let bytes = fixture_bytes();
    let doc = load_from_bytes(&bytes).expect("load_from_bytes");

    let style_id = style_id_for_number_format(&doc.workbook, "0.00");
    assert_row_col_style_defaults(&doc.workbook, style_id);
}

#[test]
fn reads_row_and_col_default_styles_read_workbook_model_from_bytes() {
    let bytes = fixture_bytes();
    let workbook = read_workbook_model_from_bytes(&bytes).expect("read_workbook_model_from_bytes");
    let style_id = style_id_for_number_format(&workbook, "0.00");
    assert_row_col_style_defaults(&workbook, style_id);
}
