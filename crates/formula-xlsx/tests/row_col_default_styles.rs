use std::io::{Cursor, Read, Write};

use formula_xlsx::{load_from_bytes, read_workbook_model_from_bytes};
use zip::ZipArchive;

fn build_fixture() -> Vec<u8> {
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

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <cols>
    <col min="1" max="1" style="1"/>
  </cols>
  <sheetData>
    <row r="1" s="1" customFormat="1"/>
  </sheetData>
</worksheet>"#;

    // Based on the crate's default styles.xml, but with 2 cellXfs so xf index 1 is valid.
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"#;

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

    zip.start_file("xl/styles.xml", options).unwrap();
    zip.write_all(styles_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn read_sheet_xml(bytes: &[u8]) -> String {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("valid zip");
    let mut xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet1.xml exists")
        .read_to_string(&mut xml)
        .expect("read sheet xml");
    xml
}

#[test]
fn row_and_col_default_styles_roundtrip() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = build_fixture();

    // Reader populates the semantic model.
    let wb = read_workbook_model_from_bytes(&bytes)?;
    let sheet_id = wb.sheets[0].id;
    let sheet = wb.sheet(sheet_id).expect("sheet exists");
    let row_style = sheet.row_properties.get(&0).and_then(|p| p.style_id);
    let col_style = sheet.col_properties.get(&0).and_then(|p| p.style_id);
    assert!(row_style.is_some(), "expected row 1 to have a default style");
    assert_eq!(
        row_style, col_style,
        "expected row 1 and col A to reference the same style id"
    );

    // No-op load/save preserves the SpreadsheetML attributes.
    let doc = load_from_bytes(&bytes)?;
    let out = doc.save_to_vec()?;
    let sheet_xml = read_sheet_xml(&out);
    let doc = roxmltree::Document::parse(&sheet_xml)?;
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    let row = doc
        .descendants()
        .find(|n| n.has_tag_name((ns, "row")) && n.attribute("r") == Some("1"))
        .expect("row 1 exists");
    assert_eq!(row.attribute("s"), Some("1"));
    assert_eq!(row.attribute("customFormat"), Some("1"));

    let col = doc
        .descendants()
        .find(|n| {
            n.has_tag_name((ns, "col")) && n.attribute("min") == Some("1") && n.attribute("max") == Some("1")
        })
        .expect("col A exists");
    assert_eq!(col.attribute("style"), Some("1"));

    Ok(())
}

