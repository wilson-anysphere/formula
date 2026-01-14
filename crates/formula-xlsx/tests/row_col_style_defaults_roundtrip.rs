use std::io::{Cursor, Read, Write};

use formula_model::{Style, Workbook};
use formula_xlsx::{load_from_bytes, write, XlsxDocument};
use zip::ZipArchive;

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

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet_xml.as_bytes()).unwrap();

    zip.start_file("xl/styles.xml", options).unwrap();
    zip.write_all(styles_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

fn read_zip_part(bytes: &[u8], name: &str) -> String {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut file = archive.by_name(name).unwrap();
    let mut out = String::new();
    file.read_to_string(&mut out).unwrap();
    out
}

#[test]
fn noop_roundtrip_preserves_row_and_col_default_styles() -> Result<(), Box<dyn std::error::Error>>
{
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
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cols>
    <col min="1" max="1" style="1" customFormat="1"/>
  </cols>
  <sheetData>
    <row r="1" s="1" customFormat="1"/>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml, styles_xml);
    let doc = load_from_bytes(&bytes)?;
    let out = write::write_to_vec(&doc)?;

    let out_sheet_xml = read_zip_part(&out, "xl/worksheets/sheet1.xml");
    assert!(
        out_sheet_xml.contains(r#"style="1" customFormat="1""#),
        "expected column default style to survive, got:\n{out_sheet_xml}"
    );
    assert!(
        out_sheet_xml.contains(r#"s="1" customFormat="1""#),
        "expected row default style to survive, got:\n{out_sheet_xml}"
    );

    Ok(())
}

#[test]
fn patch_writer_emits_row_default_style_when_xf_index_is_zero(
) -> Result<(), Box<dyn std::error::Error>> {
    // Some producers place custom xfs at index 0. Ensure we still emit `row/@s="0"` when the
    // model references that non-default style.
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
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" s="0"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml, styles_xml);
    let mut doc = load_from_bytes(&bytes)?;

    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).expect("sheet exists");
    let xf0_style_id = sheet
        .cell(formula_model::CellRef::from_a1("A1")?)
        .expect("A1 exists")
        .style_id;
    assert_ne!(xf0_style_id, 0, "expected xf index 0 to map to a non-default style");

    // Row properties are 0-based; apply to row 2 and force the patch writer to synthesize that row.
    sheet.set_row_style_id(1, Some(xf0_style_id));
    sheet.set_value(formula_model::CellRef::from_a1("A2")?, formula_model::CellValue::Number(2.0));

    let out = write::write_to_vec(&doc)?;
    let out_sheet_xml = read_zip_part(&out, "xl/worksheets/sheet1.xml");
    let parsed = roxmltree::Document::parse(&out_sheet_xml)?;

    let row2 = parsed
        .descendants()
        .find(|n| n.is_element() && n.tag_name().name() == "row" && n.attribute("r") == Some("2"))
        .expect("row 2 exists");
    assert_eq!(row2.attribute("s"), Some("0"));
    assert_eq!(row2.attribute("customFormat"), Some("1"));

    Ok(())
}

#[test]
fn patch_writer_emits_col_default_style_when_xf_index_is_zero(
) -> Result<(), Box<dyn std::error::Error>> {
    // Some producers place custom xfs at index 0. Ensure we still emit `col/@style="0"` when the
    // model references that non-default style.
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
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1"><c r="A1" s="0"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml, styles_xml);
    let mut doc = load_from_bytes(&bytes)?;

    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet_mut(sheet_id).expect("sheet exists");
    let xf0_style_id = sheet
        .cell(formula_model::CellRef::from_a1("A1")?)
        .expect("A1 exists")
        .style_id;
    assert_ne!(xf0_style_id, 0, "expected xf index 0 to map to a non-default style");

    // Column properties are 0-based; apply to column B.
    sheet.set_col_style_id(1, Some(xf0_style_id));

    let out = write::write_to_vec(&doc)?;
    let out_sheet_xml = read_zip_part(&out, "xl/worksheets/sheet1.xml");
    let parsed = roxmltree::Document::parse(&out_sheet_xml)?;

    let col_b = parsed
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "col"
                && n.attribute("min") == Some("2")
                && n.attribute("max") == Some("2")
        })
        .expect("col B exists");
    assert_eq!(col_b.attribute("style"), Some("0"));
    assert_eq!(col_b.attribute("customFormat"), Some("1"));

    Ok(())
}

#[test]
fn noop_roundtrip_preserves_col_default_style_when_xf_index_is_zero_and_unknown_col_attrs(
) -> Result<(), Box<dyn std::error::Error>> {
    // If a producer stores a non-default xf at index 0, we should treat `style="0"` as a real
    // override and avoid rewriting `<cols>` on a no-op save. This also ensures unknown `<col>`
    // attributes (like `bestFit`) are preserved byte-for-byte.
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
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cols>
    <col min="2" max="2" style="0" customFormat="1" bestFit="1"/>
  </cols>
  <sheetData/>
</worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml, styles_xml);
    let doc = load_from_bytes(&bytes)?;
    let out = write::write_to_vec(&doc)?;

    let out_sheet_xml = read_zip_part(&out, "xl/worksheets/sheet1.xml");
    assert!(
        out_sheet_xml.contains(r#"bestFit="1""#),
        "expected unknown col attribute to be preserved, got:\n{out_sheet_xml}"
    );
    assert!(
        out_sheet_xml.contains(r#"style="0" customFormat="1""#),
        "expected col style xf0 override to be preserved, got:\n{out_sheet_xml}"
    );

    Ok(())
}

#[test]
fn noop_roundtrip_preserves_row_default_style_when_xf_index_is_zero_and_unknown_row_attrs(
) -> Result<(), Box<dyn std::error::Error>> {
    // Mirror the column test above: for some producers, `cellXfs[0]` is non-default.
    // Ensure we treat `row s="0"` as a real override and avoid rewriting `<row>` on a no-op save,
    // preserving unknown row attributes like `spans`.
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
    <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
  <dxfs count="0"/>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>"#;

    let sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="2" s="0" customFormat="1" spans="1:1"/>
  </sheetData>
</worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml, styles_xml);
    let doc = load_from_bytes(&bytes)?;
    let out = write::write_to_vec(&doc)?;

    let out_sheet_xml = read_zip_part(&out, "xl/worksheets/sheet1.xml");
    assert!(
        out_sheet_xml.contains(r#"s="0" customFormat="1""#),
        "expected row style xf0 override to be preserved, got:\n{out_sheet_xml}"
    );
    assert!(
        out_sheet_xml.contains(r#"spans="1:1""#),
        "expected unknown row attribute spans to be preserved, got:\n{out_sheet_xml}"
    );

    Ok(())
}

#[test]
fn new_document_emits_row_and_col_default_styles() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1")?;

    // Create a style distinct from the default so it is guaranteed to allocate a non-zero xf.
    let style_id = workbook.styles.intern(Style {
        number_format: Some("0.00".to_string()),
        ..Default::default()
    });

    let sheet = workbook.sheet_mut(sheet_id).unwrap();
    sheet.set_row_style_id(0, Some(style_id));
    sheet.set_col_style_id(0, Some(style_id));

    let doc = XlsxDocument::new(workbook);
    let out = write::write_to_vec(&doc)?;

    let out_sheet_xml = read_zip_part(&out, "xl/worksheets/sheet1.xml");
    assert!(
        out_sheet_xml.contains(r#"style="1" customFormat="1""#),
        "expected column default style to be emitted, got:\n{out_sheet_xml}"
    );
    assert!(
        out_sheet_xml.contains(r#"s="1" customFormat="1""#),
        "expected row default style to be emitted, got:\n{out_sheet_xml}"
    );

    Ok(())
}
