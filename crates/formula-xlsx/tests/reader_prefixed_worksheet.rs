use std::io::{Cursor, Write};

use formula_model::{CellRef, CellValue};
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

#[test]
fn reader_parses_prefixed_worksheet_elements_and_inline_strings(
) -> Result<(), Box<dyn std::error::Error>> {
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
<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <x:sheetViews>
    <x:sheetView workbookViewId="0" zoomScale="125">
      <x:pane state="frozen" xSplit="1" ySplit="2"/>
    </x:sheetView>
  </x:sheetViews>
  <x:cols>
    <x:col min="1" max="1" width="12" customWidth="1"/>
    <x:col min="2" max="2" hidden="1"/>
  </x:cols>
  <x:sheetData>
    <x:row r="1" ht="18" customHeight="1">
      <x:c r="A1"><x:v>42</x:v></x:c>
      <x:c r="B1" t="inlineStr">
        <x:is>
          <x:r><x:t xml:space="preserve">Hello </x:t></x:r>
          <x:r><x:rPr><x:b/></x:rPr><x:t>World</x:t></x:r>
        </x:is>
      </x:c>
      <x:c r="C1"><x:f>A1+1</x:f><x:v>43</x:v></x:c>
      <x:c r="D1" s="0" x:s="1"><x:v>7</x:v></x:c>
      <x:c r="E1" s="1"><x:v>8</x:v></x:c>
    </x:row>
  </x:sheetData>
</x:worksheet>"#;

    let bytes = build_minimal_xlsx(sheet_xml, styles_xml);

    let full = load_from_bytes(&bytes)?.workbook;
    let fast = read_workbook_model_from_bytes(&bytes)?;

    for workbook in [&full, &fast] {
        assert_eq!(workbook.sheets.len(), 1);
        let sheet = &workbook.sheets[0];

        assert_eq!(sheet.value(CellRef::from_a1("A1")?), CellValue::Number(42.0));
        match sheet.value(CellRef::from_a1("B1")?) {
            CellValue::RichText(rich) => assert_eq!(rich.text, "Hello World"),
            other => panic!("expected rich text inline string, got {other:?}"),
        }
        assert_eq!(sheet.value(CellRef::from_a1("C1")?), CellValue::Number(43.0));
        assert_eq!(sheet.formula(CellRef::from_a1("C1")?), Some("A1+1"));

        // Worksheet metadata parsed from prefixed elements.
        assert_eq!(sheet.zoom, 1.25);
        assert_eq!(sheet.frozen_cols, 1);
        assert_eq!(sheet.frozen_rows, 2);
        assert_eq!(
            sheet.col_properties(0).and_then(|p| p.width),
            Some(12.0)
        );
        assert_eq!(sheet.col_properties(1).map(|p| p.hidden), Some(true));
        assert_eq!(
            sheet.row_properties(0).and_then(|p| p.height),
            Some(18.0)
        );

        // Only the unprefixed `s` attribute should be treated as the style index.
        assert_eq!(sheet.cell(CellRef::from_a1("D1")?).unwrap().style_id, 0);
        assert_ne!(sheet.cell(CellRef::from_a1("E1")?).unwrap().style_id, 0);
    }

    Ok(())
}
