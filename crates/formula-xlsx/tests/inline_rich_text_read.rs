use std::io::{Cursor, Write};

use formula_model::{CellRef, CellValue};
use formula_xlsx::load_from_bytes;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipWriter};

fn build_fixture_xlsx(worksheet_xml: &str) -> Vec<u8> {
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

    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"#;

    let cursor = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(root_rels.as_bytes()).unwrap();

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(workbook_xml.as_bytes()).unwrap();

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(workbook_rels.as_bytes()).unwrap();

    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(worksheet_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

#[test]
fn load_inline_rich_text_parses_runs() -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <r>
            <rPr><b/></rPr>
            <t>Bold</t>
          </r>
          <r>
            <t>Plain</t>
          </r>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

    let bytes = build_fixture_xlsx(worksheet_xml);
    let doc = load_from_bytes(&bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).expect("sheet exists");

    let value = sheet.value(CellRef::from_a1("A1")?);
    let rich = match value {
        CellValue::RichText(rich) => rich,
        other => panic!("expected CellValue::RichText, got {other:?}"),
    };

    assert_eq!(rich.text, "BoldPlain");
    assert_eq!(rich.runs.len(), 2);
    assert_eq!(rich.runs[0].style.bold, Some(true));

    Ok(())
}

#[test]
fn load_inline_rich_text_ignores_phonetic_text() -> Result<(), Box<dyn std::error::Error>> {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <r>
            <rPr><b/></rPr>
            <t>Base</t>
          </r>
          <phoneticPr fontId="0" type="noConversion"/>
          <rPh sb="0" eb="4"><t>PHO</t></rPh>
          <r>
            <t>Text</t>
          </r>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>"#;

    let bytes = build_fixture_xlsx(worksheet_xml);
    let doc = load_from_bytes(&bytes)?;
    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc.workbook.sheet(sheet_id).expect("sheet exists");

    let value = sheet.value(CellRef::from_a1("A1")?);
    let rich = match value {
        CellValue::RichText(rich) => rich,
        other => panic!("expected CellValue::RichText, got {other:?}"),
    };

    assert_eq!(rich.text, "BaseText");
    assert!(!rich.text.contains("PHO"));
    assert_eq!(rich.runs.len(), 2);

    Ok(())
}
