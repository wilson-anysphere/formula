use xlsx_diff::{diff_xml, NormalizedXml, Severity};

fn utf16le_with_bom(text: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(2 + text.len() * 2);
    out.extend_from_slice(&[0xFF, 0xFE]);
    for unit in text.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn utf16le_without_bom(text: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(text.len() * 2);
    for unit in text.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn utf16be_without_bom(text: &str) -> Vec<u8> {
    let mut out = Vec::with_capacity(text.len() * 2);
    for unit in text.encode_utf16() {
        out.extend_from_slice(&unit.to_be_bytes());
    }
    out
}

#[test]
fn utf16le_rels_parses_and_diffs_identically_to_utf8() {
    let rels_utf8 = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="t1" Target="a.xml"/>
</Relationships>"#;

    let rels_utf16le = utf16le_with_bom(rels_utf8);

    let utf8 = NormalizedXml::parse("xl/_rels/workbook.xml.rels", rels_utf8.as_bytes()).unwrap();
    let utf16 = NormalizedXml::parse("xl/_rels/workbook.xml.rels", &rels_utf16le).unwrap();

    let diffs = diff_xml(&utf8, &utf16, Severity::Critical);
    assert!(diffs.is_empty(), "expected no diffs, got {diffs:#?}");
}

#[test]
fn utf16le_worksheet_xml_snippet_parses() {
    let worksheet_utf8 = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    let worksheet_utf16le = utf16le_with_bom(worksheet_utf8);
    let parsed = NormalizedXml::parse("xl/worksheets/sheet1.xml", &worksheet_utf16le).unwrap();

    // Spot-check that we got a real document back.
    assert_eq!(
        parsed.root,
        NormalizedXml::parse("xl/worksheets/sheet1.xml", worksheet_utf8.as_bytes())
            .unwrap()
            .root
    );
}

#[test]
fn utf16_without_bom_with_leading_whitespace_is_detected() {
    // Leading whitespace is only valid when the XML declaration is omitted.
    let xml = "\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>";

    let utf8 = NormalizedXml::parse("xl/worksheets/sheet1.xml", xml.as_bytes()).unwrap();
    let utf16le =
        NormalizedXml::parse("xl/worksheets/sheet1.xml", &utf16le_without_bom(xml)).unwrap();
    let utf16be =
        NormalizedXml::parse("xl/worksheets/sheet1.xml", &utf16be_without_bom(xml)).unwrap();

    assert_eq!(utf16le, utf8);
    assert_eq!(utf16be, utf8);
}
