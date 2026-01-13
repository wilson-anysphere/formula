use xlsx_diff::{diff_xml, NormalizedXml, Severity};

#[test]
fn relationships_order_is_ignored() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="t2" Target="b.xml"/>
  <Relationship Id="rId1" Type="t1" Target="a.xml"/>
</Relationships>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="t1" Target="a.xml"/>
  <Relationship Id="rId2" Type="t2" Target="b.xml"/>
</Relationships>"#;

    let ax = NormalizedXml::parse("xl/_rels/workbook.xml.rels", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/_rels/workbook.xml.rels", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(diffs.is_empty(), "expected no diffs, got {diffs:#?}");
}

#[test]
fn relationships_are_keyed_by_id_for_paths() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="t1" Target="a.xml"/>
</Relationships>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="t1" Target="b.xml"/>
</Relationships>"#;

    let ax = NormalizedXml::parse("xl/_rels/workbook.xml.rels", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/_rels/workbook.xml.rels", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert_eq!(diffs.len(), 1);
    assert_eq!(diffs[0].kind, "attribute_changed");
    assert!(
        diffs[0].path.contains("Relationship[@Id=\"rId1\"]@Target"),
        "unexpected path: {}",
        diffs[0].path
    );
}

#[test]
fn sheetdata_rows_are_keyed_by_r() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
    <row r="2"><c r="A2"><v>2</v></c></row>
  </sheetData>
</worksheet>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    let ax = NormalizedXml::parse("xl/worksheets/sheet1.xml", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/worksheets/sheet1.xml", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(
        diffs
            .iter()
            .any(|d| d.kind == "child_missing" && d.path.contains("row[@r=\"2\"]")),
        "expected a missing row[2] diff, got {diffs:#?}"
    );
}

#[test]
fn cols_are_sorted_by_min_max() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="2" max="2" width="10"/>
    <col min="1" max="1" width="8"/>
  </cols>
</worksheet>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols>
    <col min="1" max="1" width="8"/>
    <col min="2" max="2" width="10"/>
  </cols>
</worksheet>"#;

    let ax = NormalizedXml::parse("xl/worksheets/sheet1.xml", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/worksheets/sheet1.xml", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(diffs.is_empty(), "expected no diffs, got {diffs:#?}");
}

#[test]
fn defined_names_are_sorted_by_name() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <definedNames>
    <definedName name="ZedName">Sheet1!$A$1</definedName>
    <definedName name="MyRange">Sheet1!$A$1:$A$1</definedName>
  </definedNames>
</workbook>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <definedNames>
    <definedName name="MyRange">Sheet1!$A$1:$A$1</definedName>
    <definedName name="ZedName">Sheet1!$A$1</definedName>
  </definedNames>
</workbook>"#;

    let ax = NormalizedXml::parse("xl/workbook.xml", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/workbook.xml", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(diffs.is_empty(), "expected no diffs, got {diffs:#?}");
}

#[test]
fn merge_cells_order_is_ignored() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <mergeCells count="2">
    <mergeCell ref="C3:D4"/>
    <mergeCell ref="A1:B2"/>
  </mergeCells>
</worksheet>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <mergeCells count="2">
    <mergeCell ref="A1:B2"/>
    <mergeCell ref="C3:D4"/>
  </mergeCells>
</worksheet>"#;

    let ax = NormalizedXml::parse("xl/worksheets/sheet1.xml", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/worksheets/sheet1.xml", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(diffs.is_empty(), "expected no diffs, got {diffs:#?}");
}

#[test]
fn hyperlinks_order_is_ignored() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <hyperlinks>
    <hyperlink ref="B2" r:id="rId2"/>
    <hyperlink ref="A1" r:id="rId1"/>
  </hyperlinks>
</worksheet>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <hyperlinks>
    <hyperlink ref="A1" r:id="rId1"/>
    <hyperlink ref="B2" r:id="rId2"/>
  </hyperlinks>
</worksheet>"#;

    let ax = NormalizedXml::parse("xl/worksheets/sheet1.xml", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/worksheets/sheet1.xml", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(diffs.is_empty(), "expected no diffs, got {diffs:#?}");
}

#[test]
fn data_validations_order_is_ignored() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dataValidations count="2">
    <dataValidation type="whole" sqref="B2"/>
    <dataValidation type="list" sqref="A1"/>
  </dataValidations>
</worksheet>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dataValidations count="2">
    <dataValidation type="list" sqref="A1"/>
    <dataValidation type="whole" sqref="B2"/>
  </dataValidations>
</worksheet>"#;

    let ax = NormalizedXml::parse("xl/worksheets/sheet1.xml", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/worksheets/sheet1.xml", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(diffs.is_empty(), "expected no diffs, got {diffs:#?}");
}

#[test]
fn conditional_formatting_rules_are_sorted_by_priority() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <conditionalFormatting sqref="A1">
    <cfRule type="expression" priority="2"/>
    <cfRule type="cellIs" priority="1"/>
  </conditionalFormatting>
</worksheet>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <conditionalFormatting sqref="A1">
    <cfRule type="cellIs" priority="1"/>
    <cfRule type="expression" priority="2"/>
  </conditionalFormatting>
</worksheet>"#;

    let ax = NormalizedXml::parse("xl/worksheets/sheet1.xml", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/worksheets/sheet1.xml", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(diffs.is_empty(), "expected no diffs, got {diffs:#?}");
}
