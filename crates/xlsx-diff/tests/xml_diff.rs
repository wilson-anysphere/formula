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
fn worksheet_part_name_is_normalized_for_ordering_rules() {
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

    // Intentionally use un-normalized part names to ensure `NormalizedXml::parse` applies path
    // normalization before checking whether worksheet-specific ordering rules should apply.
    let ax = NormalizedXml::parse(r"\xl\worksheets\..\worksheets\sheet1.xml", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("/xl/worksheets/sheet1.xml", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(diffs.is_empty(), "expected no diffs, got {diffs:#?}");
}

#[test]
fn merge_cells_content_changes_are_not_ignored() {
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
    <mergeCell ref="C3:E4"/>
  </mergeCells>
</worksheet>"#;

    let ax = NormalizedXml::parse("xl/worksheets/sheet1.xml", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/worksheets/sheet1.xml", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(
        diffs.iter().any(|d| d.kind == "attribute_changed" && d.path.contains("@ref")),
        "expected a ref attribute change diff, got {diffs:#?}"
    );
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
fn hyperlinks_content_changes_are_not_ignored() {
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
    <hyperlink ref="A1" r:id="rId9"/>
    <hyperlink ref="B2" r:id="rId2"/>
  </hyperlinks>
</worksheet>"#;

    let ax = NormalizedXml::parse("xl/worksheets/sheet1.xml", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/worksheets/sheet1.xml", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(
        diffs.iter().any(|d| d.kind == "attribute_changed" && d.path.contains("@{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")),
        "expected an r:id attribute change diff, got {diffs:#?}"
    );
}

#[test]
fn hyperlinks_sort_by_r_id_when_ref_is_missing() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <hyperlinks>
    <hyperlink r:id="rId2"/>
    <hyperlink r:id="rId1"/>
  </hyperlinks>
</worksheet>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <hyperlinks>
    <hyperlink r:id="rId1"/>
    <hyperlink r:id="rId2"/>
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
fn data_validations_content_changes_are_not_ignored() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dataValidations count="2">
    <dataValidation type="whole" sqref="B2"/>
    <dataValidation type="list" sqref="A1" allowBlank="1"/>
  </dataValidations>
</worksheet>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dataValidations count="2">
    <dataValidation type="list" sqref="A1" allowBlank="0"/>
    <dataValidation type="whole" sqref="B2"/>
  </dataValidations>
</worksheet>"#;

    let ax = NormalizedXml::parse("xl/worksheets/sheet1.xml", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/worksheets/sheet1.xml", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(
        diffs
            .iter()
            .any(|d| d.kind == "attribute_changed" && d.path.contains("@allowBlank")),
        "expected an allowBlank attribute change diff, got {diffs:#?}"
    );
}

#[test]
fn data_validations_sort_by_sqref_then_type() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dataValidations count="2">
    <dataValidation type="whole" sqref="A1"/>
    <dataValidation type="list" sqref="A1"/>
  </dataValidations>
</worksheet>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dataValidations count="2">
    <dataValidation type="list" sqref="A1"/>
    <dataValidation type="whole" sqref="A1"/>
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

#[test]
fn conditional_formatting_content_changes_are_not_ignored() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <conditionalFormatting sqref="A1">
    <cfRule type="expression" priority="2"/>
    <cfRule type="cellIs" priority="1" operator="equal"/>
  </conditionalFormatting>
</worksheet>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <conditionalFormatting sqref="A1">
    <cfRule type="cellIs" priority="1" operator="notEqual"/>
    <cfRule type="expression" priority="2"/>
  </conditionalFormatting>
</worksheet>"#;

    let ax = NormalizedXml::parse("xl/worksheets/sheet1.xml", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/worksheets/sheet1.xml", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(
        diffs
            .iter()
            .any(|d| d.kind == "attribute_changed" && d.path.contains("@operator")),
        "expected an operator attribute change diff, got {diffs:#?}"
    );
}

#[test]
fn conditional_formatting_rules_sort_by_priority_then_type() {
    let a = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <conditionalFormatting sqref="A1">
    <cfRule type="expression" priority="1"/>
    <cfRule type="cellIs" priority="1"/>
  </conditionalFormatting>
</worksheet>"#;

    let b = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <conditionalFormatting sqref="A1">
    <cfRule type="cellIs" priority="1"/>
    <cfRule type="expression" priority="1"/>
  </conditionalFormatting>
</worksheet>"#;

    let ax = NormalizedXml::parse("xl/worksheets/sheet1.xml", a.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("xl/worksheets/sheet1.xml", b.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert!(diffs.is_empty(), "expected no diffs, got {diffs:#?}");
}

#[test]
fn unicode_text_truncation_is_utf8_safe() {
    // Regression test: `truncate` used to slice `&value[..MAX]` where MAX is a byte
    // offset, which panics if MAX falls in the middle of a multi-byte UTF-8
    // character.
    //
    // Craft a string where the 120th byte is in the middle of `é` (2 bytes):
    // - 119 ASCII bytes ("a" * 119)
    // - followed by "é" (bytes 119..121)
    // so a naive `[..120]` slice would be invalid UTF-8.
    let prefix = format!("{}é", "a".repeat(119));
    let common_tail = "b".repeat(50);
    let expected_text = format!("{prefix}{common_tail}expected");
    let actual_text = format!("{prefix}{common_tail}actual");

    let expected_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root>{expected_text}</root>"#
    );
    let actual_xml = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root>{actual_text}</root>"#
    );

    let ax = NormalizedXml::parse("root.xml", expected_xml.as_bytes()).unwrap();
    let bx = NormalizedXml::parse("root.xml", actual_xml.as_bytes()).unwrap();

    let diffs = diff_xml(&ax, &bx, Severity::Critical);
    assert_eq!(diffs.len(), 1, "expected a single diff, got {diffs:#?}");

    let expected = diffs[0].expected.as_deref().unwrap_or_default();
    let actual = diffs[0].actual.as_deref().unwrap_or_default();
    assert!(
        expected.ends_with('…'),
        "expected truncated string to end with ellipsis, got {expected:?}"
    );
    assert!(
        actual.ends_with('…'),
        "expected truncated string to end with ellipsis, got {actual:?}"
    );
}
