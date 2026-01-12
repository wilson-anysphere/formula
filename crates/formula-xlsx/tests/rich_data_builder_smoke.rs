mod support;

use formula_xlsx::{load_from_bytes, XlsxPackage};

use support::rich_data_builder::RichDataXlsxBuilder;

#[test]
fn rich_data_xlsx_builder_emits_expected_parts() {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c></row>
  </sheetData>
</worksheet>"#;

    let bytes = RichDataXlsxBuilder::new()
        .add_sheet("Sheet1", worksheet_xml)
        .metadata_xml(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><metadata/>"#)
        .rich_value_xml(0, r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><rv0/>"#)
        .rich_value_xml(1, r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><rv/>"#)
        .rich_value_rel_xml(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><rvRel/>"#,
        )
        .rich_value_rel_rels_xml(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>"#,
        )
        .media_part("xl/media/image1.png", b"not-a-real-png")
        .build_bytes();

    let pkg = XlsxPackage::from_bytes(&bytes).expect("read zip");
    assert!(pkg.part("xl/workbook.xml").is_some());
    assert!(pkg.part("xl/_rels/workbook.xml.rels").is_some());
    assert!(pkg.part("xl/worksheets/sheet1.xml").is_some());

    assert!(pkg.part("xl/metadata.xml").is_some());
    assert!(pkg.part("xl/richData/richValue.xml").is_some());
    assert!(pkg.part("xl/richData/richValue1.xml").is_some());
    assert!(pkg.part("xl/richData/richValueRel.xml").is_some());
    assert!(pkg.part("xl/richData/_rels/richValueRel.xml.rels").is_some());
    assert!(pkg.part("xl/media/image1.png").is_some());

    // Ensure the bytes are loadable through the higher-level reader as well.
    let doc = load_from_bytes(&bytes).expect("load");
    assert!(doc.workbook.sheet_by_name("Sheet1").is_some());
}
