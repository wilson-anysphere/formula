mod support;

use formula_xlsx::XlsxPackage;

use support::rich_data_builder::RichDataXlsxBuilder;

#[test]
fn rich_data_xlsx_builder_emits_rd_rich_value_parts() {
    let worksheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData/>
</worksheet>"#;

    let bytes = RichDataXlsxBuilder::new()
        .add_sheet("Sheet1", worksheet_xml)
        .metadata_xml(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><metadata/>"#)
        .rd_rich_value_xml(
            0,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><rdRichValue/>"#,
        )
        .rd_rich_value_structure_xml(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><rdRichValueStructure/>"#,
        )
        .rd_rich_value_types_xml(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><rdRichValueTypes/>"#,
        )
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
    for expected in [
        "xl/metadata.xml",
        "xl/richData/rdrichvalue.xml",
        "xl/richData/rdrichvaluestructure.xml",
        "xl/richData/rdRichValueTypes.xml",
        "xl/richData/richValueRel.xml",
        "xl/richData/_rels/richValueRel.xml.rels",
        "xl/media/image1.png",
    ] {
        assert!(pkg.part(expected).is_some(), "missing expected part: {expected}");
    }
}

