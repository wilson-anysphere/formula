use roxmltree::Document;
use std::io::Write;
use zip::write::{FileOptions, ZipWriter};

use formula_model::{parse_range_a1, CfRule, CfRuleKind, CfRuleSchema, Cfvo, CfvoType, Color, DataBarRule};

use formula_xlsx::{load_from_bytes, rule_id_for_index, XlsxPackage};

fn build_minimal_xlsx_with_sheet1(sheet1_xml: &str) -> Vec<u8> {
    let workbook_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"#;

    let workbook_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;

    let cursor = std::io::Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(cursor);
    let options = FileOptions::<()>::default().compression_method(zip::CompressionMethod::Stored);

    zip.start_file("xl/workbook.xml", options)
        .expect("start xl/workbook.xml");
    zip.write_all(workbook_xml.as_bytes())
        .expect("write xl/workbook.xml");

    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .expect("start xl/_rels/workbook.xml.rels");
    zip.write_all(workbook_rels.as_bytes())
        .expect("write xl/_rels/workbook.xml.rels");

    zip.start_file("xl/worksheets/sheet1.xml", options)
        .expect("start xl/worksheets/sheet1.xml");
    zip.write_all(sheet1_xml.as_bytes())
        .expect("write xl/worksheets/sheet1.xml");

    zip.finish().expect("finish zip").into_inner()
}

fn x14_data_bar_rule_missing_id() -> CfRule {
    CfRule {
        schema: CfRuleSchema::X14,
        id: None,
        priority: 1,
        applies_to: vec![parse_range_a1("B1:B3").unwrap()],
        dxf_id: None,
        stop_if_true: false,
        kind: CfRuleKind::DataBar(DataBarRule {
            min: Cfvo {
                type_: CfvoType::AutoMin,
                value: None,
            },
            max: Cfvo {
                type_: CfvoType::AutoMax,
                value: None,
            },
            color: Some(Color::new_argb(0xFF638EC6)),
            min_length: Some(0),
            max_length: Some(100),
            gradient: Some(false),
            negative_fill_color: None,
            axis_color: None,
            direction: None,
        }),
        dependencies: vec![],
    }
}

fn extract_base_and_x14_rule_ids(sheet_xml: &str) -> (String, String) {
    let doc = Document::parse(sheet_xml).expect("valid worksheet xml");
    let main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let x14_ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    let base = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "cfRule"
                && n.tag_name().namespace() == Some(main_ns)
                && n.attribute("type") == Some("dataBar")
        })
        .and_then(|n| n.attribute("id"))
        .expect("expected main cfRule/@id to be present")
        .to_string();

    let ext = doc
        .descendants()
        .find(|n| {
            n.is_element()
                && n.tag_name().name() == "cfRule"
                && n.tag_name().namespace() == Some(x14_ns)
                && n.attribute("type") == Some("dataBar")
        })
        .and_then(|n| n.attribute("id"))
        .expect("expected x14 cfRule/@id to be present")
        .to_string();

    (base, ext)
}

#[test]
fn xlsxdocument_writer_generates_deterministic_x14_cf_rule_ids_when_missing() {
    let sheet1_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;
    let input = build_minimal_xlsx_with_sheet1(sheet1_xml);

    let mut doc = load_from_bytes(&input).expect("load minimal xlsx");
    let sheet_id = doc.workbook.sheets[0].id;
    doc.workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists")
        .set_conditional_formatting(vec![x14_data_bar_rule_missing_id()], vec![]);

    let bytes1 = doc.save_to_vec().expect("save xlsx");
    let bytes2 = doc.save_to_vec().expect("save xlsx again");

    let pkg1 = XlsxPackage::from_bytes(&bytes1).expect("open pkg1");
    let pkg2 = XlsxPackage::from_bytes(&bytes2).expect("open pkg2");

    let sheet1_xml_1 = std::str::from_utf8(pkg1.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let sheet1_xml_2 = std::str::from_utf8(pkg2.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();

    let (base1, x14_1) = extract_base_and_x14_rule_ids(sheet1_xml_1);
    let (base2, x14_2) = extract_base_and_x14_rule_ids(sheet1_xml_2);

    assert_eq!(base1, x14_1, "base/x14 ids must match for linking");
    assert_eq!(base2, x14_2, "base/x14 ids must match for linking");
    assert_eq!(base1, base2, "ids must be deterministic across saves");

    // Writer seeds ids with the worksheet sheetId from workbook.xml (1-based), and uses rule index.
    let expected = rule_id_for_index(1u128, 0);
    assert_eq!(base1, expected);
}

