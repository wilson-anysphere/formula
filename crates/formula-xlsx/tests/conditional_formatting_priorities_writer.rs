use std::io::{Cursor, Read};

use formula_model::{parse_range_a1, CfRule, CfRuleKind, CfRuleSchema, Workbook};
use roxmltree::Document;
use zip::ZipArchive;

#[test]
fn workbook_writer_normalizes_conditional_formatting_priorities() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").expect("add sheet");

    let range = parse_range_a1("A1:A1").expect("range");
    let rule1 = CfRule {
        schema: CfRuleSchema::Office2007,
        id: None,
        priority: u32::MAX,
        applies_to: vec![range],
        dxf_id: None,
        stop_if_true: false,
        kind: CfRuleKind::Expression {
            formula: "A1>0".to_string(),
        },
        dependencies: vec![range],
    };
    let rule2 = CfRule {
        schema: CfRuleSchema::Office2007,
        id: None,
        priority: u32::MAX,
        applies_to: vec![range],
        dxf_id: None,
        stop_if_true: false,
        kind: CfRuleKind::Expression {
            formula: "A1>1".to_string(),
        },
        dependencies: vec![range],
    };

    let sheet = workbook.sheet_mut(sheet_id).expect("sheet exists");
    sheet.add_conditional_formatting_rule(rule1);
    sheet.add_conditional_formatting_rule(rule2);

    let mut buf = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut buf).expect("write workbook");
    let bytes = buf.into_inner();

    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("open zip");
    let mut sheet_xml = String::new();
    archive
        .by_name("xl/worksheets/sheet1.xml")
        .expect("sheet1.xml exists")
        .read_to_string(&mut sheet_xml)
        .expect("read sheet xml");

    let doc = Document::parse(&sheet_xml).expect("parse xml");
    let mut priorities: Vec<u32> = doc
        .descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "cfRule")
        .filter_map(|n| n.attribute("priority").and_then(|p| p.parse::<u32>().ok()))
        .collect();
    priorities.sort_unstable();

    assert_eq!(
        priorities,
        vec![1, 2],
        "expected normalized priorities, got {priorities:?} in:\n{sheet_xml}"
    );
}

