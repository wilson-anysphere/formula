use formula_model::{parse_range_a1, CellIsOperator, CfRule, CfRuleKind, CfRuleSchema};
use formula_xlsx::{load_from_bytes, XlsxPackage};

const FIXTURE: &[u8] = include_bytes!("fixtures/rt_mc.xlsx");

#[test]
fn inserting_conditional_formatting_preserves_mc_alternatecontent_block() {
    let mut doc = load_from_bytes(FIXTURE).expect("load fixture");

    let sheet_id = doc.workbook.sheets[0].id;
    let sheet = doc
        .workbook
        .sheet_mut(sheet_id)
        .expect("sheet exists in workbook");

    let range = parse_range_a1("A1:A1").expect("valid range");
    sheet.add_conditional_formatting_rule(CfRule {
        schema: CfRuleSchema::Office2007,
        id: None,
        priority: 1,
        applies_to: vec![range],
        dxf_id: None,
        stop_if_true: false,
        kind: CfRuleKind::CellIs {
            operator: CellIsOperator::GreaterThan,
            formulas: vec!["0".to_string()],
        },
        dependencies: vec![range],
    });

    let saved = doc.save_to_vec().expect("save");

    let pkg = XlsxPackage::from_bytes(&saved).expect("reopen");
    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap())
        .expect("worksheet xml is utf-8");

    assert!(
        sheet_xml.contains("<mc:AlternateContent"),
        "expected worksheet XML to preserve mc:AlternateContent, got:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<mc:Choice"),
        "expected worksheet XML to preserve mc:Choice, got:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains("<mc:Fallback"),
        "expected worksheet XML to preserve mc:Fallback, got:\n{sheet_xml}"
    );

    assert!(
        sheet_xml.contains("<conditionalFormatting"),
        "expected worksheet XML to contain conditionalFormatting, got:\n{sheet_xml}"
    );

    let cf_pos = sheet_xml
        .find("<conditionalFormatting")
        .expect("conditionalFormatting position");
    let mc_pos = sheet_xml
        .find("<mc:AlternateContent")
        .expect("mc:AlternateContent position");
    assert!(
        cf_pos < mc_pos,
        "expected conditionalFormatting to appear before mc:AlternateContent, got:\n{sheet_xml}"
    );
}

