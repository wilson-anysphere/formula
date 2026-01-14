use formula_model::{CellIsOperator, CellRef, CellValue, CfRule, CfRuleKind, CfRuleSchema, CfStyleOverride, Color, Range, Workbook};
use formula_xlsx::{write_workbook_to_writer, XlsxPackage};
use std::io::Cursor;

#[test]
fn export_writes_conditional_formatting_and_styles_dxfs() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").unwrap();
    let sheet = workbook.sheet_mut(sheet_id).unwrap();

    sheet.set_value(CellRef::from_a1("A1").unwrap(), CellValue::Number(15.0));

    sheet.conditional_formatting_dxfs = vec![CfStyleOverride {
        fill: Some(Color::new_argb(0xFFFF0000)),
        font_color: Some(Color::new_argb(0xFF00FF00)),
        bold: Some(true),
        italic: None,
    }];

    sheet.conditional_formatting_rules = vec![CfRule {
        schema: CfRuleSchema::Office2007,
        id: None,
        priority: 1,
        applies_to: vec![Range::new(
            CellRef::from_a1("A1").unwrap(),
            CellRef::from_a1("A1").unwrap(),
        )],
        dxf_id: Some(0),
        stop_if_true: false,
        kind: CfRuleKind::CellIs {
            operator: CellIsOperator::GreaterThan,
            formulas: vec!["10".to_string()],
        },
        dependencies: vec![],
    }];

    let mut cursor = Cursor::new(Vec::new());
    write_workbook_to_writer(&workbook, &mut cursor).unwrap();
    let bytes = cursor.into_inner();

    let pkg = XlsxPackage::from_bytes(&bytes).unwrap();
    let sheet_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let styles_xml = std::str::from_utf8(pkg.part("xl/styles.xml").unwrap()).unwrap();

    assert!(
        sheet_xml.contains("<conditionalFormatting"),
        "expected sheet1.xml to contain conditionalFormatting, got:\n{sheet_xml}"
    );
    assert!(
        sheet_xml.contains(r#"dxfId="0""#),
        "expected sheet1.xml to reference dxfId=0, got:\n{sheet_xml}"
    );

    assert!(
        styles_xml.contains(r#"<dxfs count="1">"#),
        "expected styles.xml to contain dxfs count=1, got:\n{styles_xml}"
    );
    assert!(
        styles_xml.contains(r#"fgColor rgb="FFFF0000""#),
        "expected styles.xml to contain fill fgColor, got:\n{styles_xml}"
    );
    assert!(
        styles_xml.contains(r#"color rgb="FF00FF00""#),
        "expected styles.xml to contain font color, got:\n{styles_xml}"
    );
}

