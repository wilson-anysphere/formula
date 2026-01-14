use std::io::Cursor;

use formula_model::{parse_range_a1, CfRule, CfRuleKind, CfRuleSchema, CfStyleOverride, Color, Workbook};
use roxmltree::Document;

fn extract_cf_rule_dxf_ids(sheet_xml: &str) -> Vec<Option<u32>> {
    let doc = Document::parse(sheet_xml).expect("valid worksheet xml");
    let main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    doc.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "cfRule" && n.tag_name().namespace() == Some(main_ns))
        .map(|n| n.attribute("dxfId").and_then(|v| v.parse::<u32>().ok()))
        .collect()
}

fn build_workbook() -> Workbook {
    let mut wb = Workbook::new();
    let sheet1_id = wb.add_sheet("Sheet1").unwrap();
    let sheet2_id = wb.add_sheet("Sheet2").unwrap();

    let dxf_red = CfStyleOverride {
        fill: Some(Color::new_argb(0xFFFF0000)),
        ..Default::default()
    };
    let dxf_blue = CfStyleOverride {
        fill: Some(Color::new_argb(0xFF0000FF)),
        ..Default::default()
    };
    let dxf_green_font = CfStyleOverride {
        font_color: Some(Color::new_argb(0xFF00FF00)),
        ..Default::default()
    };

    let s1_rules = vec![
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![parse_range_a1("A1").unwrap()],
            dxf_id: Some(0), // red
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "A1>0".to_string(),
            },
            dependencies: vec![],
        },
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 2,
            applies_to: vec![parse_range_a1("B1").unwrap()],
            dxf_id: None, // stays None
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "B1>0".to_string(),
            },
            dependencies: vec![],
        },
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 3,
            applies_to: vec![parse_range_a1("C1").unwrap()],
            dxf_id: Some(99), // out of bounds -> omitted
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "C1>0".to_string(),
            },
            dependencies: vec![],
        },
    ];
    let s1_dxfs = vec![dxf_red.clone(), dxf_blue.clone()];
    wb.sheet_mut(sheet1_id)
        .unwrap()
        .set_conditional_formatting(s1_rules, s1_dxfs);

    let s2_rules = vec![
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![parse_range_a1("A1").unwrap()],
            dxf_id: Some(0), // blue (local)
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "A1>0".to_string(),
            },
            dependencies: vec![],
        },
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 2,
            applies_to: vec![parse_range_a1("B1").unwrap()],
            dxf_id: Some(1), // green font (local)
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "B1>0".to_string(),
            },
            dependencies: vec![],
        },
    ];
    let s2_dxfs = vec![dxf_blue.clone(), dxf_green_font.clone()];
    wb.sheet_mut(sheet2_id)
        .unwrap()
        .set_conditional_formatting(s2_rules, s2_dxfs);

    wb
}

#[test]
fn workbook_writer_aggregates_dxfs_and_remaps_cf_rule_dxfid() {
    let wb = build_workbook();
    let mut cursor = Cursor::new(Vec::<u8>::new());
    formula_xlsx::write_workbook_to_writer(&wb, &mut cursor).expect("write workbook");
    let bytes = cursor.into_inner();

    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("open written workbook");
    let styles_xml = std::str::from_utf8(pkg.part("xl/styles.xml").unwrap()).unwrap();
    let styles = formula_xlsx::Styles::parse(styles_xml).unwrap();

    let expected = vec![
        CfStyleOverride {
            fill: Some(Color::new_argb(0xFFFF0000)),
            ..Default::default()
        },
        CfStyleOverride {
            fill: Some(Color::new_argb(0xFF0000FF)),
            ..Default::default()
        },
        CfStyleOverride {
            font_color: Some(Color::new_argb(0xFF00FF00)),
            ..Default::default()
        },
    ];
    assert_eq!(styles.dxfs, expected, "styles.xml should contain deduped union of dxfs");

    let sheet1_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let sheet2_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet2.xml").unwrap()).unwrap();

    assert_eq!(extract_cf_rule_dxf_ids(sheet1_xml), vec![Some(0), None, None]);
    assert_eq!(extract_cf_rule_dxf_ids(sheet2_xml), vec![Some(1), Some(2)]);
}

#[test]
fn xlsxdocument_writer_aggregates_dxfs_and_remaps_cf_rule_dxfid() {
    let wb = build_workbook();
    let doc = formula_xlsx::XlsxDocument::new(wb);
    let bytes = doc.save_to_vec().expect("save xlsx document");

    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("open written workbook");
    let styles_xml = std::str::from_utf8(pkg.part("xl/styles.xml").unwrap()).unwrap();
    let styles = formula_xlsx::Styles::parse(styles_xml).unwrap();

    assert_eq!(styles.dxfs.len(), 3);
    assert_eq!(
        styles.dxfs[0].fill,
        Some(Color::new_argb(0xFFFF0000)),
        "first-seen ordering should be stable"
    );
    assert_eq!(styles.dxfs[1].fill, Some(Color::new_argb(0xFF0000FF)));
    assert_eq!(
        styles.dxfs[2].font_color,
        Some(Color::new_argb(0xFF00FF00))
    );

    let sheet1_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let sheet2_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet2.xml").unwrap()).unwrap();

    assert_eq!(extract_cf_rule_dxf_ids(sheet1_xml), vec![Some(0), None, None]);
    assert_eq!(extract_cf_rule_dxf_ids(sheet2_xml), vec![Some(1), Some(2)]);
}

#[test]
fn xlsxdocument_roundtrip_injects_cf_and_remaps_cf_rule_dxfid() {
    // Regression test: when saving an `XlsxDocument` loaded from an existing package, we may need
    // to *inject* `<conditionalFormatting>` into an existing worksheet XML payload. In that case,
    // `CfRule.dxf_id` values (worksheet-local indices) must be remapped to the workbook-global
    // `styles.xml` `<dxfs>` index space.
    let mut base = Workbook::new();
    base.add_sheet("Sheet1").unwrap();
    base.add_sheet("Sheet2").unwrap();

    let mut cursor = Cursor::new(Vec::<u8>::new());
    formula_xlsx::write_workbook_to_writer(&base, &mut cursor).expect("write base workbook");
    let base_bytes = cursor.into_inner();

    let mut doc = formula_xlsx::load_from_bytes(&base_bytes).expect("load xlsx document");
    let sheet1_id = doc.workbook.sheets[0].id;
    let sheet2_id = doc.workbook.sheets[1].id;

    let dxf_red = CfStyleOverride {
        fill: Some(Color::new_argb(0xFFFF0000)),
        ..Default::default()
    };
    let dxf_blue = CfStyleOverride {
        fill: Some(Color::new_argb(0xFF0000FF)),
        ..Default::default()
    };
    let dxf_green_font = CfStyleOverride {
        font_color: Some(Color::new_argb(0xFF00FF00)),
        ..Default::default()
    };

    let s1_rules = vec![
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![parse_range_a1("A1").unwrap()],
            dxf_id: Some(0), // red
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "A1>0".to_string(),
            },
            dependencies: vec![],
        },
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 2,
            applies_to: vec![parse_range_a1("B1").unwrap()],
            dxf_id: None, // stays None
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "B1>0".to_string(),
            },
            dependencies: vec![],
        },
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 3,
            applies_to: vec![parse_range_a1("C1").unwrap()],
            dxf_id: Some(99), // out of bounds -> omitted
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "C1>0".to_string(),
            },
            dependencies: vec![],
        },
    ];
    doc.workbook
        .sheet_mut(sheet1_id)
        .unwrap()
        .set_conditional_formatting(s1_rules, vec![dxf_red.clone(), dxf_blue.clone()]);

    let s2_rules = vec![
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![parse_range_a1("A1").unwrap()],
            dxf_id: Some(0), // blue (local) -> global 1
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "A1>0".to_string(),
            },
            dependencies: vec![],
        },
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 2,
            applies_to: vec![parse_range_a1("B1").unwrap()],
            dxf_id: Some(1), // green (local) -> global 2
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "B1>0".to_string(),
            },
            dependencies: vec![],
        },
    ];
    doc.workbook
        .sheet_mut(sheet2_id)
        .unwrap()
        .set_conditional_formatting(s2_rules, vec![dxf_blue.clone(), dxf_green_font.clone()]);

    let bytes = doc.save_to_vec().expect("save updated xlsx document");
    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("open updated workbook");

    let styles_xml = std::str::from_utf8(pkg.part("xl/styles.xml").unwrap()).unwrap();
    let styles = formula_xlsx::Styles::parse(styles_xml).unwrap();
    assert_eq!(styles.dxfs.len(), 3);
    assert_eq!(styles.dxfs[0].fill, Some(Color::new_argb(0xFFFF0000)));
    assert_eq!(styles.dxfs[1].fill, Some(Color::new_argb(0xFF0000FF)));
    assert_eq!(
        styles.dxfs[2].font_color,
        Some(Color::new_argb(0xFF00FF00))
    );

    let sheet1_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet1.xml").unwrap()).unwrap();
    let sheet2_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet2.xml").unwrap()).unwrap();
    assert_eq!(extract_cf_rule_dxf_ids(sheet1_xml), vec![Some(0), None, None]);
    assert_eq!(extract_cf_rule_dxf_ids(sheet2_xml), vec![Some(1), Some(2)]);
}
