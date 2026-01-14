use std::io::Cursor;

use formula_model::{
    parse_range_a1, CellIsOperator, CfRule, CfRuleKind, CfRuleSchema, CfStyleOverride, Color,
    Workbook,
};
use roxmltree::Document;

fn extract_cf_rule_dxf_ids(sheet_xml: &str) -> Vec<Option<u32>> {
    let doc = Document::parse(sheet_xml).expect("valid worksheet xml");
    let main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    doc.descendants()
        .filter(|n| n.is_element() && n.tag_name().name() == "cfRule" && n.tag_name().namespace() == Some(main_ns))
        .map(|n| n.attribute("dxfId").and_then(|v| v.parse::<u32>().ok()))
        .collect()
}

fn extract_cf_rule_types(sheet_xml: &str) -> Vec<String> {
    let doc = Document::parse(sheet_xml).expect("valid worksheet xml");
    let main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    doc.descendants()
        .filter(|n| {
            n.is_element()
                && n.tag_name().name() == "cfRule"
                && n.tag_name().namespace() == Some(main_ns)
        })
        .map(|n| n.attribute("type").unwrap_or_default().to_string())
        .collect()
}

fn extract_cf_rule_formulas(sheet_xml: &str) -> Vec<String> {
    let doc = Document::parse(sheet_xml).expect("valid worksheet xml");
    let main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    doc.descendants()
        .filter(|n| {
            n.is_element()
                && n.tag_name().name() == "cfRule"
                && n.tag_name().namespace() == Some(main_ns)
        })
        .map(|rule| {
            rule.children()
                .find(|n| {
                    n.is_element()
                        && n.tag_name().name() == "formula"
                        && n.tag_name().namespace() == Some(main_ns)
                })
                .and_then(|n| n.text())
                .unwrap_or_default()
                .to_string()
        })
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
                // Use a function that requires `_xlfn.` so we can assert writers normalize CF
                // formulas (strip leading `=` and restore the correct file form).
                formula: "=ROWS(SEQUENCE(3))=3".to_string(),
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
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 4,
            applies_to: vec![parse_range_a1("D1").unwrap()],
            dxf_id: Some(1), // blue
            stop_if_true: false,
            kind: CfRuleKind::CellIs {
                operator: CellIsOperator::GreaterThan,
                formulas: vec!["=SEQUENCE(3)".to_string()],
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

    assert_eq!(extract_cf_rule_dxf_ids(sheet1_xml), vec![Some(0), None, None, Some(1)]);
    assert_eq!(extract_cf_rule_dxf_ids(sheet2_xml), vec![Some(1), Some(2)]);

    assert_eq!(
        extract_cf_rule_types(sheet1_xml),
        vec![
            "expression".to_string(),
            "expression".to_string(),
            "expression".to_string(),
            "cellIs".to_string(),
        ]
    );

    // Regression test: writers must strip a leading `=` and apply `_xlfn.` prefixes to modern
    // functions inside conditional formatting formulas.
    assert_eq!(
        extract_cf_rule_formulas(sheet1_xml),
        vec![
            "ROWS(_xlfn.SEQUENCE(3))=3".to_string(),
            "B1>0".to_string(),
            "C1>0".to_string(),
            "_xlfn.SEQUENCE(3)".to_string(),
        ]
    );
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

    assert_eq!(extract_cf_rule_dxf_ids(sheet1_xml), vec![Some(0), None, None, Some(1)]);
    assert_eq!(extract_cf_rule_dxf_ids(sheet2_xml), vec![Some(1), Some(2)]);

    assert_eq!(
        extract_cf_rule_formulas(sheet1_xml),
        vec![
            "ROWS(_xlfn.SEQUENCE(3))=3".to_string(),
            "B1>0".to_string(),
            "C1>0".to_string(),
            "_xlfn.SEQUENCE(3)".to_string(),
        ]
    );
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
                formula: "=ROWS(SEQUENCE(3))=3".to_string(),
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
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 4,
            applies_to: vec![parse_range_a1("D1").unwrap()],
            dxf_id: Some(1), // blue
            stop_if_true: false,
            kind: CfRuleKind::CellIs {
                operator: CellIsOperator::GreaterThan,
                formulas: vec!["=SEQUENCE(3)".to_string()],
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
    assert_eq!(extract_cf_rule_dxf_ids(sheet1_xml), vec![Some(0), None, None, Some(1)]);
    assert_eq!(extract_cf_rule_dxf_ids(sheet2_xml), vec![Some(1), Some(2)]);

    assert_eq!(
        extract_cf_rule_formulas(sheet1_xml),
        vec![
            "ROWS(_xlfn.SEQUENCE(3))=3".to_string(),
            "B1>0".to_string(),
            "C1>0".to_string(),
            "_xlfn.SEQUENCE(3)".to_string(),
        ]
    );
}

#[test]
fn xlsxdocument_roundtrip_injects_cf_with_existing_base_dxfs() {
    // Like `xlsxdocument_roundtrip_injects_cf_and_remaps_cf_rule_dxfid`, but starts from an XLSX
    // that already has `styles.xml <dxfs>` entries. This ensures we remap newly-injected rules to
    // the correct *offset* in the workbook-global dxf table.
    let mut base = Workbook::new();
    let sheet1_id = base.add_sheet("Sheet1").unwrap();
    base.add_sheet("Sheet2").unwrap();

    let dxf_existing = CfStyleOverride {
        fill: Some(Color::new_argb(0xFFFFFF00)),
        ..Default::default()
    };
    let s1_rules = vec![CfRule {
        schema: CfRuleSchema::Office2007,
        id: None,
        priority: 1,
        applies_to: vec![parse_range_a1("A1").unwrap()],
        dxf_id: Some(0), // existing dxf
        stop_if_true: false,
        kind: CfRuleKind::Expression {
            formula: "A1>0".to_string(),
        },
        dependencies: vec![],
    }];
    base.sheet_mut(sheet1_id)
        .unwrap()
        .set_conditional_formatting(s1_rules, vec![dxf_existing.clone()]);

    let mut cursor = Cursor::new(Vec::<u8>::new());
    formula_xlsx::write_workbook_to_writer(&base, &mut cursor).expect("write base workbook");
    let base_bytes = cursor.into_inner();

    let mut doc = formula_xlsx::load_from_bytes(&base_bytes).expect("load xlsx document");
    let sheet2_id = doc.workbook.sheets[1].id;

    let dxf_new = CfStyleOverride {
        font_color: Some(Color::new_argb(0xFF00FF00)),
        ..Default::default()
    };
    let s2_rules = vec![CfRule {
        schema: CfRuleSchema::Office2007,
        id: None,
        priority: 1,
        applies_to: vec![parse_range_a1("B1").unwrap()],
        dxf_id: Some(0), // local idx 0 -> global idx 1 (because `dxf_existing` already occupies 0)
        stop_if_true: false,
        kind: CfRuleKind::Expression {
            formula: "B1>0".to_string(),
        },
        dependencies: vec![],
    }];
    doc.workbook
        .sheet_mut(sheet2_id)
        .unwrap()
        .set_conditional_formatting(s2_rules, vec![dxf_new.clone()]);

    let bytes = doc.save_to_vec().expect("save updated xlsx document");
    let pkg = formula_xlsx::XlsxPackage::from_bytes(&bytes).expect("open updated workbook");

    let styles_xml = std::str::from_utf8(pkg.part("xl/styles.xml").unwrap()).unwrap();
    let styles = formula_xlsx::Styles::parse(styles_xml).unwrap();
    assert_eq!(styles.dxfs.len(), 2);
    assert_eq!(
        styles.dxfs[0].fill,
        Some(Color::new_argb(0xFFFFFF00)),
        "existing dxf should remain at index 0"
    );
    assert_eq!(
        styles.dxfs[1].font_color,
        Some(Color::new_argb(0xFF00FF00)),
        "newly-added dxf should be appended"
    );

    let sheet2_xml = std::str::from_utf8(pkg.part("xl/worksheets/sheet2.xml").unwrap()).unwrap();
    assert_eq!(extract_cf_rule_dxf_ids(sheet2_xml), vec![Some(1)]);
}
