use formula_model::{parse_range_a1, CfRule, CfRuleKind, CfRuleSchema, CfStyleOverride, Color};
use formula_xlsx::{
    load_from_bytes, ConditionalFormattingDxfAggregation, StylesPart, XlsxPackage,
};
use roxmltree::Document;

const OFFICE2007_FIXTURE: &[u8] = include_bytes!("fixtures/conditional_formatting_2007.xlsx");

fn extract_cf_rule_dxf_ids(sheet_xml: &str) -> Vec<Option<u32>> {
    let doc = Document::parse(sheet_xml).expect("valid worksheet xml");
    let main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    doc.descendants()
        .filter(|n| {
            n.is_element() && n.tag_name().name() == "cfRule" && n.tag_name().namespace() == Some(main_ns)
        })
        .map(|n| n.attribute("dxfId").and_then(|v| v.parse::<u32>().ok()))
        .collect()
}

#[test]
fn xlsxdocument_appends_new_dxfs_when_editing_existing_workbook() {
    let mut doc = load_from_bytes(OFFICE2007_FIXTURE).expect("load fixture via XlsxDocument");

    let original_pkg = XlsxPackage::from_bytes(OFFICE2007_FIXTURE).expect("open original fixture");
    let original_styles_xml = std::str::from_utf8(original_pkg.part("xl/styles.xml").unwrap())
        .expect("styles.xml should be utf-8");
    let original_styles = formula_xlsx::Styles::parse(original_styles_xml).expect("parse styles");
    let base_dxf_count = original_styles.dxfs.len();
    assert!(base_dxf_count > 0, "fixture should contain at least one dxf");
    let mut style_table = doc.workbook.styles.clone();
    let styles_part = StylesPart::parse_or_default(
        Some(original_pkg.part("xl/styles.xml").unwrap()),
        &mut style_table,
    )
    .expect("parse styles part");
    let base_dxfs_styles_part = styles_part.conditional_formatting_dxfs();
    assert_eq!(
        base_dxfs_styles_part, original_styles.dxfs,
        "expected styles.xml dxf parsing to be consistent across Styles and StylesPart"
    );
    assert_eq!(
        doc.workbook
            .sheets
            .first()
            .expect("fixture should contain at least one sheet")
            .conditional_formatting_dxfs
            .len(),
        base_dxf_count,
        "expected loaded workbook to populate worksheet conditional formatting dxfs from styles.xml"
    );
    assert!(
        !doc.workbook
            .sheets
            .first()
            .unwrap()
            .conditional_formatting_rules
            .is_empty(),
        "expected loaded workbook to populate conditional formatting rules"
    );

    let sheet2_id = doc.workbook.add_sheet("Sheet2").expect("add sheet");
    let sheet1_id = doc
        .workbook
        .sheets
        .first()
        .expect("fixture should contain at least one sheet")
        .id;
    assert_eq!(
        doc.workbook.sheet(sheet1_id).unwrap().conditional_formatting_dxfs,
        original_styles.dxfs,
        "expected existing worksheet conditional formatting dxfs to stay in sync with styles.xml"
    );
    assert_ne!(
        sheet2_id, sheet1_id,
        "expected newly added sheet to receive a distinct WorksheetId"
    );
    let dxf_blue = CfStyleOverride {
        fill: Some(Color::new_argb(0xFF0000FF)),
        ..Default::default()
    };
    let rules = vec![CfRule {
        schema: CfRuleSchema::Office2007,
        id: None,
        priority: 1,
        applies_to: vec![parse_range_a1("A1").unwrap()],
        dxf_id: Some(0), // worksheet-local index; should be remapped to the appended global index.
        stop_if_true: false,
        kind: CfRuleKind::Expression {
            formula: "A1>0".to_string(),
        },
        dependencies: vec![],
    }];
    doc.workbook
        .sheet_mut(sheet2_id)
        .unwrap()
        .set_conditional_formatting(rules, vec![dxf_blue.clone()]);
    assert_eq!(
        doc.workbook
            .sheet(sheet2_id)
            .unwrap()
            .conditional_formatting_dxfs,
        vec![dxf_blue.clone()],
        "expected newly added worksheet to retain its local dxfs vector"
    );
    let agg = ConditionalFormattingDxfAggregation::from_worksheets_with_base_global_dxfs(
        &doc.workbook.sheets,
        &original_styles.dxfs,
    );
    assert_eq!(
        agg.local_to_global_by_sheet
            .get(&sheet2_id)
            .expect("expected aggregation to include newly added sheet")
            .as_slice(),
        &[base_dxf_count as u32],
        "expected dxf aggregation to remap the new sheet's local dxf indices after the existing base table"
    );

    let saved = doc.save_to_vec().expect("save XlsxDocument");
    let saved_pkg = XlsxPackage::from_bytes(&saved).expect("open saved package");

    let styles_xml = std::str::from_utf8(saved_pkg.part("xl/styles.xml").unwrap())
        .expect("saved styles.xml should be utf-8");
    let styles = formula_xlsx::Styles::parse(styles_xml).expect("parse saved styles");
    assert_eq!(
        styles.dxfs.len(),
        base_dxf_count + 1,
        "expected newly introduced dxf to be appended to styles.xml"
    );
    assert_eq!(
        styles.dxfs.first(),
        original_styles.dxfs.first(),
        "expected existing dxfs to remain at the start of the table so existing dxfId indices stay stable"
    );
    assert_eq!(
        styles.dxfs.last().unwrap().fill,
        Some(Color::new_argb(0xFF0000FF)),
        "expected appended dxf to match the newly added conditional formatting style"
    );

    let sheet2_xml = std::str::from_utf8(saved_pkg.part("xl/worksheets/sheet2.xml").unwrap())
        .expect("sheet2.xml should be utf-8");
    assert_eq!(
        extract_cf_rule_dxf_ids(sheet2_xml),
        vec![Some(base_dxf_count as u32)],
        "expected worksheet-local dxfId=0 to be remapped to the appended global dxf index"
    );
}
