use formula_model::{
    CellIsOperator, CellRef, CfRule, CfRuleKind, CfRuleSchema, Cfvo, CfvoType, DataBarRule,
    DataValidation, DataValidationKind, DataValidationOperator, Hyperlink, HyperlinkTarget, Range,
    Table, TableColumn, Workbook,
};

#[test]
fn rename_sheet_rewrites_all_modeled_surfaces() {
    let mut workbook = Workbook::new();
    let old_sheet_id = workbook.add_sheet("OldSheet").unwrap();
    let other_sheet_id = workbook.add_sheet("Sheet2").unwrap();

    let other_sheet = workbook.sheet_mut(other_sheet_id).expect("sheet2");

    // Cell formula.
    other_sheet.set_formula(CellRef::new(0, 0), Some("=OldSheet!A1".to_string()));

    // Table formulas (stored without leading '=').
    other_sheet.tables.push(Table {
        id: 1,
        name: "Table1".to_string(),
        display_name: "Table1".to_string(),
        range: Range::new(CellRef::new(0, 0), CellRef::new(1, 0)),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![TableColumn {
            id: 1,
            name: "Col1".to_string(),
            formula: Some("OldSheet!A1".to_string()),
            totals_formula: Some("SUM(OldSheet!A1)".to_string()),
        }],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    });

    // Conditional formatting formulas.
    other_sheet.conditional_formatting_rules.extend([
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![Range::new(CellRef::new(0, 0), CellRef::new(0, 0))],
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "OldSheet!A1>0".to_string(),
            },
            dependencies: Vec::new(),
        },
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 2,
            applies_to: vec![Range::new(CellRef::new(0, 0), CellRef::new(0, 0))],
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::CellIs {
                operator: CellIsOperator::GreaterThan,
                formulas: vec!["OldSheet!A1".to_string()],
            },
            dependencies: Vec::new(),
        },
        CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 3,
            applies_to: vec![Range::new(CellRef::new(0, 0), CellRef::new(0, 0))],
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::DataBar(DataBarRule {
                min: Cfvo {
                    type_: CfvoType::Formula,
                    value: Some("OldSheet!A1".to_string()),
                },
                max: Cfvo {
                    type_: CfvoType::Max,
                    value: None,
                },
                color: None,
                min_length: None,
                max_length: None,
                gradient: None,
                negative_fill_color: None,
                axis_color: None,
                direction: None,
            }),
            dependencies: Vec::new(),
        },
    ]);

    // Internal hyperlink target (case-insensitive sheet matching).
    other_sheet.hyperlinks.push(Hyperlink::for_cell(
        CellRef::new(1, 1),
        HyperlinkTarget::Internal {
            sheet: "oldsheet".to_string(),
            cell: CellRef::new(1, 1),
        },
    ));

    let list_validation_id = other_sheet.add_data_validation(
        vec![Range::new(CellRef::new(2, 0), CellRef::new(2, 0))],
        DataValidation {
            kind: DataValidationKind::List,
            operator: None,
            formula1: "OldSheet!A1:A3".to_string(),
            formula2: None,
            allow_blank: false,
            show_input_message: false,
            show_error_message: false,
            show_drop_down: false,
            input_message: None,
            error_alert: None,
        },
    );

    let between_validation_id = other_sheet.add_data_validation(
        vec![Range::new(CellRef::new(3, 0), CellRef::new(3, 0))],
        DataValidation {
            kind: DataValidationKind::Decimal,
            operator: Some(DataValidationOperator::Between),
            formula1: "OldSheet!A1".to_string(),
            formula2: Some("OldSheet!A2".to_string()),
            allow_blank: false,
            show_input_message: false,
            show_error_message: false,
            show_drop_down: false,
            input_message: None,
            error_alert: None,
        },
    );

    workbook
        .rename_sheet(old_sheet_id, "New Sheet")
        .expect("rename succeeds");

    assert_eq!(workbook.sheet(old_sheet_id).unwrap().name, "New Sheet");

    let other_sheet = workbook.sheet(other_sheet_id).expect("sheet2");

    assert_eq!(
        other_sheet.formula(CellRef::new(0, 0)).unwrap(),
        "'New Sheet'!A1"
    );

    let table = other_sheet.tables.first().expect("table1");
    let col = table.columns.first().expect("col1");
    assert_eq!(col.formula.as_deref().unwrap(), "'New Sheet'!A1");
    assert_eq!(
        col.totals_formula.as_deref().unwrap(),
        "SUM('New Sheet'!A1)"
    );

    let expr_rule = other_sheet
        .conditional_formatting_rules
        .iter()
        .find(|r| matches!(r.kind, CfRuleKind::Expression { .. }))
        .expect("expression rule");
    match &expr_rule.kind {
        CfRuleKind::Expression { formula } => assert_eq!(formula, "'New Sheet'!A1>0"),
        _ => unreachable!(),
    }

    let cell_is_rule = other_sheet
        .conditional_formatting_rules
        .iter()
        .find(|r| matches!(r.kind, CfRuleKind::CellIs { .. }))
        .expect("cellIs rule");
    match &cell_is_rule.kind {
        CfRuleKind::CellIs { formulas, .. } => {
            assert_eq!(formulas, &vec!["'New Sheet'!A1".to_string()])
        }
        _ => unreachable!(),
    }

    let data_bar_rule = other_sheet
        .conditional_formatting_rules
        .iter()
        .find(|r| matches!(r.kind, CfRuleKind::DataBar(_)))
        .expect("dataBar rule");
    match &data_bar_rule.kind {
        CfRuleKind::DataBar(db) => {
            assert_eq!(db.min.value.as_deref().unwrap(), "'New Sheet'!A1");
        }
        _ => unreachable!(),
    }

    let link = other_sheet.hyperlinks.first().expect("hyperlink");
    match &link.target {
        HyperlinkTarget::Internal { sheet, cell } => {
            assert_eq!(sheet, "New Sheet");
            assert_eq!(*cell, CellRef::new(1, 1));
        }
        _ => panic!("expected internal hyperlink"),
    }

    let list_validation = other_sheet
        .data_validations
        .iter()
        .find(|dv| dv.id == list_validation_id)
        .expect("list validation");
    assert_eq!(list_validation.validation.formula1, "'New Sheet'!A1:A3");

    let between_validation = other_sheet
        .data_validations
        .iter()
        .find(|dv| dv.id == between_validation_id)
        .expect("between validation");
    assert_eq!(between_validation.validation.formula1, "'New Sheet'!A1");
    assert_eq!(
        between_validation.validation.formula2.as_deref().unwrap(),
        "'New Sheet'!A2"
    );
}
