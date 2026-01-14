use formula_model::{
    Cell, CellRef, CfRule, CfRuleKind, CfRuleSchema, DataValidation, DataValidationKind,
    DefinedNameScope, Hyperlink, HyperlinkTarget, Range,
};
use formula_storage::{ImportModelWorkbookOptions, Storage};

#[test]
fn rename_sheet_rewrites_cell_formulas_and_defined_names() {
    let mut workbook = formula_model::Workbook::new();
    let data_id = workbook.add_sheet("Data").expect("add sheet Data");
    let summary_id = workbook.add_sheet("Summary").expect("add sheet Summary");

    workbook
        .create_defined_name(
            DefinedNameScope::Workbook,
            "MyRange",
            "=Data!$A$1",
            None,
            false,
            None,
        )
        .expect("create workbook defined name");
    workbook
        .create_defined_name(
            DefinedNameScope::Sheet(data_id),
            "MyLocal",
            "=Data!$A$2",
            None,
            false,
            None,
        )
        .expect("create sheet defined name");

    // Add print settings so we can validate they get renamed too.
    assert!(
        workbook.set_sheet_print_area(
            data_id,
            Some(vec![Range::new(CellRef::new(0, 0), CellRef::new(2, 2))]),
        ),
        "set print area"
    );

    // Formula in Summary referencing the sheet we're going to rename.
    {
        let sheet = workbook.sheet_mut(summary_id).expect("summary sheet");
        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                value: formula_model::CellValue::Empty,
                formula: Some("Data!A1".to_string()),
                phonetic: None,
                style_id: 0,
                phonetic: None,
            },
        );

        sheet.add_data_validation(
            vec![Range::new(CellRef::new(0, 0), CellRef::new(0, 0))],
            DataValidation {
                kind: DataValidationKind::Custom,
                operator: None,
                formula1: "Data!A1".to_string(),
                formula2: None,
                allow_blank: false,
                show_input_message: false,
                show_error_message: false,
                show_drop_down: false,
                input_message: None,
                error_alert: None,
            },
        );

        sheet.conditional_formatting_rules.push(CfRule {
            schema: CfRuleSchema::Office2007,
            id: None,
            priority: 1,
            applies_to: vec![Range::new(CellRef::new(0, 0), CellRef::new(0, 0))],
            dxf_id: None,
            stop_if_true: false,
            kind: CfRuleKind::Expression {
                formula: "Data!A1=1".to_string(),
            },
            dependencies: Vec::new(),
        });

        sheet.hyperlinks.push(Hyperlink::for_cell(
            CellRef::new(0, 0),
            HyperlinkTarget::Internal {
                sheet: "Data".to_string(),
                cell: CellRef::new(0, 0),
            },
        ));
    }

    let storage = Storage::open_in_memory().expect("open storage");
    let meta = storage
        .import_model_workbook(&workbook, ImportModelWorkbookOptions::new("Book"))
        .expect("import");
    let data_sheet_uuid = storage
        .list_sheets(meta.id)
        .expect("list sheets")
        .iter()
        .find(|s| s.name == "Data")
        .expect("data sheet")
        .id;

    storage
        .rename_sheet(data_sheet_uuid, "Renamed")
        .expect("rename sheet");

    // Sheet-scoped named ranges should now use the new sheet name as their scope identifier.
    assert!(storage
        .get_named_range(meta.id, "MyLocal", "Renamed")
        .expect("get local range")
        .is_some());

    let exported = storage.export_model_workbook(meta.id).expect("export");

    let summary_sheet = exported
        .sheets
        .iter()
        .find(|s| s.name == "Summary")
        .expect("summary sheet");
    let formula = summary_sheet
        .cell(CellRef::new(0, 0))
        .and_then(|c| c.formula.as_deref())
        .expect("formula exists");
    assert_eq!(formula, "Renamed!A1");

    let dv_formula = summary_sheet
        .data_validations
        .first()
        .map(|dv| dv.validation.formula1.as_str())
        .expect("data validation exists");
    assert_eq!(dv_formula, "Renamed!A1");

    let cf_formula = summary_sheet
        .conditional_formatting_rules
        .first()
        .and_then(|rule| match &rule.kind {
            CfRuleKind::Expression { formula } => Some(formula.as_str()),
            _ => None,
        })
        .expect("conditional formatting rule exists");
    assert_eq!(cf_formula, "Renamed!A1=1");

    let hyperlink_target = summary_sheet
        .hyperlinks
        .first()
        .and_then(|link| match &link.target {
            HyperlinkTarget::Internal { sheet, .. } => Some(sheet.as_str()),
            _ => None,
        })
        .expect("hyperlink exists");
    assert_eq!(hyperlink_target, "Renamed");

    assert!(exported.defined_names.iter().any(|n| {
        n.name == "MyRange" && n.scope == DefinedNameScope::Workbook && n.refers_to == "Renamed!$A$1"
    }));

    let renamed_sheet_id = exported
        .sheets
        .iter()
        .find(|s| s.name == "Renamed")
        .expect("renamed sheet")
        .id;
    assert!(exported.defined_names.iter().any(|n| {
        n.name == "MyLocal"
            && n.scope == DefinedNameScope::Sheet(renamed_sheet_id)
            && n.refers_to == "Renamed!$A$2"
    }));

    assert!(exported
        .print_settings
        .sheets
        .iter()
        .any(|s| s.sheet_name == "Renamed"));
}
