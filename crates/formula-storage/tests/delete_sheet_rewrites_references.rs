use formula_model::{
    Cell, CellRef, CfRule, CfRuleKind, CfRuleSchema, DataValidation, DataValidationKind,
    DefinedNameScope, Range,
};
use formula_storage::{ImportModelWorkbookOptions, Storage};

#[test]
fn delete_sheet_rewrites_references_in_formulas_and_metadata() {
    let mut workbook = formula_model::Workbook::new();
    let data_id = workbook.add_sheet("Data").expect("add sheet Data");
    let summary_id = workbook.add_sheet("Summary").expect("add sheet Summary");
    workbook.add_sheet("Other").expect("add sheet Other");

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

    assert!(
        workbook.set_sheet_print_area(
            data_id,
            Some(vec![Range::new(CellRef::new(0, 0), CellRef::new(2, 2))]),
        ),
        "set print area"
    );

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

    storage.delete_sheet(data_sheet_uuid).expect("delete sheet");

    assert!(storage
        .get_named_range(meta.id, "MyLocal", "Data")
        .expect("get local range")
        .is_none());

    let renamed = storage
        .get_named_range(meta.id, "MyRange", "workbook")
        .expect("get range")
        .expect("workbook range exists");
    assert_eq!(renamed.reference, "#REF!");

    let exported = storage.export_model_workbook(meta.id).expect("export");
    assert!(exported.sheets.iter().all(|s| s.name != "Data"));

    let summary_sheet = exported
        .sheets
        .iter()
        .find(|s| s.name == "Summary")
        .expect("summary sheet");
    let formula = summary_sheet
        .cell(CellRef::new(0, 0))
        .and_then(|c| c.formula.as_deref())
        .expect("formula exists");
    assert_eq!(formula, "#REF!");

    let dv_formula = summary_sheet
        .data_validations
        .first()
        .map(|dv| dv.validation.formula1.as_str())
        .expect("data validation exists");
    assert_eq!(dv_formula, "#REF!");

    let cf_formula = summary_sheet
        .conditional_formatting_rules
        .first()
        .and_then(|rule| match &rule.kind {
            CfRuleKind::Expression { formula } => Some(formula.as_str()),
            _ => None,
        })
        .expect("conditional formatting rule exists");
    assert_eq!(cf_formula, "#REF!=1");

    assert!(exported.defined_names.iter().any(|n| {
        n.name == "MyRange" && n.scope == DefinedNameScope::Workbook && n.refers_to == "#REF!"
    }));
    assert!(exported.defined_names.iter().all(|n| n.name != "MyLocal"));

    assert!(exported
        .print_settings
        .sheets
        .iter()
        .all(|s| s.sheet_name != "Data"));
}
