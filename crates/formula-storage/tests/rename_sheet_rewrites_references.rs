use formula_model::{Cell, CellRef, DefinedNameScope, Range};
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
                style_id: 0,
            },
        );
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

