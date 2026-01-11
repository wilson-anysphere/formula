use formula_model::DefinedNameScope;
use formula_storage::{NamedRange, Storage};

#[test]
fn export_model_workbook_includes_named_ranges_from_legacy_storage() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    storage
        .upsert_named_range(&NamedRange {
            workbook_id: workbook.id,
            name: "MyRange".to_string(),
            scope: "workbook".to_string(),
            reference: "Sheet1!$A$1".to_string(),
        })
        .expect("insert workbook named range");

    storage
        .upsert_named_range(&NamedRange {
            workbook_id: workbook.id,
            name: "LocalRange".to_string(),
            scope: "Sheet1".to_string(),
            reference: "=Sheet1!$B$2".to_string(),
        })
        .expect("insert sheet named range");

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");
    let sheet_id = exported
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("sheet exists")
        .id;

    assert!(exported.defined_names.iter().any(|n| {
        n.name == "MyRange"
            && n.scope == DefinedNameScope::Workbook
            && n.refers_to == "Sheet1!$A$1"
    }));

    assert!(exported.defined_names.iter().any(|n| {
        n.name == "LocalRange"
            && n.scope == DefinedNameScope::Sheet(sheet_id)
            && n.refers_to == "Sheet1!$B$2"
    }));
}

