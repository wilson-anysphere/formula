use formula_storage::Storage;
use std::collections::HashMap;

#[test]
fn export_model_workbook_preserves_sheet_ids_after_reorder() {
    let storage = Storage::open_in_memory().expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet_a = storage
        .create_sheet(workbook.id, "SheetA", 0, None)
        .expect("create sheet A");
    storage
        .create_sheet(workbook.id, "SheetB", 1, None)
        .expect("create sheet B");
    let sheet_c = storage
        .create_sheet(workbook.id, "SheetC", 2, None)
        .expect("create sheet C");

    let exported1 = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");
    let ids1: HashMap<_, _> = exported1
        .sheets
        .iter()
        .map(|s| (s.name.clone(), s.id))
        .collect();

    storage
        .reorder_sheet(sheet_c.id, 0)
        .expect("reorder sheet");
    let exported2 = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");
    let ids2: HashMap<_, _> = exported2
        .sheets
        .iter()
        .map(|s| (s.name.clone(), s.id))
        .collect();

    assert_eq!(ids2, ids1);
    assert_eq!(
        exported2.sheets.iter().map(|s| s.name.as_str()).collect::<Vec<_>>(),
        vec!["SheetC", sheet_a.name.as_str(), "SheetB"]
    );
}

