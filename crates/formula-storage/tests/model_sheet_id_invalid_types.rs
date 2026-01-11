use formula_model::DefinedNameScope;
use formula_storage::{NamedRange, Storage};
use rusqlite::{Connection, OpenFlags};

#[test]
fn export_and_create_sheet_tolerate_invalid_model_sheet_id_types() {
    let uri = "file:model_sheet_id_invalid_types?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet1 = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");
    conn.execute(
        "UPDATE sheets SET model_sheet_id = 'bogus' WHERE id = ?1",
        rusqlite::params![sheet1.id.to_string()],
    )
    .expect("corrupt model_sheet_id");

    // Export should ignore the invalid `model_sheet_id` row instead of failing.
    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");
    assert!(exported.sheets.iter().any(|s| s.name == "Sheet1"));

    // Creating a new sheet should still succeed even if existing rows contain invalid `model_sheet_id`
    // values.
    let _sheet2 = storage
        .create_sheet(workbook.id, "Sheet2", 1, None)
        .expect("create second sheet");

    // Ensure sheet-scoped named range sync tolerates invalid `model_sheet_id` values on other sheets.
    storage
        .upsert_named_range(&NamedRange {
            workbook_id: workbook.id,
            name: "Scoped".to_string(),
            scope: "Sheet2".to_string(),
            reference: "Sheet2!$A$1".to_string(),
        })
        .expect("upsert named range");

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");
    assert!(exported.defined_names.iter().any(|name| {
        name.name == "Scoped"
            && name.scope == DefinedNameScope::Sheet(1)
            && name.refers_to == "Sheet2!$A$1"
    }));

    // Deleting a sheet with a corrupt `model_sheet_id` should still succeed.
    storage.delete_sheet(sheet1.id).expect("delete sheet");
}
