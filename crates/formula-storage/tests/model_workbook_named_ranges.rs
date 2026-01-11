use formula_model::DefinedNameScope;
use formula_storage::{NamedRange, Storage};
use rusqlite::{Connection, OpenFlags};

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

#[test]
fn export_model_workbook_prefers_latest_named_range_for_unicode_scope_duplicates() {
    let uri = "file:model_workbook_named_ranges_unicode_scope_dup?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    storage
        .create_sheet(workbook.id, "Äbc", 0, None)
        .expect("create sheet");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");

    // Insert an older row with a scope that sorts after the newer row under SQLite `COLLATE NOCASE`
    // (which is ASCII-only). Export should still prefer the newest row by `rowid`.
    conn.execute(
        "INSERT INTO named_ranges (workbook_id, name, scope, reference) VALUES (?1, ?2, ?3, ?4)",
        rusqlite::params![workbook.id.to_string(), "LocalRange", "äbc", "=1"],
    )
    .expect("insert older named range");
    conn.execute(
        "INSERT INTO named_ranges (workbook_id, name, scope, reference) VALUES (?1, ?2, ?3, ?4)",
        rusqlite::params![workbook.id.to_string(), "LocalRange", "Äbc", "=2"],
    )
    .expect("insert newer named range");

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");
    let sheet_id = exported
        .sheets
        .iter()
        .find(|s| s.name == "Äbc")
        .expect("sheet exists")
        .id;

    let defined_name = exported
        .defined_names
        .iter()
        .find(|n| n.name == "LocalRange" && n.scope == DefinedNameScope::Sheet(sheet_id))
        .expect("defined name exists");
    assert_eq!(defined_name.refers_to, "2");
}
