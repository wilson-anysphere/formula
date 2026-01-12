use formula_storage::Storage;
use rusqlite::{Connection, OpenFlags};

#[test]
fn list_workbooks_skips_invalid_ids_and_tolerates_invalid_names() {
    let uri = "file:list_workbooks_invalid_types?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Good", None)
        .expect("create workbook");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");

    // Insert a corrupt workbook row with a non-TEXT id/name.
    conn.execute(
        "INSERT INTO workbooks (id, name) VALUES (X'00', X'00')",
        [],
    )
    .expect("insert corrupt workbook");

    // Corrupt the name of the valid workbook as well; listing should still succeed and surface
    // the workbook with a placeholder name.
    conn.execute(
        "UPDATE workbooks SET name = X'00' WHERE id = ?1",
        rusqlite::params![workbook.id.to_string()],
    )
    .expect("corrupt workbook name");

    let workbooks = storage.list_workbooks().expect("list workbooks");
    assert_eq!(workbooks.len(), 1);
    assert_eq!(workbooks[0].id, workbook.id);
    assert!(workbooks[0].name.starts_with("_invalid_workbook_"));

    let fetched = storage.get_workbook(workbook.id).expect("get workbook");
    assert!(fetched.name.starts_with("_invalid_workbook_"));

    storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet even with corrupt workbook name");
}
