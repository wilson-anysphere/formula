use formula_storage::{NamedRange, Storage};
use rusqlite::{params, Connection, OpenFlags};

#[test]
fn sheet_ops_tolerate_missing_workbook_row() {
    let uri = "file:missing_workbook_row?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    storage
        .upsert_named_range(&NamedRange {
            workbook_id: workbook.id,
            name: "Test".to_string(),
            scope: "workbook".to_string(),
            reference: "=Sheet1!A1".to_string(),
        })
        .expect("insert named range");

    // Simulate a corrupted database where the workbook row was removed without cascading.
    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");
    conn.execute_batch("PRAGMA foreign_keys = OFF;")
        .expect("disable foreign keys");
    conn.execute(
        "DELETE FROM workbooks WHERE id = ?1",
        params![workbook.id.to_string()],
    )
    .expect("delete workbook row");
    drop(conn);

    storage
        .rename_sheet(sheet.id, "Renamed")
        .expect("rename should succeed even if workbook row is missing");

    let range = storage
        .get_named_range(workbook.id, "Test", "workbook")
        .expect("get named range")
        .expect("named range exists");
    assert!(
        range.reference.contains("Renamed"),
        "expected named range reference to be rewritten, got {}",
        range.reference
    );

    storage
        .upsert_named_range(&NamedRange {
            workbook_id: workbook.id,
            name: "Test".to_string(),
            scope: "workbook".to_string(),
            reference: "=Renamed!B1".to_string(),
        })
        .expect("upsert should succeed even if workbook row is missing");

    storage
        .delete_sheet(sheet.id)
        .expect("delete should succeed even if workbook row is missing");
}

