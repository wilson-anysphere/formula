use formula_storage::Storage;
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
    storage
        .create_sheet(workbook.id, "Sheet2", 1, None)
        .expect("create second sheet");
}

