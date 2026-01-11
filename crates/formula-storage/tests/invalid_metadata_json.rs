use formula_storage::Storage;
use rusqlite::{Connection, OpenFlags};

#[test]
fn sheet_and_workbook_metadata_ignore_invalid_json() {
    let uri = "file:invalid_metadata_json?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");

    conn.execute(
        "UPDATE workbooks SET metadata = '{' WHERE id = ?1",
        rusqlite::params![workbook.id.to_string()],
    )
    .expect("corrupt workbook metadata");
    conn.execute(
        "UPDATE sheets SET metadata = '{' WHERE id = ?1",
        rusqlite::params![sheet.id.to_string()],
    )
    .expect("corrupt sheet metadata");

    let fetched_workbook = storage.get_workbook(workbook.id).expect("get workbook");
    assert!(fetched_workbook.metadata.is_none());

    let workbooks = storage.list_workbooks().expect("list workbooks");
    assert_eq!(workbooks.len(), 1);
    assert!(workbooks[0].metadata.is_none());

    let fetched_sheet = storage.get_sheet_meta(sheet.id).expect("get sheet");
    assert!(fetched_sheet.metadata.is_none());

    // `create_sheet` uses `list_sheets_tx`, which must also tolerate invalid metadata.
    storage
        .create_sheet(workbook.id, "Sheet2", 1, None)
        .expect("create sheet with corrupt metadata present");

    let sheets = storage.list_sheets(workbook.id).expect("list sheets");
    assert_eq!(sheets.len(), 2);
    assert!(sheets.iter().any(|s| s.name == "Sheet1" && s.metadata.is_none()));

    storage
        .rename_sheet(sheet.id, "Renamed")
        .expect("rename sheet with corrupt metadata present");
}

