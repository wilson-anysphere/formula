use formula_storage::Storage;
use rusqlite::{Connection, OpenFlags};

#[test]
fn sheet_operations_tolerate_invalid_sheet_name_types() {
    let uri = "file:sheet_name_invalid_types?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet1 = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet1");
    let sheet2 = storage
        .create_sheet(workbook.id, "Sheet2", 1, None)
        .expect("create sheet2");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");

    // Corrupt the `name` column by storing a BLOB. APIs that enumerate sheets should tolerate the
    // row and surface it with a deterministic placeholder name so it can be renamed/deleted.
    conn.execute(
        "UPDATE sheets SET name = X'00' WHERE id = ?1",
        rusqlite::params![sheet1.id.to_string()],
    )
    .expect("corrupt sheet name");

    let sheet3 = storage
        .create_sheet(workbook.id, "Sheet3", 2, None)
        .expect("create sheet3");

    let sheets = storage.list_sheets(workbook.id).expect("list sheets");
    assert_eq!(sheets.len(), 3);
    let sheet1_meta = sheets.iter().find(|s| s.id == sheet1.id).expect("sheet1 present");
    assert!(sheet1_meta.name.starts_with("_invalid_"));
    assert!(sheets.iter().any(|s| s.id == sheet2.id && s.name == "Sheet2"));
    assert!(sheets.iter().any(|s| s.id == sheet3.id && s.name == "Sheet3"));

    // Recovery path: rename the corrupted sheet to a valid name.
    storage
        .rename_sheet(sheet1.id, "Restored")
        .expect("rename corrupted sheet");

    storage
        .rename_sheet(sheet2.id, "Renamed")
        .expect("rename sheet");

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");
    assert!(exported.sheets.iter().any(|s| s.name == "Restored"));
    assert!(exported.sheets.iter().any(|s| s.name == "Renamed"));
    assert!(exported.sheets.iter().any(|s| s.name == "Sheet3"));

    storage.delete_sheet(sheet1.id).expect("delete restored sheet");
    storage.delete_sheet(sheet2.id).expect("delete sheet");
}
