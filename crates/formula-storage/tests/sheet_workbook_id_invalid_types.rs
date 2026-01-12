use formula_model::TabColor;
use formula_storage::{SheetVisibility, Storage};
use rusqlite::{params, Connection, OpenFlags};

#[test]
fn sheet_crud_tolerates_invalid_workbook_id_types() {
    let uri = "file:sheet_workbook_id_invalid_types?mode=memory&cache=shared";
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
    conn.execute_batch("PRAGMA foreign_keys = OFF;")
        .expect("disable foreign keys");

    // Corrupt the sheet's workbook_id with a non-TEXT value.
    conn.execute(
        "UPDATE sheets SET workbook_id = X'00' WHERE id = ?1",
        params![sheet.id.to_string()],
    )
    .expect("corrupt workbook_id");
    drop(conn);

    storage
        .rename_sheet(sheet.id, "Renamed")
        .expect("rename sheet even if workbook_id is corrupt");
    storage
        .set_sheet_visibility(sheet.id, SheetVisibility::Hidden)
        .expect("set visibility even if workbook_id is corrupt");
    let tab_color = TabColor::rgb("FF112233");
    storage
        .set_sheet_tab_color(sheet.id, Some(&tab_color))
        .expect("set tab color even if workbook_id is corrupt");
    storage
        .set_sheet_xlsx_metadata(sheet.id, Some(42), Some("rId7"))
        .expect("set xlsx metadata even if workbook_id is corrupt");
    storage
        .reorder_sheet(sheet.id, 5)
        .expect("reorder sheet even if workbook_id is corrupt");
    storage
        .delete_sheet(sheet.id)
        .expect("delete sheet even if workbook_id is corrupt");

    // Verify the sheet row is gone.
    let conn = Connection::open_with_flags(uri, flags).expect("reopen raw connection");
    let remaining: i64 = conn
        .query_row(
            "SELECT COUNT(*) FROM sheets WHERE id = ?1",
            params![sheet.id.to_string()],
            |r| r.get(0),
        )
        .expect("count sheets");
    assert_eq!(remaining, 0);
}
