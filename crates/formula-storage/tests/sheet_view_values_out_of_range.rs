use formula_storage::Storage;
use rusqlite::{Connection, OpenFlags};

#[test]
fn export_model_workbook_ignores_out_of_range_sheet_view_values() {
    let uri = "file:sheet_view_values_out_of_range?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet_meta = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");

    conn.execute(
        "UPDATE sheets SET frozen_rows = ?2, frozen_cols = ?3, zoom = ?4 WHERE id = ?1",
        rusqlite::params![
            sheet_meta.id.to_string(),
            u32::MAX as i64 + 1,
            -1i64,
            -2.0f64
        ],
    )
    .expect("write invalid sheet view values");

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");
    let sheet = exported
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("sheet exists");

    assert_eq!(sheet.frozen_rows, 0);
    assert_eq!(sheet.frozen_cols, 0);
    assert_eq!(sheet.zoom, 1.0);
    assert_eq!(sheet.view.pane.frozen_rows, 0);
    assert_eq!(sheet.view.pane.frozen_cols, 0);
    assert_eq!(sheet.view.zoom, 1.0);
}

#[test]
fn sheet_metadata_uses_defaults_when_view_fields_are_null() {
    let uri = "file:sheet_view_values_null?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet_meta = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");

    conn.execute(
        "UPDATE sheets SET frozen_rows = NULL, frozen_cols = NULL, zoom = NULL WHERE id = ?1",
        rusqlite::params![sheet_meta.id.to_string()],
    )
    .expect("write NULL sheet view values");

    let loaded = storage.get_sheet_meta(sheet_meta.id).expect("get sheet meta");
    assert_eq!(loaded.frozen_rows, 0);
    assert_eq!(loaded.frozen_cols, 0);
    assert_eq!(loaded.zoom, 1.0);

    let listed = storage.list_sheets(workbook.id).expect("list sheets");
    let listed = listed
        .iter()
        .find(|s| s.id == sheet_meta.id)
        .expect("sheet listed");
    assert_eq!(listed.frozen_rows, 0);
    assert_eq!(listed.frozen_cols, 0);
    assert_eq!(listed.zoom, 1.0);

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");
    let sheet = exported
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("sheet exists");

    assert_eq!(sheet.frozen_rows, 0);
    assert_eq!(sheet.frozen_cols, 0);
    assert_eq!(sheet.zoom, 1.0);
    assert_eq!(sheet.view.pane.frozen_rows, 0);
    assert_eq!(sheet.view.pane.frozen_cols, 0);
    assert_eq!(sheet.view.zoom, 1.0);
}
