use formula_storage::Storage;
use rusqlite::{Connection, OpenFlags};

#[test]
fn sheet_metadata_uses_defaults_when_view_fields_have_invalid_types() {
    let uri = "file:sheet_meta_invalid_types?mode=memory&cache=shared";
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

    // Store non-numeric TEXT values into numeric columns to simulate corruption. These would
    // previously cause `r.get::<_, i64/f64>` to return an error.
    conn.execute(
        r#"
        UPDATE sheets
        SET position = 'bogus',
            frozen_rows = 'bogus',
            frozen_cols = 'bogus',
            zoom = 'bogus',
            xlsx_sheet_id = 'bogus',
            xlsx_rel_id = X'00',
            tab_color = X'00'
        WHERE id = ?1
        "#,
        rusqlite::params![sheet.id.to_string()],
    )
    .expect("corrupt sheet view fields");

    let meta = storage.get_sheet_meta(sheet.id).expect("get sheet meta");
    assert_eq!(meta.position, 0);
    assert_eq!(meta.frozen_rows, 0);
    assert_eq!(meta.frozen_cols, 0);
    assert_eq!(meta.zoom, 1.0);
    assert!(meta.tab_color.is_none());
    assert!(meta.xlsx_sheet_id.is_none());
    assert!(meta.xlsx_rel_id.is_none());

    let sheets = storage.list_sheets(workbook.id).expect("list sheets");
    assert_eq!(sheets.len(), 1);
    assert_eq!(sheets[0].id, sheet.id);
    assert_eq!(sheets[0].zoom, 1.0);

    // `create_sheet` relies on the transactional `list_sheets_tx` helper; ensure it still works.
    storage
        .create_sheet(workbook.id, "Sheet2", 1, None)
        .expect("create sheet with corrupted view fields present");
}
