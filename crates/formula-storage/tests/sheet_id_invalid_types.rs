use formula_storage::Storage;
use rusqlite::{params, Connection, OpenFlags};

#[test]
fn rename_and_delete_skip_sheets_with_invalid_id_types_in_metadata_rewrite_passes() {
    let uri = "file:sheet_id_invalid_types?mode=memory&cache=shared";
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

    // Insert a corrupt sheet row that has a non-TEXT id and a non-NULL model_sheet_json so that
    // rename/delete metadata rewrite passes will encounter it.
    conn.execute(
        r#"
        INSERT INTO sheets (id, workbook_id, name, position, model_sheet_json)
        VALUES (?1, ?2, 'Corrupt', 1, '{')
        "#,
        params![&[0u8][..], workbook.id.to_string()],
    )
    .expect("insert corrupt sheet");
    drop(conn);

    storage
        .rename_sheet(sheet.id, "Renamed")
        .expect("rename should skip corrupt sheet ids");
    storage
        .delete_sheet(sheet.id)
        .expect("delete should skip corrupt sheet ids");
}

