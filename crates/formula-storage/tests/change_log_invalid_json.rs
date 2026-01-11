use formula_storage::{CellChange, CellData, CellValue, Storage};
use rusqlite::{Connection, OpenFlags};

#[test]
fn latest_change_tolerates_invalid_json_columns() {
    let uri = "file:change_log_invalid_json?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    storage
        .apply_cell_changes(&[CellChange {
            sheet_id: sheet.id,
            row: 0,
            col: 0,
            data: CellData {
                value: CellValue::Number(1.0),
                formula: None,
                style: None,
            },
            user_id: None,
        }])
        .expect("apply change");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");
    conn.execute(
        "UPDATE change_log SET target = '{', old_value = '{', new_value = '{' WHERE sheet_id = ?1",
        rusqlite::params![sheet.id.to_string()],
    )
    .expect("corrupt json columns");

    let entry = storage
        .latest_change(sheet.id)
        .expect("latest_change")
        .expect("entry exists");
    assert!(entry.target.is_null());
    assert!(entry.old_value.is_null());
    assert!(entry.new_value.is_null());
}

