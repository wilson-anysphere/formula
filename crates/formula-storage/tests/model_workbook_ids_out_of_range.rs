use formula_storage::Storage;
use rusqlite::{Connection, OpenFlags};

#[test]
fn export_model_workbook_ignores_out_of_range_model_workbook_fields() {
    let uri = "file:model_workbook_ids_out_of_range?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");

    conn.execute(
        "UPDATE workbooks SET model_schema_version = ?2, model_workbook_id = ?3 WHERE id = ?1",
        rusqlite::params![
            workbook.id.to_string(),
            -1i64,
            u32::MAX as i64 + 1
        ],
    )
    .expect("write invalid workbook metadata");

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");

    assert_eq!(exported.schema_version, formula_model::SCHEMA_VERSION);
    assert_eq!(exported.id, 0);
}

