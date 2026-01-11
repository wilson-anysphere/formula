use formula_storage::Storage;
use rusqlite::{params, Connection};
use std::collections::HashSet;
use tempfile::NamedTempFile;

#[test]
fn migration_v8_deduplicates_model_sheet_ids() {
    let tmp = NamedTempFile::new().expect("tmp file");
    let path = tmp.path();

    let storage = Storage::open_path(path).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet_a = storage
        .create_sheet(workbook.id, "SheetA", 0, None)
        .expect("create sheet a");
    let sheet_b = storage
        .create_sheet(workbook.id, "SheetB", 1, None)
        .expect("create sheet b");
    let sheet_c = storage
        .create_sheet(workbook.id, "SheetC", 2, None)
        .expect("create sheet c");
    drop(storage);

    // Corrupt the DB to simulate a pre-v8 client that wrote duplicate model_sheet_id values.
    let conn = Connection::open(path).expect("open raw db");
    conn.execute_batch("DROP INDEX IF EXISTS idx_sheets_workbook_model_sheet_id;")
        .expect("drop unique index");
    conn.execute("UPDATE schema_version SET version = 7 WHERE id = 1", [])
        .expect("downgrade schema version");
    let sheet_a_model_id: i64 = conn
        .query_row(
            "SELECT model_sheet_id FROM sheets WHERE id = ?1",
            params![sheet_a.id.to_string()],
            |r| r.get(0),
        )
        .expect("fetch sheet a model id");
    conn.execute(
        "UPDATE sheets SET model_sheet_id = ?1 WHERE id = ?2",
        params![sheet_a_model_id, sheet_b.id.to_string()],
    )
    .expect("force duplicate model_sheet_id");
    drop(conn);

    let storage = Storage::open_path(path).expect("open with v8 migration");
    drop(storage);

    let conn = Connection::open(path).expect("reopen raw db");
    let version: i64 = conn
        .query_row("SELECT version FROM schema_version WHERE id = 1", [], |r| r.get(0))
        .expect("schema version");
    assert_eq!(version, 8);

    let mut stmt = conn
        .prepare("SELECT model_sheet_id FROM sheets WHERE workbook_id = ?1")
        .expect("prepare select");
    let ids = stmt
        .query_map(params![workbook.id.to_string()], |row| row.get::<_, i64>(0))
        .expect("query ids")
        .collect::<rusqlite::Result<Vec<_>>>()
        .expect("collect ids");
    assert_eq!(ids.len(), 3);
    let mut seen = HashSet::new();
    for id in ids {
        assert!(
            seen.insert(id),
            "duplicate model_sheet_id after migration: {id}"
        );
    }

    // Ensure the unique index is active by attempting to introduce a duplicate.
    let result = conn.execute(
        "UPDATE sheets SET model_sheet_id = ?1 WHERE id = ?2",
        params![sheet_a_model_id, sheet_c.id.to_string()],
    );
    assert!(result.is_err(), "expected unique index to reject duplicates");
}

