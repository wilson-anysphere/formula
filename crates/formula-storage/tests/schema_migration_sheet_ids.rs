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

#[test]
fn migration_v8_tolerates_invalid_workbook_and_sheet_id_types() {
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
    drop(storage);

    // Downgrade to schema v7 and inject corrupt rows that would previously crash the v8 migration
    // when it attempted to read TEXT ids as Strings.
    let conn = Connection::open(path).expect("open raw db");
    conn.execute_batch("DROP INDEX IF EXISTS idx_sheets_workbook_model_sheet_id;")
        .expect("drop unique index");
    conn.execute("UPDATE schema_version SET version = 7 WHERE id = 1", [])
        .expect("downgrade schema version");

    // Corrupt an existing sheet's model_sheet_id with a non-integer type so the migration has to
    // clear and reallocate it.
    conn.execute(
        "UPDATE sheets SET model_sheet_id = X'00' WHERE id = ?1",
        params![sheet_a.id.to_string()],
    )
    .expect("corrupt model_sheet_id");

    // Insert a corrupt workbook row with a non-TEXT primary key.
    conn.execute(
        "INSERT INTO workbooks (id, name) VALUES (X'00', 'Corrupt')",
        [],
    )
    .expect("insert corrupt workbook");

    // Insert a corrupt sheet row with a non-TEXT id (but a valid workbook_id) and a non-integer
    // model_sheet_id.
    conn.execute(
        "INSERT INTO sheets (id, workbook_id, name, model_sheet_id) VALUES (X'00', ?1, 'Bad', X'00')",
        params![workbook.id.to_string()],
    )
    .expect("insert corrupt sheet");
    drop(conn);

    let storage = Storage::open_path(path).expect("open with v8 migration");
    drop(storage);

    let conn = Connection::open(path).expect("reopen raw db");
    let version: i64 = conn
        .query_row("SELECT version FROM schema_version WHERE id = 1", [], |r| r.get(0))
        .expect("schema version");
    assert_eq!(version, 8);

    let sheet_b_model_id: i64 = conn
        .query_row(
            "SELECT model_sheet_id FROM sheets WHERE id = ?1",
            params![sheet_b.id.to_string()],
            |r| r.get(0),
        )
        .expect("fetch sheet b model id");

    // Ensure the unique index is active by attempting to introduce a duplicate.
    let result = conn.execute(
        "UPDATE sheets SET model_sheet_id = ?1 WHERE id = ?2",
        params![sheet_b_model_id, sheet_a.id.to_string()],
    );
    assert!(result.is_err(), "expected unique index to reject duplicates");
}

#[test]
fn migration_v8_ignores_orphaned_sheets_when_creating_unique_index() {
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
    drop(storage);

    let conn = Connection::open(path).expect("open raw db");
    conn.execute_batch("DROP INDEX IF EXISTS idx_sheets_workbook_model_sheet_id;")
        .expect("drop unique index");
    conn.execute("UPDATE schema_version SET version = 7 WHERE id = 1", [])
        .expect("downgrade schema version");
    conn.execute_batch("PRAGMA foreign_keys = OFF;")
        .expect("disable foreign keys");

    // Insert orphaned sheet rows with an invalid workbook_id type and duplicate model_sheet_id
    // values. These rows are not addressable via the public API, but they must not prevent the v8
    // unique index from being created.
    conn.execute(
        "INSERT INTO sheets (id, workbook_id, name, model_sheet_id) VALUES ('orphan1', X'00', 'Orphan', 1)",
        [],
    )
    .expect("insert orphan sheet 1");
    conn.execute(
        "INSERT INTO sheets (id, workbook_id, name, model_sheet_id) VALUES ('orphan2', X'00', 'Orphan', 1)",
        [],
    )
    .expect("insert orphan sheet 2");
    drop(conn);

    let storage = Storage::open_path(path).expect("open with v8 migration");
    drop(storage);

    let conn = Connection::open(path).expect("reopen raw db");
    let version: i64 = conn
        .query_row("SELECT version FROM schema_version WHERE id = 1", [], |r| r.get(0))
        .expect("schema version");
    assert_eq!(version, 8);

    let orphan_ids_with_model_id: i64 = conn
        .query_row(
            "SELECT COUNT(*) FROM sheets WHERE typeof(workbook_id) != 'text' AND model_sheet_id IS NOT NULL",
            [],
            |r| r.get(0),
        )
        .expect("count orphan sheet ids");
    assert_eq!(orphan_ids_with_model_id, 0);

    let sheet_a_model_id: i64 = conn
        .query_row(
            "SELECT model_sheet_id FROM sheets WHERE id = ?1",
            params![sheet_a.id.to_string()],
            |r| r.get(0),
        )
        .expect("fetch sheet a model id");

    // Ensure the unique index is active by attempting to introduce a duplicate.
    let result = conn.execute(
        "UPDATE sheets SET model_sheet_id = ?1 WHERE id = ?2",
        params![sheet_a_model_id, sheet_b.id.to_string()],
    );
    assert!(result.is_err(), "expected unique index to reject duplicates");
}
