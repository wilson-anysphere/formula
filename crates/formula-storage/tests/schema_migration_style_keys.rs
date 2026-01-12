use formula_storage::Storage;
use rusqlite::{params, Connection};
use tempfile::NamedTempFile;

#[test]
fn migration_v3_tolerates_duplicate_style_component_keys() {
    let tmp = NamedTempFile::new().expect("tmp file");
    let path = tmp.path();

    let storage = Storage::open_path(path).expect("open storage");
    storage
        .create_workbook("Book", None)
        .expect("create workbook");
    drop(storage);

    let conn = Connection::open(path).expect("open raw db");
    conn.execute_batch("DROP INDEX IF EXISTS idx_fonts_key;")
        .expect("drop unique index");
    // Force migration v3 to re-run so we attempt to recreate the unique index.
    conn.execute("UPDATE schema_version SET version = 2 WHERE id = 1", [])
        .expect("downgrade schema version");

    // Insert two rows with the same non-NULL key; unique index creation would fail without
    // deduplication.
    conn.execute(
        "INSERT INTO fonts (key, data) VALUES (?1, ?2)",
        params!["dup", "{}"],
    )
    .expect("insert dup font 1");
    conn.execute(
        "INSERT INTO fonts (key, data) VALUES (?1, ?2)",
        params!["dup", "{}"],
    )
    .expect("insert dup font 2");
    drop(conn);

    let storage = Storage::open_path(path).expect("open with migrations");
    drop(storage);

    let conn = Connection::open(path).expect("reopen raw db");
    let count: i64 = conn
        .query_row("SELECT COUNT(*) FROM fonts WHERE key = 'dup'", [], |r| r.get(0))
        .expect("count keys");
    assert_eq!(count, 1, "expected duplicate keys to be cleared before index");

    let result = conn.execute("INSERT INTO fonts (key, data) VALUES ('dup', '{}')", []);
    assert!(
        result.is_err(),
        "expected unique index to reject reintroducing duplicate keys"
    );
}

