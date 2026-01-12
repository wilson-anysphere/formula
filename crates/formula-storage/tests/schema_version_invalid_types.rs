use formula_storage::Storage;
use rusqlite::{params, Connection};
use tempfile::NamedTempFile;

#[test]
fn schema_init_tolerates_invalid_schema_version_types() {
    let tmp = NamedTempFile::new().expect("tmp file");
    let path = tmp.path();

    // Create a valid database at the latest schema version.
    let storage = Storage::open_path(path).expect("open storage");
    drop(storage);

    // Corrupt the schema version value so it can't be decoded as an integer.
    let conn = Connection::open(path).expect("open raw db");
    conn.execute(
        "UPDATE schema_version SET version = ?1 WHERE id = 1",
        params![&[0u8][..]],
    )
    .expect("corrupt schema version");
    drop(conn);

    // Reopening should succeed and restore schema_version.version to a valid integer.
    let storage = Storage::open_path(path).expect("reopen after schema version corruption");
    drop(storage);

    let conn = Connection::open(path).expect("reopen raw db");
    let (version, type_name): (i64, String) = conn
        .query_row(
            "SELECT version, typeof(version) FROM schema_version WHERE id = 1",
            [],
            |r| Ok((r.get(0)?, r.get(1)?)),
        )
        .expect("read schema version");
    assert_eq!(version, 8);
    assert_eq!(type_name, "integer");
}

