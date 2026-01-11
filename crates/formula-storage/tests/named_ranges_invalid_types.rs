use formula_model::DefinedNameScope;
use formula_storage::Storage;
use rusqlite::{Connection, OpenFlags};

#[test]
fn export_model_workbook_skips_invalid_named_ranges_rows() {
    let uri = "file:named_ranges_invalid_types?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");

    conn.execute(
        r#"
        INSERT INTO named_ranges (workbook_id, name, scope, reference)
        VALUES (?1, 'MyRange', 'workbook', '=Sheet1!$A$1')
        "#,
        rusqlite::params![workbook.id.to_string()],
    )
    .expect("insert named range");

    // Corrupt the `name` column by storing a BLOB. Export should ignore this row instead of failing.
    conn.execute(
        r#"
        INSERT INTO named_ranges (workbook_id, name, scope, reference)
        VALUES (?1, X'00', 'workbook', '=Sheet1!$A$2')
        "#,
        rusqlite::params![workbook.id.to_string()],
    )
    .expect("insert invalid named range");

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");

    assert_eq!(exported.defined_names.len(), 1);
    assert!(exported.defined_names.iter().any(|name| {
        name.name == "MyRange"
            && name.scope == DefinedNameScope::Workbook
            && name.refers_to == "Sheet1!$A$1"
    }));
}

