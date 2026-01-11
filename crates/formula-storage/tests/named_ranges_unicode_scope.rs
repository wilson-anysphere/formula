use formula_model::rewrite_sheet_names_in_formula;
use formula_storage::{NamedRange, Storage};
use rusqlite::{Connection, OpenFlags};

#[test]
fn named_ranges_sheet_scopes_are_unicode_case_insensitive() {
    let uri = "file:named_ranges_unicode_scope?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Äbc", 0, None)
        .expect("create sheet");

    storage
        .upsert_named_range(&NamedRange {
            workbook_id: workbook.id,
            name: "Local".to_string(),
            scope: "Äbc".to_string(),
            reference: "Äbc!$A$1".to_string(),
        })
        .expect("insert named range");

    let fetched = storage
        .get_named_range(workbook.id, "local", "äbc")
        .expect("get named range")
        .expect("named range exists");
    assert_eq!(fetched.reference, "Äbc!$A$1");

    // Updating the same named range with a different Unicode case in the scope should update the
    // existing row rather than inserting a duplicate.
    storage
        .upsert_named_range(&NamedRange {
            workbook_id: workbook.id,
            name: "LOCAL".to_string(),
            scope: "äbc".to_string(),
            reference: "Äbc!$B$2".to_string(),
        })
        .expect("update named range");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");
    let count: i64 = conn
        .query_row(
            "SELECT COUNT(*) FROM named_ranges WHERE workbook_id = ?1",
            rusqlite::params![workbook.id.to_string()],
            |r| r.get(0),
        )
        .expect("count named ranges");
    assert_eq!(count, 1);

    // Renaming the sheet should update the sheet-scoped named range scopes even when Unicode
    // case folding is involved.
    storage.rename_sheet(sheet.id, "äbc").expect("rename sheet");
    let fetched = storage
        .get_named_range(workbook.id, "Local", "ÄBC")
        .expect("get renamed scope")
        .expect("named range exists");
    assert_eq!(fetched.scope, "äbc");
    // Non-ASCII sheet names are emitted as quoted sheet references in formulas.
    assert_eq!(fetched.reference, "'äbc'!$B$2");
}

#[test]
fn rename_sheet_deduplicates_unicode_case_scoped_named_ranges() {
    let uri = "file:named_ranges_unicode_scope_dedup?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Äbc", 0, None)
        .expect("create sheet");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");

    conn.execute(
        "INSERT INTO named_ranges (workbook_id, name, scope, reference) VALUES (?1, ?2, ?3, ?4)",
        rusqlite::params![workbook.id.to_string(), "Local", "Äbc", "Äbc!$A$1"],
    )
    .expect("insert named range A1");
    conn.execute(
        "INSERT INTO named_ranges (workbook_id, name, scope, reference) VALUES (?1, ?2, ?3, ?4)",
        rusqlite::params![workbook.id.to_string(), "Local", "äbc", "Äbc!$B$2"],
    )
    .expect("insert named range B2");

    storage
        .rename_sheet(sheet.id, "Renamed")
        .expect("rename sheet");

    let fetched = storage
        .get_named_range(workbook.id, "LOCAL", "renamed")
        .expect("get named range")
        .expect("named range exists");
    assert_eq!(fetched.scope, "Renamed");
    assert_eq!(
        fetched.reference,
        rewrite_sheet_names_in_formula("Äbc!$B$2", "Äbc", "Renamed")
    );

    let count: i64 = conn
        .query_row(
            "SELECT COUNT(*) FROM named_ranges WHERE workbook_id = ?1 AND name = ?2",
            rusqlite::params![workbook.id.to_string(), "Local"],
            |r| r.get(0),
        )
        .expect("count named ranges");
    assert_eq!(count, 1);
}
