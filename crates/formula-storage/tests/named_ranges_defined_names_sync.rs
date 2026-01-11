use formula_model::{DefinedName, DefinedNameScope};
use formula_storage::{NamedRange, Storage};
use rusqlite::{Connection, OpenFlags};

#[test]
fn upsert_named_range_keeps_workbook_defined_names_in_sync() {
    let uri = "file:named_ranges_defined_names_sync?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    storage
        .upsert_named_range(&NamedRange {
            workbook_id: workbook.id,
            name: "MyRange".to_string(),
            scope: "workbook".to_string(),
            reference: "=Sheet1!$A$1".to_string(),
        })
        .expect("insert workbook named range");

    // Updating the range with a different ASCII case should update the existing defined name rather
    // than inserting a duplicate.
    storage
        .upsert_named_range(&NamedRange {
            workbook_id: workbook.id,
            name: "MYRANGE".to_string(),
            scope: "workbook".to_string(),
            reference: "=Sheet1!$B$2".to_string(),
        })
        .expect("update workbook named range");

    storage
        .upsert_named_range(&NamedRange {
            workbook_id: workbook.id,
            name: "LocalRange".to_string(),
            scope: "Sheet1".to_string(),
            reference: "=Sheet1!$C$3".to_string(),
        })
        .expect("insert sheet named range");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");

    let sheet_model_id: u32 = conn
        .query_row(
            "SELECT model_sheet_id FROM sheets WHERE id = ?1",
            rusqlite::params![sheet.id.to_string()],
            |r| {
                let raw: i64 = r.get(0)?;
                u32::try_from(raw).map_err(|_| rusqlite::Error::InvalidQuery)
            },
        )
        .expect("read model_sheet_id");

    let raw_defined_names: Option<serde_json::Value> = conn
        .query_row(
            "SELECT defined_names FROM workbooks WHERE id = ?1",
            rusqlite::params![workbook.id.to_string()],
            |r| r.get(0),
        )
        .expect("read defined_names json");
    let defined_names: Vec<DefinedName> =
        serde_json::from_value(raw_defined_names.expect("defined_names stored"))
            .expect("parse defined_names");

    assert!(defined_names.iter().any(|name| {
        name.scope == DefinedNameScope::Workbook
            && name.name.eq_ignore_ascii_case("MyRange")
            && name.refers_to == "Sheet1!$B$2"
    }));

    assert!(defined_names.iter().any(|name| {
        name.scope == DefinedNameScope::Sheet(sheet_model_id)
            && name.name.eq_ignore_ascii_case("LocalRange")
            && name.refers_to == "Sheet1!$C$3"
    }));

    let workbook_scope_count = defined_names
        .iter()
        .filter(|n| n.scope == DefinedNameScope::Workbook && n.name.eq_ignore_ascii_case("MyRange"))
        .count();
    assert_eq!(workbook_scope_count, 1);
}

