use formula_storage::Storage;
use rusqlite::{Connection, OpenFlags};

#[test]
fn export_model_workbook_ignores_invalid_workbook_metadata_shapes() {
    let uri = "file:model_workbook_invalid_metadata?mode=memory&cache=shared";
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

    // Store valid JSON values with the wrong shapes so `serde_json::from_value::<T>` fails.
    conn.execute(
        r#"
        UPDATE workbooks
        SET date_system = 'bogus',
            calc_settings = '"bogus"',
            theme = '"bogus"',
            workbook_protection = '"bogus"',
            defined_names = '"bogus"',
            print_settings = '"bogus"',
            view = '"bogus"'
        WHERE id = ?1
        "#,
        rusqlite::params![workbook.id.to_string()],
    )
    .expect("write invalid workbook metadata");

    // Also write a broken tab_color_json shape; export should fall back to the legacy string column.
    conn.execute(
        "UPDATE sheets SET tab_color = ?2, tab_color_json = ?3 WHERE workbook_id = ?1",
        rusqlite::params![workbook.id.to_string(), "FFFF0000", "\"bogus\""],
    )
    .expect("write invalid tab_color_json");

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");

    assert_eq!(exported.date_system, formula_model::DateSystem::Excel1900);
    assert_eq!(exported.calc_settings, formula_model::CalcSettings::default());
    assert!(exported.defined_names.is_empty());
    assert!(exported.print_settings.is_empty());
    assert_eq!(exported.view, formula_model::WorkbookView::default());

    let sheet = exported
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("sheet exists");
    assert_eq!(sheet.tab_color.as_ref().and_then(|c| c.rgb.clone()), Some("FFFF0000".to_string()));
}

