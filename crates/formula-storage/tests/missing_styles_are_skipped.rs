use formula_storage::Storage;
use rusqlite::{Connection, OpenFlags};

#[test]
fn export_model_workbook_skips_cells_with_missing_style_rows() {
    let uri = "file:missing_styles_are_skipped?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let sheet = storage
        .create_sheet(workbook.id, "Sheet1", 0, None)
        .expect("create sheet");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");

    // Insert a cell that references a non-existent style row. The `cells` table does not enforce
    // foreign keys for `style_id`, so corrupted databases can contain such rows.
    conn.execute(
        r#"
        INSERT INTO cells (sheet_id, row, col, value_type, value_number, style_id)
        VALUES (?1, 0, 0, 'number', 42.0, 999)
        "#,
        rusqlite::params![sheet.id.to_string()],
    )
    .expect("insert cell");

    // Also simulate a broken workbook_styles entry (this table has an FK, so we need to disable
    // enforcement to inject invalid data).
    conn.pragma_update(None, "foreign_keys", "OFF")
        .expect("disable foreign keys");
    conn.execute(
        r#"
        INSERT INTO workbook_styles (workbook_id, style_index, style_id)
        VALUES (?1, 0, 999)
        "#,
        rusqlite::params![workbook.id.to_string()],
    )
    .expect("insert invalid workbook_styles row");
    conn.pragma_update(None, "foreign_keys", "ON")
        .expect("re-enable foreign keys");

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");

    assert_eq!(exported.styles.len(), 1, "missing styles should be skipped");

    let exported_sheet = exported
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("sheet exists");
    let cell = exported_sheet
        .cell(formula_model::CellRef::new(0, 0))
        .expect("cell exists");
    assert_eq!(cell.value, formula_model::CellValue::Number(42.0));
    assert_eq!(cell.style_id, 0);
}

