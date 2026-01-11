use formula_model::{CellRef, CellValue};
use formula_storage::{CellRange, Storage};
use rusqlite::{Connection, OpenFlags};

#[test]
fn cell_reads_and_export_tolerate_invalid_value_json_and_types() {
    let uri = "file:cells_invalid_json?mode=memory&cache=shared";
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

    // Corrupt the canonical JSON cell payload. Previously this would fail JSON parsing and break
    // both reads and model export.
    conn.execute(
        r#"
        INSERT INTO cells (sheet_id, row, col, value_type, value_number, value_string, value_json)
        VALUES (?1, 0, 0, 'number', NULL, NULL, '{')
        "#,
        rusqlite::params![sheet.id.to_string()],
    )
    .expect("insert invalid json cell");

    // Store a non-TEXT value in `value_json`; rusqlite cannot deserialize it into a `String`.
    conn.execute(
        r#"
        INSERT INTO cells (sheet_id, row, col, value_type, value_number, value_string, value_json)
        VALUES (?1, 0, 1, 'string', NULL, NULL, 123)
        "#,
        rusqlite::params![sheet.id.to_string()],
    )
    .expect("insert invalid type value_json cell");

    // Insert a valid scalar-only cell (no value_json) for sanity.
    conn.execute(
        r#"
        INSERT INTO cells (sheet_id, row, col, value_type, value_number)
        VALUES (?1, 0, 2, 'number', 7.0)
        "#,
        rusqlite::params![sheet.id.to_string()],
    )
    .expect("insert scalar cell");

    // Insert a row/col value with an invalid type; the row should be skipped rather than causing
    // the query iterator to fail.
    conn.execute(
        r#"
        INSERT INTO cells (sheet_id, row, col, value_type, value_number)
        VALUES (?1, 'bogus', 0, 'number', 1.0)
        "#,
        rusqlite::params![sheet.id.to_string()],
    )
    .expect("insert invalid row type cell");

    let cells = storage
        .load_cells_in_range(sheet.id, CellRange::new(-1, 1, -1, 2))
        .expect("load cells");

    assert_eq!(cells.len(), 3, "invalid row type cell should be skipped");

    let cell_00 = cells
        .iter()
        .find(|(coord, _)| *coord == (0, 0))
        .expect("cell 0,0 exists");
    assert_eq!(cell_00.1.value, CellValue::Number(0.0));

    let cell_01 = cells
        .iter()
        .find(|(coord, _)| *coord == (0, 1))
        .expect("cell 0,1 exists");
    assert_eq!(cell_01.1.value, CellValue::String(String::new()));

    let cell_02 = cells
        .iter()
        .find(|(coord, _)| *coord == (0, 2))
        .expect("cell 0,2 exists");
    assert_eq!(cell_02.1.value, CellValue::Number(7.0));

    let exported = storage
        .export_model_workbook(workbook.id)
        .expect("export workbook");
    let exported_sheet = exported
        .sheets
        .iter()
        .find(|s| s.name == "Sheet1")
        .expect("sheet exists");

    assert_eq!(
        exported_sheet
            .cell(CellRef::new(0, 0))
            .expect("exported cell 0,0")
            .value,
        CellValue::Number(0.0)
    );
    assert_eq!(
        exported_sheet
            .cell(CellRef::new(0, 1))
            .expect("exported cell 0,1")
            .value,
        CellValue::String(String::new())
    );
    assert_eq!(
        exported_sheet
            .cell(CellRef::new(0, 2))
            .expect("exported cell 0,2")
            .value,
        CellValue::Number(7.0)
    );
}

