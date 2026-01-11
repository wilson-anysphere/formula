use formula_storage::{CellChange, CellData, CellRange, CellValue, Storage};
use rusqlite::{Connection, OpenFlags};

#[test]
fn rename_sheet_skips_corrupt_formula_cells() {
    let uri = "file:rename_sheet_invalid_cell_types?mode=memory&cache=shared";
    let storage = Storage::open_uri(uri).expect("open storage");
    let workbook = storage
        .create_workbook("Book", None)
        .expect("create workbook");
    let data_sheet = storage
        .create_sheet(workbook.id, "Data", 0, None)
        .expect("create data sheet");
    let summary_sheet = storage
        .create_sheet(workbook.id, "Summary", 1, None)
        .expect("create summary sheet");

    storage
        .apply_cell_changes(&[CellChange {
            sheet_id: summary_sheet.id,
            row: 0,
            col: 0,
            data: CellData {
                value: CellValue::Empty,
                formula: Some("Data!A1".to_string()),
                style: None,
            },
            user_id: None,
        }])
        .expect("set formula");

    let flags =
        OpenFlags::SQLITE_OPEN_READ_WRITE | OpenFlags::SQLITE_OPEN_CREATE | OpenFlags::SQLITE_OPEN_URI;
    let conn = Connection::open_with_flags(uri, flags).expect("open raw connection");

    // Insert a corrupt formula cell with an invalid row type; rename should ignore it rather than failing.
    conn.execute(
        r#"
        INSERT INTO cells (sheet_id, row, col, formula)
        VALUES (?1, 'bogus', 0, 'Data!A1')
        "#,
        rusqlite::params![summary_sheet.id.to_string()],
    )
    .expect("insert corrupt cell");

    storage
        .rename_sheet(data_sheet.id, "Renamed")
        .expect("rename sheet");

    let cells = storage
        .load_cells_in_range(summary_sheet.id, CellRange::new(0, 0, 0, 0))
        .expect("load cells");
    assert_eq!(cells.len(), 1);
    assert_eq!(cells[0].1.formula.as_deref(), Some("Renamed!A1"));
}

