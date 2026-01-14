use std::path::Path;

use formula_engine::{Engine, Value};
use formula_xlsx::load_from_path;
use pretty_assertions::assert_eq;

#[test]
fn imported_column_metadata_drives_cell_width_excel_encoding() {
    let fixture = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/xlsx/metadata/cell-width.xlsx");
    let doc = load_from_path(&fixture).expect("load fixture");
    let sheet = doc.workbook.sheet_by_name("Sheet1").expect("sheet exists");

    // The fixture is constructed to cover:
    // - a sheet default column width (20)
    // - a per-column width override (column B -> 25)
    // - a hidden column (column C)
    assert_eq!(sheet.default_col_width, Some(20.0));

    let col_b = sheet.col_properties(1).expect("col B should have props");
    assert_eq!(col_b.width, Some(25.0));
    assert!(!col_b.hidden);

    let col_c = sheet.col_properties(2).expect("col C should have props");
    assert_eq!(col_c.width, None);
    assert!(col_c.hidden);

    // Bridge the imported metadata into the formula engine and assert that `CELL("width")`
    // matches Excel's numeric encoding rules.
    let mut engine = Engine::new();
    engine.set_sheet_default_col_width("Sheet1", sheet.default_col_width);

    engine.set_col_width("Sheet1", 1, col_b.width);
    engine.set_col_hidden("Sheet1", 1, col_b.hidden);

    engine.set_col_width("Sheet1", 2, col_c.width);
    engine.set_col_hidden("Sheet1", 2, col_c.hidden);

    engine
        .set_cell_formula("Sheet1", "D1", "=CELL(\"width\",A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D2", "=CELL(\"width\",B1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D3", "=CELL(\"width\",C1)")
        .unwrap();
    engine.recalculate_single_threaded();

    // Sheet default width uses the `.0` marker.
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(20.0));
    // Explicit per-column width uses the `.1` marker.
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(25.1));
    // Hidden columns always return 0.
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(0.0));
}

