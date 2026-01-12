use formula_engine::{Engine, Value};
use formula_model::EXCEL_MAX_COLS;

#[test]
fn sheet_dimensions_affect_full_column_rows() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 2_000_000, EXCEL_MAX_COLS)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=ROWS(A:A)")
        .unwrap();

    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Number(2_000_000.0)
    );

    // Updating sheet dimensions should mark the cell dirty and recompute.
    engine
        .set_sheet_dimensions("Sheet1", 3_000_000, EXCEL_MAX_COLS)
        .unwrap();
    assert!(engine.is_dirty("Sheet1", "B1"));
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Number(3_000_000.0)
    );
}

#[test]
fn sheet_dimensions_affect_full_row_columns() {
    let mut engine = Engine::new();
    engine.set_sheet_dimensions("Sheet1", 10, 100).unwrap();
    // Avoid a circular reference: `1:1` includes row 1, so the formula can't be on row 1.
    engine
        .set_cell_formula("Sheet1", "B2", "=COLUMNS(1:1)")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(100.0));

    // Updating sheet dimensions should mark the cell dirty and recompute.
    engine.set_sheet_dimensions("Sheet1", 10, 120).unwrap();
    assert!(engine.is_dirty("Sheet1", "B2"));
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(120.0));
}

#[test]
fn row_and_column_handle_row_and_column_refs_with_custom_dimensions() {
    let mut engine = Engine::new();
    engine.set_sheet_dimensions("Sheet1", 2_000_000, 100).unwrap();

    // Avoid a circular ref: `5:7` includes row 5..7, so keep the formula outside those rows.
    engine
        .set_cell_formula("Sheet1", "J1", "=ROW(5:7)")
        .unwrap();
    // Avoid a circular ref: `D:F` includes column D..F, so keep the formula in another column.
    engine
        .set_cell_formula("Sheet1", "A10", "=COLUMN(D:F)")
        .unwrap();

    engine.recalculate();

    // ROW(5:7) -> {5;6;7}
    assert_eq!(engine.get_cell_value("Sheet1", "J1"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J2"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J3"), Value::Number(7.0));

    // COLUMN(D:F) -> {4,5,6}
    assert_eq!(engine.get_cell_value("Sheet1", "A10"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B10"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C10"), Value::Number(6.0));
}
