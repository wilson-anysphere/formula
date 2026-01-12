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
