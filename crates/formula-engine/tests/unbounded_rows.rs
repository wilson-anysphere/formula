use formula_engine::{Engine, Value};
use pretty_assertions::assert_eq;

#[test]
fn formulas_can_reference_rows_beyond_excel_max() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A2000000", 7.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=A2000000")
        .unwrap();

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(7.0));
}

#[test]
fn whole_column_references_use_dynamic_sheet_dimensions() {
    let mut engine = Engine::new();

    // Writing beyond Excel's default bounds should grow the sheet, and full-column references
    // should include the newly addressable rows.
    engine.set_cell_value("Sheet1", "A2000000", 10.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A:A)")
        .unwrap();

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(10.0));
}

#[test]
fn whole_column_dependents_are_marked_dirty_when_writing_far_rows() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A:A)")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(0.0));

    engine.set_cell_value("Sheet1", "A2000000", 1.0).unwrap();
    assert!(engine.is_dirty("Sheet1", "B1"));

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
}

#[test]
fn indirect_can_resolve_rows_beyond_excel_max() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A2000000", 9.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", r#"=INDIRECT("A2000000")"#)
        .unwrap();

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(9.0));
}

