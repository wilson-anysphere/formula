use formula_engine::{Engine, ErrorKind, Value};
use formula_model::EXCEL_MAX_COLS;

#[test]
fn out_of_bounds_cell_reference_evaluates_to_ref_error() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=A2000000")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn growing_sheet_dimensions_makes_far_references_valid() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=A2000000")
        .unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Error(ErrorKind::Ref)
    );

    // Grow the sheet by writing a value in a far row (but not in A2000000 itself).
    engine.set_cell_value("Sheet1", "B2000000", 1.0).unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "A2000000"), Value::Blank);
    engine.recalculate();

    // The referenced cell is now in-bounds; since it's unset it behaves as blank.
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Blank);

    // Setting the referenced cell should surface its value.
    engine.set_cell_value("Sheet1", "A2000000", 42.0).unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(42.0));
}

#[test]
fn out_of_bounds_range_reference_evaluates_to_ref_error() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A2000000)")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn sheet_dimensions_expand_whole_column_references() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 2_100_000, EXCEL_MAX_COLS)
        .unwrap();
    engine.set_cell_value("Sheet1", "A2000000", 5.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A:A)")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(5.0));
}

#[test]
fn sheet_dimensions_shrink_whole_row_references() {
    let mut engine = Engine::new();
    // Limit the sheet to just 10 columns (A..J) so `1:1` should resolve to `A1:J1` rather than
    // erroring due to an out-of-bounds endpoint.
    engine.set_sheet_dimensions("Sheet1", 100, 10).unwrap();
    engine.set_cell_value("Sheet1", "J1", 7.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=SUM(1:1)")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(7.0));
}
