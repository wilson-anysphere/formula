use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn column_whole_column_reference_scalarizes_with_custom_sheet_rows() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 16_384)
        .expect("set sheet dimensions");

    engine
        // Avoid a circular dependency by placing the formula outside the referenced column.
        .set_cell_formula("Sheet1", "B1", "=COLUMN(A:A)")
        .expect("set formula");
    engine.recalculate_single_threaded();

    // COLUMN(A:A) should evaluate to a scalar, not a spilled 10x1 array.
    assert!(engine.spill_range("Sheet1", "B1").is_none());
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
}

#[test]
fn offset_returns_ref_error_when_result_is_outside_custom_sheet_bounds() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 16_384)
        .expect("set sheet dimensions");

    engine
        .set_cell_value("Sheet1", "A1", 1.0)
        .expect("set value");
    engine
        .set_cell_formula("Sheet1", "B1", "=OFFSET(A1,10,0)")
        .expect("set formula");
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Error(ErrorKind::Ref)
    );
}

#[test]
fn indirect_expands_whole_column_refs_using_custom_sheet_dimensions() {
    let mut engine = Engine::new();
    engine
        .set_sheet_dimensions("Sheet1", 10, 16_384)
        .expect("set sheet dimensions");

    // Avoid a circular dependency by placing the formula outside the referenced column.
    engine
        .set_cell_formula("Sheet1", "B1", "=ROWS(INDIRECT(\"A:A\"))")
        .expect("set formula");
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(10.0));
}
