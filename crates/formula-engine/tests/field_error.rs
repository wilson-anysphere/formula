use formula_engine::{Engine, ErrorKind, Value};
use pretty_assertions::assert_eq;

#[test]
fn parses_field_error_literal_and_propagates() {
    let mut engine = Engine::new();

    engine.set_cell_formula("Sheet1", "A1", "=#FIELD!").unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=#FIELD!+1")
        .unwrap();

    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Field)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::Field)
    );
}

#[test]
fn field_error_propagates_through_bytecode_formulas() {
    let mut engine = Engine::new();
    engine.set_bytecode_enabled(true);

    engine
        .set_cell_value("Sheet1", "A1", Value::Error(ErrorKind::Field))
        .unwrap();
    engine.set_cell_formula("Sheet1", "B1", "=A1+1").unwrap();

    // Ensure the bytecode compiler accepted the formula so this test exercises the
    // engine<->bytecode error-kind mappings.
    assert!(engine.bytecode_program_count() > 0);

    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Error(ErrorKind::Field)
    );
}
