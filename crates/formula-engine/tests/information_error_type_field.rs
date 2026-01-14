use formula_engine::value::{EntityValue, RecordValue};
use formula_engine::{Engine, Value};

fn assert_number(value: &Value, expected: f64) {
    match value {
        Value::Number(n) => assert!((*n - expected).abs() < 1e-9, "expected {expected}, got {n}"),
        other => panic!("expected number {expected}, got {other:?}"),
    }
}

#[test]
fn error_type_includes_field_error_code() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=ERROR.TYPE(#FIELD!)")
        .unwrap();
    engine.recalculate();
    assert_number(&engine.get_cell_value("Sheet1", "A1"), 11.0);
}

#[test]
fn type_treats_rich_values_as_text() {
    let mut engine = Engine::new();
    engine
        .set_cell_value("Sheet1", "A1", Value::Entity(EntityValue::new("entity")))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A2", Value::Record(RecordValue::new("record")))
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=TYPE(A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=TYPE(A2)")
        .unwrap();
    engine.recalculate();
    assert_number(&engine.get_cell_value("Sheet1", "B1"), 2.0);
    assert_number(&engine.get_cell_value("Sheet1", "B2"), 2.0);
}
