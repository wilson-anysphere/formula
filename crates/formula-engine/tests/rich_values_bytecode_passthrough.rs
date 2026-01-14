use formula_engine::value::EntityValue;
use formula_engine::{Engine, Value};

#[test]
fn rich_values_roundtrip_through_bytecode_and_enable_chaining() {
    // Bytecode is enabled by default in `Engine::new()`. This test ensures that a rich value
    // loaded from a cell reference can be returned from a bytecode program without degrading
    // to text, so downstream formulas can still access fields/properties.
    let mut engine = Engine::new();

    let entity = EntityValue::new("Widget").property("Price", Value::Number(42.0));
    engine
        .set_cell_value("Sheet1", "A1", Value::Entity(entity.clone()))
        .expect("set A1");

    engine
        .set_cell_formula("Sheet1", "B1", "=A1")
        .expect("set B1 formula");
    engine.recalculate_single_threaded();

    assert!(
        engine.bytecode_program_count() >= 1,
        "expected at least one compiled bytecode program"
    );

    let b1 = engine.get_cell_value("Sheet1", "B1");
    assert!(
        matches!(b1, Value::Entity(_)),
        "expected B1 to be an entity"
    );
    assert_eq!(b1, Value::Entity(entity.clone()));

    engine
        .set_cell_formula("Sheet1", "C1", "=B1.Price")
        .expect("set C1 formula");
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(42.0));
}
