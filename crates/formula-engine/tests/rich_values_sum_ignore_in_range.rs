use formula_engine::value::{EntityValue, RecordValue};
use formula_engine::{Engine, Value};

#[test]
fn sum_ignores_entity_in_range_bytecode() {
    let mut engine = Engine::new();
    // Bytecode is enabled by default; assert we actually compile at least one program so this
    // test exercises the SIMD column-slice cache population path.
    engine.set_bytecode_enabled(true);

    engine.set_cell_value("Sheet1", "A1", 1.0).expect("set A1");
    engine
        .set_cell_value(
            "Sheet1",
            "A2",
            Value::Entity(EntityValue::new("SomeEntity")),
        )
        .expect("set A2");
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A2)")
        .expect("set B1 formula");

    // Ensure the formula was eligible for bytecode compilation.
    assert!(
        engine.bytecode_program_count() > 0,
        "expected SUM range formula to compile to bytecode"
    );

    engine.recalculate_single_threaded();

    match engine.get_cell_value("Sheet1", "B1") {
        Value::Number(n) => assert!((n - 1.0).abs() < 1e-9, "expected 1, got {n}"),
        other => panic!("expected number, got {other:?}"),
    }
}

#[test]
fn sum_ignores_record_in_range_bytecode() {
    let mut engine = Engine::new();
    engine.set_bytecode_enabled(true);

    engine.set_cell_value("Sheet1", "A1", 1.0).expect("set A1");
    engine
        .set_cell_value(
            "Sheet1",
            "A2",
            Value::Record(RecordValue::new("SomeRecord")),
        )
        .expect("set A2");
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A2)")
        .expect("set B1 formula");

    assert!(
        engine.bytecode_program_count() > 0,
        "expected SUM range formula to compile to bytecode"
    );

    engine.recalculate_single_threaded();

    match engine.get_cell_value("Sheet1", "B1") {
        Value::Number(n) => assert!((n - 1.0).abs() < 1e-9, "expected 1, got {n}"),
        other => panic!("expected number, got {other:?}"),
    }
}
