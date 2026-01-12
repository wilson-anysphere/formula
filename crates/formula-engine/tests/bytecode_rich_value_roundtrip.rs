#![cfg(not(target_arch = "wasm32"))]

use formula_engine::value::EntityValue;
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn bytecode_backend_preserves_entity_values_through_cell_ref_pass_through() {
    let mut engine = Engine::new();

    engine
        .set_cell_value("Sheet1", "A1", Value::Entity(EntityValue::new("Widget")))
        .unwrap();

    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    // Field access lowers to an internal `_FIELDACCESS` call in the AST evaluator and is not
    // currently supported by the bytecode backend, so this formula should fall back to the AST
    // path while still reading the (bytecode-evaluated) `B1` result.
    engine.set_cell_formula("Sheet1", "C1", "=B1.Price").unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 1);
    assert_eq!(stats.fallback, 1);

    engine.recalculate_single_threaded();

    match engine.get_cell_value("Sheet1", "B1") {
        Value::Entity(ent) => assert_eq!(ent.display, "Widget"),
        other => panic!("expected B1 to remain an Entity, got {other:?}"),
    }

    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Error(ErrorKind::Field)
    );
}

