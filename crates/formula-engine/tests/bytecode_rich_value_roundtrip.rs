#![cfg(not(target_arch = "wasm32"))]

use formula_engine::value::{EntityValue, RecordValue};
use formula_engine::{Engine, Value};
use std::collections::HashMap;

#[test]
fn bytecode_backend_preserves_entity_values_through_cell_ref_pass_through() {
    let mut engine = Engine::new();

    let mut fields = HashMap::new();
    fields.insert("Price".to_string(), Value::Number(9.99));

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Entity(EntityValue::with_fields("Widget", fields)),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    // Field access is lowered to an internal `_FIELDACCESS` builtin. Ensure the bytecode backend
    // can evaluate the lowered form and preserve rich values across cells.
    engine
        .set_cell_formula("Sheet1", "C1", "=B1.Price")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 2);
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    match engine.get_cell_value("Sheet1", "B1") {
        Value::Entity(ent) => {
            assert_eq!(ent.display, "Widget");
            assert_eq!(ent.fields.get("Price"), Some(&Value::Number(9.99)));
        }
        other => panic!("expected B1 to remain an Entity, got {other:?}"),
    }

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(9.99));
}

#[test]
fn bytecode_backend_preserves_record_values_through_cell_ref_pass_through() {
    let mut engine = Engine::new();

    let mut fields = HashMap::new();
    fields.insert("Price".to_string(), Value::Number(123.0));

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Record(RecordValue::with_fields("Widget record", fields)),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "B1", "=A1").unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine
        .set_cell_formula("Sheet1", "C1", "=B1.Price")
        .unwrap();

    let stats = engine.bytecode_compile_stats();
    assert_eq!(stats.total_formula_cells, 2);
    assert_eq!(stats.compiled, 2);
    assert_eq!(stats.fallback, 0);

    engine.recalculate_single_threaded();

    match engine.get_cell_value("Sheet1", "B1") {
        Value::Record(rec) => {
            assert_eq!(rec.display, "Widget record");
            assert_eq!(rec.fields.get("Price"), Some(&Value::Number(123.0)));
        }
        other => panic!("expected B1 to remain a Record, got {other:?}"),
    }

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(123.0));
}
