use std::collections::HashMap;

use formula_engine::value::RecordValue;
use formula_engine::{Engine, Value};

fn assert_t_and_switch_use_record_display_field(mut engine: Engine) {
    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Record(RecordValue {
                display: "Fallback".to_string(),
                // Deliberately use different casing than the stored field key to validate
                // case-insensitive display_field resolution.
                display_field: Some("name".to_string()),
                fields: HashMap::from([("Name".to_string(), Value::Text("Apple".to_string()))]),
            }),
        )
        .unwrap();

    engine.set_cell_formula("Sheet1", "B1", "=T(A1)").unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", r#"=SWITCH(A1,"Apple",1,0)"#)
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("Apple".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(1.0));
}

#[test]
fn record_display_field_parity_bytecode() {
    assert_t_and_switch_use_record_display_field(Engine::new());
}

#[test]
fn record_display_field_parity_ast() {
    let mut engine = Engine::new();
    engine.set_bytecode_enabled(false);
    assert_t_and_switch_use_record_display_field(engine);
}
