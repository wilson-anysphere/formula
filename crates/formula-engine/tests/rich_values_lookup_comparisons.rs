use formula_engine::{value::EntityValue, Engine, Value};

#[test]
fn match_compares_entity_and_text_by_display_string_case_insensitive() {
    let mut engine = Engine::new();
    // Ensure we exercise the AST evaluator path, which preserves rich Value variants.
    engine.set_bytecode_enabled(false);

    engine
        .set_cell_value("Sheet1", "A1", Value::Entity(EntityValue::new("Apple")))
        .unwrap();
    engine.set_cell_value("Sheet1", "B1", "apple").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=MATCH(B1, A1:A1, 0)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
}

#[test]
fn match_wildcards_compare_against_entity_display_string() {
    let mut engine = Engine::new();
    // Ensure we exercise the AST evaluator path, which preserves rich Value variants.
    engine.set_bytecode_enabled(false);

    engine
        .set_cell_value("Sheet1", "A1", Value::Entity(EntityValue::new("Apple")))
        .unwrap();
    engine.set_cell_value("Sheet1", "B1", "app*").unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=MATCH(B1, A1:A1, 0)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
}

#[test]
fn xlookup_compares_entity_and_text_by_display_string_case_insensitive() {
    let mut engine = Engine::new();
    // Ensure we exercise the AST evaluator path, which preserves rich Value variants.
    engine.set_bytecode_enabled(false);

    engine
        .set_cell_value("Sheet1", "A1", Value::Entity(EntityValue::new("Apple")))
        .unwrap();
    engine.set_cell_value("Sheet1", "B1", "apple").unwrap();
    engine.set_cell_value("Sheet1", "C1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=XLOOKUP(B1, A1:A1, C1:C1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
}
