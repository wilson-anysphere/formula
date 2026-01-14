use formula_engine::value::{EntityValue, RecordValue};
use formula_engine::{Engine, Value};

fn assert_rich_values_compare_and_sort_like_text(mut engine: Engine) {
    engine
        .set_cell_value("Sheet1", "A1", Value::Entity(EntityValue::new("Apple")))
        .unwrap();
    // Excel comparisons are case-insensitive; entities/records should behave the same by using
    // their display strings.
    engine.set_cell_value("Sheet1", "B1", "apple").unwrap();
    engine
        .set_cell_value("Sheet1", "D1", Value::Record(RecordValue::new("Apple")))
        .unwrap();

    engine.set_cell_formula("Sheet1", "C1", "=A1=B1").unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", r#"=A1>"Aardvark""#)
        .unwrap();
    engine.set_cell_formula("Sheet1", "E1", "=D1=B1").unwrap();

    // Validate lookup binary-search semantics as well: type precedence should treat entities
    // as text-like values between numbers and booleans.
    engine.set_cell_value("Sheet1", "A2", true).unwrap();
    engine
        .set_cell_formula("Sheet1", "C3", "=XMATCH(TRUE, A1:A2, 0, 2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C4", r#"=XMATCH("APPLE", A1:A2, 0, 2)"#)
        .unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C4"), Value::Number(1.0));
}

#[test]
fn rich_values_compare_and_sort_like_text_bytecode() {
    assert_rich_values_compare_and_sort_like_text(Engine::new());
}

#[test]
fn rich_values_compare_and_sort_like_text_ast() {
    let mut engine = Engine::new();
    engine.set_bytecode_enabled(false);
    assert_rich_values_compare_and_sort_like_text(engine);
}
