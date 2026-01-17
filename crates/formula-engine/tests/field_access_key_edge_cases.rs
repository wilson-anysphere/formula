use std::collections::HashMap;

use formula_engine::value::EntityValue;
use formula_engine::{Engine, Value};
use pretty_assertions::assert_eq;

#[test]
fn field_access_does_not_strip_brackets_from_key() {
    let mut engine = Engine::new();

    let mut fields = HashMap::new();
    fields.insert("Price".to_string(), Value::Number(1.0));
    fields.insert("[Price]".to_string(), Value::Number(2.0));

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Entity(EntityValue {
                display: "Example".to_string(),
                entity_type: None,
                entity_id: None,
                fields,
            }),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", r#"=A1.["[Price]"]"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
}

#[test]
fn field_access_does_not_trim_spaces_from_key() {
    let mut engine = Engine::new();

    let mut fields = HashMap::new();
    fields.insert("Price".to_string(), Value::Number(1.0));
    fields.insert(" Price ".to_string(), Value::Number(3.0));

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Entity(EntityValue {
                display: "Example".to_string(),
                entity_type: None,
                entity_id: None,
                fields,
            }),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", r#"=A1.[" Price "]"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
}

#[test]
fn field_access_unescapes_doubled_quotes_in_bracket_selectors() {
    let mut engine = Engine::new();

    let mut fields = HashMap::new();
    fields.insert("He said \"Hi\"".to_string(), Value::Number(4.0));

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Entity(EntityValue {
                display: "Example".to_string(),
                entity_type: None,
                entity_id: None,
                fields,
            }),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", r#"=A1.["He said ""Hi"""]"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(4.0));
}

#[test]
fn field_access_allows_rbracket_in_quoted_key() {
    let mut engine = Engine::new();

    let mut fields = HashMap::new();
    fields.insert("Price]".to_string(), Value::Number(7.0));

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Entity(EntityValue {
                display: "Example".to_string(),
                entity_type: None,
                entity_id: None,
                fields,
            }),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", r#"=A1.["Price]"]"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(7.0));
}
