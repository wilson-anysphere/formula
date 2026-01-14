use std::collections::HashMap;

use formula_engine::value::{EntityValue, RecordValue};
use formula_engine::{Engine, ErrorKind, Value};
use pretty_assertions::assert_eq;

#[test]
fn entity_field_access_variants() {
    let mut engine = Engine::new();

    let mut fields = HashMap::new();
    fields.insert("Price".to_string(), Value::Number(178.5));
    fields.insert("Change%".to_string(), Value::Number(0.0133));

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Entity(EntityValue {
                display: "Apple".to_string(),
                entity_type: Some("Stock".to_string()),
                entity_id: Some("AAPL".to_string()),
                fields,
            }),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=A1.Price")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", r#"=A1.["Change%"]"#)
        .unwrap();
    engine.set_cell_formula("Sheet1", "D1", "=A1.Nope").unwrap();

    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=A2.Price")
        .unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(178.5));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(0.0133));
    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Error(ErrorKind::Field)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn nested_field_access_entity_to_record() {
    let mut engine = Engine::new();

    let mut address_fields = HashMap::new();
    address_fields.insert("City".to_string(), Value::Text("Seattle".to_string()));

    let address = Value::Record(RecordValue {
        display: "Seattle".to_string(),
        fields: address_fields,
        display_field: Some("City".to_string()),
    });

    let mut fields = HashMap::new();
    fields.insert("Address".to_string(), address);

    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Entity(EntityValue {
                display: "Alice".to_string(),
                entity_type: Some("Person".to_string()),
                entity_id: Some("123".to_string()),
                fields,
            }),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=A1.Address.City")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("Seattle".to_string())
    );
}
