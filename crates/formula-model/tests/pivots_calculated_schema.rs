use formula_model::pivots::{CalculatedField, CalculatedItem};

use serde_json::json;

#[test]
fn calculated_field_json_round_trip() {
    let field = CalculatedField {
        name: "Profit".to_string(),
        formula: "=Sales-Cost".to_string(),
    };

    let serialized = serde_json::to_value(&field).unwrap();
    assert_eq!(
        serialized,
        json!({
            "name": "Profit",
            "formula": "=Sales-Cost",
        })
    );

    let deserialized: CalculatedField = serde_json::from_value(serialized).unwrap();
    assert_eq!(deserialized, field);
}

#[test]
fn calculated_item_json_round_trip() {
    let item = CalculatedItem {
        field: "Month".to_string(),
        name: "Q1".to_string(),
        formula: "=Jan+Feb+Mar".to_string(),
    };

    let serialized = serde_json::to_value(&item).unwrap();
    assert_eq!(
        serialized,
        json!({
            "field": "Month",
            "name": "Q1",
            "formula": "=Jan+Feb+Mar",
        })
    );

    let deserialized: CalculatedItem = serde_json::from_value(serialized).unwrap();
    assert_eq!(deserialized, item);
}
