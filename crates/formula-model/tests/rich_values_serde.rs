use std::collections::BTreeMap;

use formula_model::{CellValue, EntityValue, RecordValue};
use serde_json::json;

#[test]
fn entity_value_json_roundtrip() {
    let mut meta_fields = BTreeMap::new();
    meta_fields.insert("symbol".to_string(), CellValue::String("AAPL".to_string()));
    meta_fields.insert("active".to_string(), CellValue::Boolean(true));

    let mut properties = BTreeMap::new();
    properties.insert("Price".to_string(), CellValue::Number(178.5));
    properties.insert("Name".to_string(), CellValue::String("Apple".to_string()));
    properties.insert(
        "Meta".to_string(),
        CellValue::Record(RecordValue {
            fields: meta_fields,
            display_field: Some("symbol".to_string()),
            ..RecordValue::default()
        }),
    );

    let value = CellValue::Entity(EntityValue {
        entity_type: "stock".to_string(),
        entity_id: "AAPL".to_string(),
        display_value: "Apple Inc.".to_string(),
        properties,
    });

    let serialized = serde_json::to_value(&value).unwrap();
    let expected = json!({
        "type": "entity",
        "value": {
            "entityType": "stock",
            "entityId": "AAPL",
            "displayValue": "Apple Inc.",
            "properties": {
                "Meta": {
                    "type": "record",
                    "value": {
                        "fields": {
                            "active": { "type": "boolean", "value": true },
                            "symbol": { "type": "string", "value": "AAPL" }
                        },
                        "displayField": "symbol"
                    }
                },
                "Name": { "type": "string", "value": "Apple" },
                "Price": { "type": "number", "value": 178.5 }
            }
        }
    });

    assert_eq!(serialized, expected);
    let roundtrip: CellValue = serde_json::from_value(serialized).unwrap();
    assert_eq!(roundtrip, value);
}

#[test]
fn record_value_json_roundtrip() {
    let mut entity_props = BTreeMap::new();
    entity_props.insert("Country".to_string(), CellValue::String("USA".to_string()));
    let employer = CellValue::Entity(EntityValue {
        entity_type: "company".to_string(),
        entity_id: "AAPL".to_string(),
        display_value: "Apple".to_string(),
        properties: entity_props,
    });

    let mut fields = BTreeMap::new();
    fields.insert("Name".to_string(), CellValue::String("Alice".to_string()));
    fields.insert("Age".to_string(), CellValue::Number(30.0));
    fields.insert("Employer".to_string(), employer);

    let value = CellValue::Record(RecordValue {
        fields,
        display_field: Some("Name".to_string()),
        ..RecordValue::default()
    });

    let serialized = serde_json::to_value(&value).unwrap();
    let expected = json!({
        "type": "record",
        "value": {
            "fields": {
                "Age": { "type": "number", "value": 30.0 },
                "Employer": {
                    "type": "entity",
                    "value": {
                        "entityType": "company",
                        "entityId": "AAPL",
                        "displayValue": "Apple",
                        "properties": {
                            "Country": { "type": "string", "value": "USA" }
                        }
                    }
                },
                "Name": { "type": "string", "value": "Alice" }
            },
            "displayField": "Name"
        }
    });

    assert_eq!(serialized, expected);
    let roundtrip: CellValue = serde_json::from_value(serialized).unwrap();
    assert_eq!(roundtrip, value);
}
