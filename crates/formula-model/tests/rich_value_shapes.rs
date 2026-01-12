use std::collections::HashMap;

use formula_format::FormatOptions;
use formula_model::{format_cell_display, CellValue, EntityValue, RecordValue};

#[test]
fn old_entity_and_record_json_with_only_display_deserializes() {
    let entity_json = r#"{ "type": "entity", "value": { "display": "X" } }"#;
    let record_json = r#"{ "type": "record", "value": { "display": "Y" } }"#;

    let entity: CellValue = serde_json::from_str(entity_json).unwrap();
    let record: CellValue = serde_json::from_str(record_json).unwrap();

    match entity {
        CellValue::Entity(e) => {
            assert_eq!(e.display_value, "X");
            assert_eq!(e.entity_type, "");
            assert_eq!(e.entity_id, "");
            assert!(e.properties.is_empty());
        }
        other => panic!("expected entity, got {other:?}"),
    }

    match record {
        CellValue::Record(r) => {
            assert_eq!(r.display_value, "Y");
            assert!(r.fields.is_empty());
            assert_eq!(r.display_field, None);
        }
        other => panic!("expected record, got {other:?}"),
    }
}

#[test]
fn entity_and_record_serialize_compactly_when_only_display_is_present() {
    let entity = CellValue::Entity(EntityValue::new("X"));
    let record = CellValue::Record(RecordValue::new("Y"));

    let entity_json = serde_json::to_value(&entity).unwrap();
    let record_json = serde_json::to_value(&record).unwrap();

    assert_eq!(
        entity_json,
        serde_json::json!({ "type": "entity", "value": { "displayValue": "X" } })
    );
    assert_eq!(
        record_json,
        serde_json::json!({ "type": "record", "value": { "displayValue": "Y" } })
    );
}

#[test]
fn rich_entity_and_record_json_roundtrip() {
    let entity = EntityValue::new("Apple Inc.")
        .with_entity_type("stock")
        .with_entity_id("AAPL")
        .with_property("Price", 178.50)
        .with_property("Change", 2.35);
    let entity_value = CellValue::Entity(entity);

    let mut fields = HashMap::new();
    fields.insert("Name".to_string(), CellValue::String("Alice".to_string()));
    fields.insert("Age".to_string(), CellValue::Number(42.0));
    let record = RecordValue::new("Alice")
        .with_fields(fields)
        .with_display_field("Name");
    let record_value = CellValue::Record(record);

    let entity_json = serde_json::to_string(&entity_value).unwrap();
    let record_json = serde_json::to_string(&record_value).unwrap();

    let entity_roundtrip: CellValue = serde_json::from_str(&entity_json).unwrap();
    let record_roundtrip: CellValue = serde_json::from_str(&record_json).unwrap();

    assert_eq!(entity_roundtrip, entity_value);
    assert_eq!(record_roundtrip, record_value);
}

#[test]
fn format_cell_display_for_entity_and_record_uses_display_string() {
    let options = FormatOptions::default();

    let entity = CellValue::Entity(
        EntityValue::new("Entity Display")
            .with_entity_type("stock")
            .with_entity_id("AAPL")
            .with_property("Price", 178.50),
    );
    let record = CellValue::Record(RecordValue::new("Record Display").with_field("Price", 178.50));

    let entity_display = format_cell_display(&entity, None, &options);
    let record_display = format_cell_display(&record, None, &options);

    assert_eq!(entity_display.text, "Entity Display");
    assert_eq!(record_display.text, "Record Display");
}
