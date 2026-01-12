use std::collections::BTreeMap;

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
fn rich_entity_serializes_with_entity_type_id_and_properties() {
    let entity_value = CellValue::Entity(
        EntityValue::new("Apple Inc.")
            .with_entity_type("stock")
            .with_entity_id("AAPL")
            .with_property("Price", 178.5),
    );

    let json = serde_json::to_value(&entity_value).unwrap();
    assert_eq!(json["type"], "entity");
    assert_eq!(json["value"]["displayValue"], "Apple Inc.");
    assert_eq!(json["value"]["entityType"], "stock");
    assert_eq!(json["value"]["entityId"], "AAPL");
    assert_eq!(
        json["value"]["properties"]["Price"],
        serde_json::json!({ "type": "number", "value": 178.5 })
    );
}

#[test]
fn rich_record_serializes_fields_and_display_field_without_display_value() {
    let record_value = CellValue::Record(
        RecordValue::new("")
            .with_field("name", "Ada")
            .with_display_field("name"),
    );

    let json = serde_json::to_value(&record_value).unwrap();
    assert_eq!(json["type"], "record");
    assert_eq!(json["value"]["displayField"], "name");
    assert_eq!(
        json["value"]["fields"]["name"],
        serde_json::json!({ "type": "string", "value": "Ada" })
    );

    let value_obj = json["value"]
        .as_object()
        .expect("record value should serialize as object");
    assert!(
        value_obj.get("displayValue").is_none(),
        "displayValue should be omitted when empty"
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

    let mut fields = BTreeMap::new();
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

#[test]
fn record_json_without_display_value_deserializes_and_uses_display_field_for_display() {
    let json = r#"
    {
      "type": "record",
      "value": {
        "fields": {
          "name": { "type": "string", "value": "Ada" }
        },
        "displayField": "name"
      }
    }
    "#;

    let value: CellValue = serde_json::from_str(json).unwrap();
    let options = FormatOptions::default();
    let display = format_cell_display(&value, None, &options);
    assert_eq!(display.text, "Ada");
}

#[test]
fn record_json_display_field_can_point_to_entity() {
    let json = r#"
    {
      "type": "record",
      "value": {
        "fields": {
          "company": {
            "type": "entity",
            "value": { "displayValue": "Apple" }
          }
        },
        "displayField": "company"
      }
    }
    "#;

    let value: CellValue = serde_json::from_str(json).unwrap();
    let options = FormatOptions::default();
    let display = format_cell_display(&value, None, &options);
    assert_eq!(display.text, "Apple");

    match &value {
        CellValue::Record(record) => assert_eq!(record.to_string(), "Apple"),
        other => panic!("expected record, got {other:?}"),
    }
}

#[test]
fn record_json_display_field_can_point_to_nested_record() {
    let json = r#"
    {
      "type": "record",
      "value": {
        "fields": {
          "person": {
            "type": "record",
            "value": {
              "fields": {
                "name": { "type": "string", "value": "Ada" }
              },
              "displayField": "name"
            }
          }
        },
        "displayField": "person"
      }
    }
    "#;

    let value: CellValue = serde_json::from_str(json).unwrap();
    let options = FormatOptions::default();
    let display = format_cell_display(&value, None, &options);
    assert_eq!(display.text, "Ada");

    match &value {
        CellValue::Record(record) => assert_eq!(record.to_string(), "Ada"),
        other => panic!("expected record, got {other:?}"),
    }
}

#[test]
fn record_json_display_field_can_point_to_image_alt_text() {
    let json = r#"
    {
      "type": "record",
      "value": {
        "fields": {
          "logo": {
            "type": "image",
            "value": { "imageId": "logo.png", "altText": "Logo" }
          }
        },
        "displayField": "logo"
      }
    }
    "#;

    let value: CellValue = serde_json::from_str(json).unwrap();
    let options = FormatOptions::default();
    let display = format_cell_display(&value, None, &options);
    assert_eq!(display.text, "Logo");

    match &value {
        CellValue::Record(record) => assert_eq!(record.to_string(), "Logo"),
        other => panic!("expected record, got {other:?}"),
    }
}

#[test]
fn record_json_display_field_can_point_to_image_without_alt_text() {
    let json = r#"
    {
      "type": "record",
      "value": {
        "fields": {
          "logo": {
            "type": "image",
            "value": { "imageId": "logo.png" }
          }
        },
        "displayField": "logo"
      }
    }
    "#;

    let value: CellValue = serde_json::from_str(json).unwrap();
    let options = FormatOptions::default();
    let display = format_cell_display(&value, None, &options);
    assert_eq!(display.text, "[Image]");

    match &value {
        CellValue::Record(record) => assert_eq!(record.to_string(), "[Image]"),
        other => panic!("expected record, got {other:?}"),
    }
}
