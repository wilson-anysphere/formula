use super::*;

use serde_json::json;

#[test]
fn pivot_config_serde_roundtrips_with_calculated_fields_and_items() {
    let allowed = vec![PivotKeyPart::Text("East".to_string())];

    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![PivotField::new("Product")],
        value_fields: vec![ValueField {
            source_field: "Sales".to_string(),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![FilterField {
            source_field: "Region".to_string(),
            allowed: Some(allowed),
        }],
        calculated_fields: vec![CalculatedField {
            name: "Profit".to_string(),
            formula: "Sales - Cost".to_string(),
        }],
        calculated_items: vec![CalculatedItem {
            field: "Region".to_string(),
            name: "East+West".to_string(),
            formula: "\"East\" + \"West\"".to_string(),
        }],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals::default(),
    };

    let json_value = serde_json::to_value(&cfg).unwrap();
    assert!(json_value.get("calculatedFields").is_some());
    assert!(json_value.get("calculatedItems").is_some());

    let decoded: PivotConfig = serde_json::from_value(json_value.clone()).unwrap();
    assert_eq!(decoded, cfg);

    // Backward-compat: missing keys should default to empty vectors.
    let mut json_without = json_value;
    if let Some(obj) = json_without.as_object_mut() {
        obj.remove("calculatedFields");
        obj.remove("calculatedItems");
    }
    let decoded: PivotConfig = serde_json::from_value(json_without).unwrap();
    assert!(decoded.calculated_fields.is_empty());
    assert!(decoded.calculated_items.is_empty());
    assert_eq!(decoded.row_fields, cfg.row_fields);
    assert_eq!(decoded.column_fields, cfg.column_fields);
    assert_eq!(decoded.value_fields, cfg.value_fields);
    assert_eq!(decoded.filter_fields, cfg.filter_fields);
    assert_eq!(decoded.layout, cfg.layout);
    assert_eq!(decoded.subtotals, cfg.subtotals);
    assert_eq!(decoded.grand_totals, cfg.grand_totals);
}

#[test]
fn pivot_value_to_key_part_canonicalizes_numbers() {
    assert_eq!(
        PivotValue::Number(0.0).to_key_part(),
        PivotValue::Number(-0.0).to_key_part()
    );

    let alt_nan = f64::from_bits(0x7ff8_0000_0000_0001);
    assert!(alt_nan.is_nan());
    assert_eq!(
        PivotValue::Number(f64::NAN).to_key_part(),
        PivotValue::Number(alt_nan).to_key_part()
    );
    assert_eq!(
        PivotValue::Number(alt_nan).to_key_part(),
        PivotKeyPart::Number(f64::NAN.to_bits())
    );

    let n = 12.5;
    assert_eq!(
        PivotValue::Number(n).to_key_part(),
        PivotKeyPart::Number(n.to_bits())
    );
}

#[test]
fn pivot_value_uses_tagged_ipc_serde_layout() {
    assert_eq!(serde_json::to_value(&PivotValue::Blank).unwrap(), json!({"type": "blank"}));
    assert_eq!(
        serde_json::to_value(&PivotValue::Text("hi".to_string())).unwrap(),
        json!({"type": "text", "value": "hi"})
    );
    assert_eq!(
        serde_json::to_value(&PivotValue::Number(1.5)).unwrap(),
        json!({"type": "number", "value": 1.5})
    );
}

#[test]
fn pivot_key_part_display_string_matches_excel_like_expectations() {
    assert_eq!(PivotKeyPart::Blank.display_string(), "(blank)");
    assert_eq!(
        PivotKeyPart::Number(PivotValue::canonical_number_bits(-0.0)).display_string(),
        "0"
    );
}
