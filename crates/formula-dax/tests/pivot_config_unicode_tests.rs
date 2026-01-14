#![cfg(feature = "pivot-model")]

use formula_dax::{pivot_crosstab_from_config, DataModel, Table, Value};
use formula_model::pivots::{
    AggregationType, FilterField, PivotConfig, PivotField, PivotFieldRef, PivotKeyPart, SortOrder,
    ValueField,
};
use pretty_assertions::assert_eq;
use std::collections::HashSet;

#[test]
fn pivot_crosstab_from_config_resolves_unicode_identifiers_case_insensitively() {
    let mut model = DataModel::new();
    let mut fact = Table::new("Straße", vec!["StraßenId", "Region", "Maß"]);
    fact.push_row(vec![1.into(), "East".into(), 10.0.into()])
        .unwrap();
    fact.push_row(vec![1.into(), "East".into(), 20.0.into()])
        .unwrap();
    fact.push_row(vec![2.into(), "West".into(), 5.0.into()])
        .unwrap();
    model.add_table(fact).unwrap();

    let cfg = PivotConfig {
        row_fields: vec![PivotField {
            source_field: PivotFieldRef::DataModelColumn {
                table: "STRASSE".to_string(),
                column: "STRASSENID".to_string(),
            },
            sort_order: SortOrder::Ascending,
            manual_sort: None,
        }],
        // Use shorthand cache field name; should resolve to `base_table[Region]` case-insensitively.
        column_fields: vec![PivotField {
            source_field: PivotFieldRef::CacheFieldName("region".to_string()),
            sort_order: SortOrder::Ascending,
            manual_sort: None,
        }],
        value_fields: vec![ValueField {
            // Shorthand cache field name resolves to `base_table[Maß]`. Use ASCII-only spelling to
            // exercise ß/SS folding (`Maß` -> `MASS`).
            source_field: PivotFieldRef::CacheFieldName("MASS".to_string()),
            name: "Total".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        ..Default::default()
    };

    let result = pivot_crosstab_from_config(&model, "strasse", &cfg).unwrap();
    assert_eq!(
        result.data,
        vec![
            vec![
                Value::from("Straße[StraßenId]"),
                Value::from("East"),
                Value::from("West"),
            ],
            vec![1.into(), 30.0.into(), Value::Blank],
            vec![2.into(), Value::Blank, 5.0.into()],
        ]
    );
}

#[test]
fn pivot_crosstab_from_config_applies_unicode_filter_fields_case_insensitively() {
    let mut model = DataModel::new();
    let mut fact = Table::new("Straße", vec!["StraßenId", "Region", "Maß"]);
    fact.push_row(vec![1.into(), "East".into(), 10.0.into()])
        .unwrap();
    fact.push_row(vec![1.into(), "East".into(), 20.0.into()])
        .unwrap();
    fact.push_row(vec![2.into(), "West".into(), 5.0.into()])
        .unwrap();
    model.add_table(fact).unwrap();

    let allowed = HashSet::from([PivotKeyPart::Text("East".to_string())]);
    let cfg = PivotConfig {
        row_fields: vec![PivotField {
            source_field: PivotFieldRef::CacheFieldName("StraßenId".to_string()),
            sort_order: SortOrder::Ascending,
            manual_sort: None,
        }],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: PivotFieldRef::CacheFieldName("MASS".to_string()),
            name: "Total".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![FilterField {
            source_field: PivotFieldRef::DataModelColumn {
                table: "STRASSE".to_string(),
                column: "REGION".to_string(),
            },
            allowed: Some(allowed),
        }],
        ..Default::default()
    };

    let result = pivot_crosstab_from_config(&model, "STRASSE", &cfg).unwrap();
    // Only the East rows should remain (StraßenId=1).
    assert_eq!(
        result.data,
        vec![
            vec![Value::from("Straße[StraßenId]"), Value::from("Total")],
            vec![1.into(), 30.0.into()],
        ]
    );
}

