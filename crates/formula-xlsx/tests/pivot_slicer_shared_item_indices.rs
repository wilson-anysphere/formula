use std::collections::HashSet;

use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotTable, PivotValue,
    SubtotalPosition, ValueField,
};
use formula_xlsx::pivots::engine_bridge::{
    pivot_cache_to_engine_source, slicer_selection_to_engine_filter_with_resolver,
};
use formula_xlsx::{
    PivotCacheDefinition, PivotCacheField, PivotCacheRecordsReader, PivotCacheValue,
    SlicerSelectionState,
};

use pretty_assertions::assert_eq;

#[test]
fn slicer_x_indices_filter_pivot_results_via_shared_items() {
    // Minimal cache definition: a text field whose records are stored as shared-item indices.
    let cache_def = PivotCacheDefinition {
        cache_fields: vec![
            PivotCacheField {
                name: "Region".to_string(),
                shared_items: Some(vec![
                    PivotCacheValue::String("East".to_string()),
                    PivotCacheValue::String("West".to_string()),
                ]),
                ..Default::default()
            },
            PivotCacheField {
                name: "Sales".to_string(),
                ..Default::default()
            },
        ],
        ..Default::default()
    };

    // Cache records use `<x v="..."/>` indices for the Region field.
    let records_xml = br#"
        <pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <r><x v="0"/><n v="100"/></r>
          <r><x v="0"/><n v="150"/></r>
          <r><x v="1"/><n v="200"/></r>
        </pivotCacheRecords>
    "#;

    let mut reader = PivotCacheRecordsReader::new(records_xml);
    let records = reader.parse_all_records();
    let source = pivot_cache_to_engine_source(&cache_def, records.into_iter());

    let base_cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: "Sales".to_string(),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: None,
            base_field: None,
            base_item: None,
        }],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: true,
            columns: false,
        },
    };

    let pivot_all = PivotTable::new("PivotTable1", &source, base_cfg.clone()).expect("pivot");
    let result_all = pivot_all.calculate().expect("calculate");
    assert_eq!(
        result_all.data,
        vec![
            vec![PivotValue::Text("Region".to_string()), PivotValue::Text("Sum of Sales".to_string())],
            vec![PivotValue::Text("East".to_string()), PivotValue::Number(250.0)],
            vec![PivotValue::Text("West".to_string()), PivotValue::Number(200.0)],
            vec![PivotValue::Text("Grand Total".to_string()), PivotValue::Number(450.0)],
        ]
    );

    // Slicer selection is stored as the shared-item index "0" -> "East".
    let selection = SlicerSelectionState {
        available_items: vec!["0".to_string(), "1".to_string()],
        selected_items: Some(HashSet::from(["0".to_string()])),
    };

    let slicer_filter = slicer_selection_to_engine_filter_with_resolver(
        "Region",
        &selection,
        |key| key.parse::<u32>().ok().and_then(|idx| cache_def.resolve_shared_item(0, idx)),
    );

    let mut filtered_cfg = base_cfg;
    filtered_cfg.filter_fields = vec![slicer_filter];

    let pivot_filtered = PivotTable::new("PivotTable1", &source, filtered_cfg).expect("pivot");
    let result_filtered = pivot_filtered.calculate().expect("calculate");

    assert_eq!(
        result_filtered.data,
        vec![
            vec![PivotValue::Text("Region".to_string()), PivotValue::Text("Sum of Sales".to_string())],
            vec![PivotValue::Text("East".to_string()), PivotValue::Number(250.0)],
            vec![PivotValue::Text("Grand Total".to_string()), PivotValue::Number(250.0)],
        ]
    );
}
