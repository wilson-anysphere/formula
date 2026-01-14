use chrono::NaiveDate;
use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotCache, PivotConfig, PivotField, PivotFieldRef,
    PivotTable, PivotValue, SubtotalPosition, ValueField,
};
use formula_xlsx::pivots::engine_bridge::{
    pivot_cache_to_engine_source, timeline_selection_to_engine_filter,
};
use formula_xlsx::{
    PivotCacheDefinition, PivotCacheField, PivotCacheRecordsReader, TimelineSelectionState,
};

use pretty_assertions::assert_eq;

#[test]
fn timeline_date_range_filters_by_date_values_in_pivot_cache() {
    let cache_def = PivotCacheDefinition {
        cache_fields: vec![
            PivotCacheField {
                name: "OrderDate".to_string(),
                ..Default::default()
            },
            PivotCacheField {
                name: "Sales".to_string(),
                ..Default::default()
            },
        ],
        ..Default::default()
    };

    let records_xml = br#"
        <pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <r><d v="2024-01-01T00:00:00Z"/><n v="100"/></r>
          <r><d v="2024-01-02T00:00:00Z"/><n v="200"/></r>
          <r><d v="2024-01-03T00:00:00Z"/><n v="300"/></r>
        </pivotCacheRecords>
    "#;

    let mut reader = PivotCacheRecordsReader::new(records_xml);
    let records = reader.parse_all_records();
    let source = pivot_cache_to_engine_source(&cache_def, records.into_iter());

    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("OrderDate")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: PivotFieldRef::CacheFieldName("Sales".to_string()),
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

    let pivot_all = PivotTable::new("PivotTable1", &source, cfg.clone()).expect("pivot");
    let result_all = pivot_all.calculate().expect("calculate");
    let d1 = NaiveDate::from_ymd_opt(2024, 1, 1).expect("valid date");
    let d2 = NaiveDate::from_ymd_opt(2024, 1, 2).expect("valid date");
    let d3 = NaiveDate::from_ymd_opt(2024, 1, 3).expect("valid date");
    assert_eq!(
        result_all.data,
        vec![
            vec![
                PivotValue::Text("OrderDate".to_string()),
                PivotValue::Text("Sum of Sales".to_string())
            ],
            vec![PivotValue::Date(d1), PivotValue::Number(100.0)],
            vec![PivotValue::Date(d2), PivotValue::Number(200.0)],
            vec![PivotValue::Date(d3), PivotValue::Number(300.0)],
            vec![PivotValue::Text("Grand Total".to_string()), PivotValue::Number(600.0)],
        ]
    );

    let selection = TimelineSelectionState {
        start: Some("2024-01-02".to_string()),
        end: Some("2024-01-02".to_string()),
    };

    let cache = PivotCache::from_range(&source).expect("cache");
    let timeline_filter =
        timeline_selection_to_engine_filter("OrderDate", &selection, &cache).expect("filter");

    let mut filtered_cfg = cfg;
    filtered_cfg.filter_fields = vec![timeline_filter];

    let pivot_filtered = PivotTable::new("PivotTable1", &source, filtered_cfg).expect("pivot");
    let result_filtered = pivot_filtered.calculate().expect("calculate");

    assert_eq!(
        result_filtered.data,
        vec![
            vec![
                PivotValue::Text("OrderDate".to_string()),
                PivotValue::Text("Sum of Sales".to_string())
            ],
            vec![PivotValue::Date(d2), PivotValue::Number(200.0)],
            vec![PivotValue::Text("Grand Total".to_string()), PivotValue::Number(200.0)],
        ]
    );
}
