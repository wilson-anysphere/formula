use formula_model::pivots::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotDestination, PivotField, PivotSource,
    PivotTableModel, SortOrder, SubtotalPosition, ValueField,
};
use formula_model::{CellRef, Range, Workbook};

use uuid::Uuid;

#[test]
fn workbook_json_roundtrips_with_pivot_tables() {
    let mut wb = Workbook::new();
    let sheet_id = wb.add_sheet("Data").unwrap();

    let pivot_id = Uuid::from_u128(1);
    let pivot = PivotTableModel {
        id: pivot_id,
        name: "PivotTable1".to_string(),
        source: PivotSource::Range {
            sheet_id,
            range: Range::from_a1("A1:C10").unwrap(),
        },
        destination: PivotDestination::Cell {
            sheet_id,
            cell: CellRef::new(0, 5), // F1
        },
        config: PivotConfig {
            row_fields: vec![PivotField {
                source_field: "Region".to_string(),
                sort_order: SortOrder::default(),
                manual_sort: None,
            }],
            value_fields: vec![ValueField {
                source_field: "Sales".to_string(),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: true,
            },
            ..PivotConfig::default()
        },
        cache_id: None,
    };

    wb.pivot_tables.push(pivot.clone());

    let json = serde_json::to_value(&wb).unwrap();
    let pivot_tables = json
        .get("pivotTables")
        .expect("workbook JSON should contain pivotTables when non-empty");
    assert_eq!(pivot_tables.as_array().unwrap().len(), 1);

    let decoded: Workbook = serde_json::from_value(json).unwrap();
    assert_eq!(decoded.pivot_tables, vec![pivot]);
}

#[test]
fn missing_pivot_tables_defaults_to_empty() {
    let wb: Workbook = serde_json::from_value(serde_json::json!({})).unwrap();
    assert!(wb.pivot_tables.is_empty());
}

