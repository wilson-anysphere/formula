use pretty_assertions::assert_eq;

use formula_engine::{Engine, Value};
use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotDestination, PivotField, PivotFieldRef,
    PivotSource, PivotTable, PivotTableDefinition, PivotValue, SubtotalPosition, ValueField,
};
use formula_model::{CellRef, Range};

fn cell(a1: &str) -> CellRef {
    CellRef::from_a1(a1).unwrap()
}

fn range(a1: &str) -> Range {
    Range::from_a1(a1).unwrap()
}

fn seed_sales_data(engine: &mut Engine, sheet: &str) {
    engine.set_cell_value(sheet, "A1", "Region").unwrap();
    engine.set_cell_value(sheet, "B1", "Sales").unwrap();
    engine.set_cell_value(sheet, "A2", "East").unwrap();
    engine.set_cell_value(sheet, "B2", 100.0).unwrap();
    engine.set_cell_value(sheet, "A3", "East").unwrap();
    engine.set_cell_value(sheet, "B3", 150.0).unwrap();
    engine.set_cell_value(sheet, "A4", "West").unwrap();
    engine.set_cell_value(sheet, "B4", 200.0).unwrap();
    engine.set_cell_value(sheet, "A5", "West").unwrap();
    engine.set_cell_value(sheet, "B5", 250.0).unwrap();
}

fn sum_sales_by_region_config() -> PivotConfig {
    PivotConfig {
        row_fields: vec![PivotField::new("Region")],
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
    }
}

fn sales_source_as_pivot_values() -> Vec<Vec<PivotValue>> {
    vec![
        vec!["Region".into(), "Sales".into()],
        vec!["East".into(), 100.into()],
        vec!["East".into(), 150.into()],
        vec!["West".into(), 200.into()],
        vec!["West".into(), 250.into()],
    ]
}

#[test]
fn pivot_sheet_rename_and_delete_update_pivot_metadata_and_prevent_resurrection() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine, "sheet1_key");
    engine.set_sheet_display_name("sheet1_key", "Sheet1");

    let cfg = sum_sales_by_region_config();
    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B5")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: cfg.clone(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    assert!(engine.rename_sheet("Sheet1", "Renamed"));
    let pivot = engine.pivot_table(pivot_id).unwrap();
    assert_eq!(pivot.destination.sheet, "Renamed");
    assert_eq!(
        pivot.source,
        PivotSource::Range {
            sheet: "Renamed".to_string(),
            range: Some(range("A1:B5")),
        }
    );

    let refresh = engine.refresh_pivot_table(pivot_id).unwrap();

    // The rename should not be able to resurrect the old sheet name via pivot refresh.
    assert_eq!(engine.sheet_id("Sheet1"), None);
    assert_eq!(
        engine.get_cell_value("Renamed", "D1"),
        Value::Text("Region".to_string())
    );
    assert_eq!(engine.get_cell_value("Renamed", "E4"), Value::Number(700.0));

    // Register metadata for `GETPIVOTDATA` and ensure it is pruned when the destination sheet is deleted.
    let pivot_table =
        PivotTable::new("Sales by Region", &sales_source_as_pivot_values(), cfg).unwrap();
    let output_range = refresh.output_range.expect("pivot should produce output");
    engine
        .register_pivot_table("Renamed", output_range, pivot_table)
        .unwrap();
    assert_eq!(engine.pivot_registry_entries().len(), 1);

    // The engine enforces Excel-like constraints: you cannot delete the last remaining sheet.
    engine.ensure_sheet("Sheet2");

    engine.delete_sheet("Renamed").unwrap();
    assert!(engine.pivot_table(pivot_id).is_none());
    assert_eq!(engine.pivot_registry_entries().len(), 0);
}

#[test]
fn pivot_sheet_display_name_change_updates_pivot_metadata_and_prevent_resurrection() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine, "sheet1_key");
    engine.set_sheet_display_name("sheet1_key", "Sheet1");

    let cfg = sum_sales_by_region_config();
    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B5")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: cfg,
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    // Metadata-only tab rename (stable sheet key unchanged) should still rewrite pivots so a refresh
    // cannot resurrect the old tab name via `ensure_sheet`.
    engine.set_sheet_display_name("sheet1_key", "Renamed");

    let pivot = engine.pivot_table(pivot_id).unwrap();
    assert_eq!(pivot.destination.sheet, "Renamed");
    assert_eq!(
        pivot.source,
        PivotSource::Range {
            sheet: "Renamed".to_string(),
            range: Some(range("A1:B5")),
        }
    );

    engine.refresh_pivot_table(pivot_id).unwrap();
    assert_eq!(engine.sheet_id("Sheet1"), None);
    assert_eq!(
        engine.get_cell_value("Renamed", "D1"),
        Value::Text("Region".to_string())
    );
}
