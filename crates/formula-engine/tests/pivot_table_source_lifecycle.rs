use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotDestination, PivotField, PivotFieldRef,
    PivotSource, PivotTableDefinition, SubtotalPosition, ValueField,
};
use formula_engine::Engine;
use formula_model::{CellRef, Range, Table, TableColumn};

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

#[test]
fn deleting_sheet_prunes_table_sourced_pivots() {
    let mut engine = Engine::new();

    seed_sales_data(&mut engine, "Source");
    engine.ensure_sheet("Dest");

    engine.set_sheet_tables(
        "Source",
        vec![Table {
            id: 1,
            name: "Table1".to_string(),
            display_name: "Table1".to_string(),
            range: range("A1:B5"),
            header_row_count: 1,
            totals_row_count: 0,
            columns: vec![
                TableColumn {
                    id: 1,
                    name: "Region".to_string(),
                    formula: None,
                    totals_formula: None,
                },
                TableColumn {
                    id: 2,
                    name: "Sales".to_string(),
                    formula: None,
                    totals_formula: None,
                },
            ],
            style: None,
            auto_filter: None,
            relationship_id: None,
            part_path: None,
        }],
    );

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Table { table_id: 1 },
        destination: PivotDestination {
            sheet: "Dest".to_string(),
            cell: cell("A1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.delete_sheet("Source").unwrap();
    assert!(engine.pivot_table(pivot_id).is_none());
    assert!(engine.sheet_id("Dest").is_some());
}

