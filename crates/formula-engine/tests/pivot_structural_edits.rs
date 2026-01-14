use formula_engine::{EditOp, Engine, Value};
use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotDestination, PivotField, PivotFieldRef,
    PivotSource, PivotTableDefinition, SubtotalPosition, ValueField,
};
use formula_model::{CellRef, Range};
use pretty_assertions::assert_eq;

fn cell(a1: &str) -> CellRef {
    CellRef::from_a1(a1).unwrap()
}

fn range(a1: &str) -> Range {
    Range::from_a1(a1).unwrap()
}

fn seed_sales_data(engine: &mut Engine) {
    engine.set_cell_value("Sheet1", "A1", "Region").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Sales").unwrap();
    engine.set_cell_value("Sheet1", "A2", "East").unwrap();
    engine.set_cell_value("Sheet1", "B2", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", "East").unwrap();
    engine.set_cell_value("Sheet1", "B3", 150.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", "West").unwrap();
    engine.set_cell_value("Sheet1", "B4", 200.0).unwrap();
    engine.set_cell_value("Sheet1", "A5", "West").unwrap();
    engine.set_cell_value("Sheet1", "B5", 250.0).unwrap();
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
fn insert_rows_shifts_pivot_definition_and_refresh_uses_updated_addresses() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine);

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
        config: sum_sales_by_region_config(),
        last_output_range: None,
        needs_refresh: true,
    });

    engine
        .apply_operation(EditOp::InsertRows {
            sheet: "Sheet1".to_string(),
            row: 0,
            count: 2,
        })
        .unwrap();

    // Both the source and destination should shift down by 2 rows.
    let pivot = engine.pivot_table(pivot_id).unwrap();
    assert_eq!(pivot.destination.cell, cell("D3"));
    assert_eq!(
        pivot.source,
        PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A3:B7")),
        }
    );

    engine.refresh_pivot_table(pivot_id).unwrap();

    // Output written at shifted destination (D3).
    assert_eq!(
        engine.get_cell_value("Sheet1", "D3"),
        Value::Text("Region".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "E6"), Value::Number(700.0));
    // Original destination remains empty.
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Blank);
}

#[test]
fn insert_cols_shifts_pivot_definition_and_refresh_uses_updated_addresses() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine);

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
        config: sum_sales_by_region_config(),
        last_output_range: None,
        needs_refresh: true,
    });

    engine
        .apply_operation(EditOp::InsertCols {
            sheet: "Sheet1".to_string(),
            col: 0,
            count: 1,
        })
        .unwrap();

    // Both the source and destination should shift right by 1 column.
    let pivot = engine.pivot_table(pivot_id).unwrap();
    assert_eq!(pivot.destination.cell, cell("E1"));
    assert_eq!(
        pivot.source,
        PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("B1:C5")),
        }
    );

    engine.refresh_pivot_table(pivot_id).unwrap();

    // Output written at shifted destination (E1).
    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        Value::Text("Region".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "F4"), Value::Number(700.0));
}

#[test]
fn move_range_updates_pivot_source_and_refresh_reads_from_moved_location() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine);

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
        config: sum_sales_by_region_config(),
        last_output_range: None,
        needs_refresh: true,
    });

    engine
        .apply_operation(EditOp::MoveRange {
            sheet: "Sheet1".to_string(),
            src: range("A1:B5"),
            dst_top_left: cell("A10"),
        })
        .unwrap();

    // Source range should track the move.
    let pivot = engine.pivot_table(pivot_id).unwrap();
    assert_eq!(
        pivot.source,
        PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A10:B14")),
        }
    );

    engine.refresh_pivot_table(pivot_id).unwrap();
    assert_eq!(engine.get_cell_value("Sheet1", "E4"), Value::Number(700.0));
}
