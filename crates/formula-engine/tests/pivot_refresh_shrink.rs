use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotDestination, PivotField, PivotFieldRef,
    PivotSource, PivotTableDefinition, SubtotalPosition, ValueField,
};
use formula_engine::{Engine, Value};
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
    engine.set_cell_value("Sheet1", "A3", "West").unwrap();
    engine.set_cell_value("Sheet1", "B3", 200.0).unwrap();
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
fn pivot_shrinking_clears_stale_cells() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine);

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B3")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();

    assert_eq!(
        engine.get_cell_value("Sheet1", "D3"),
        Value::Text("West".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(200.0));

    // Collapse West into East -> pivot output shrinks (removes the West row).
    engine.set_cell_value("Sheet1", "A3", "East").unwrap();
    engine.refresh_pivot_table(pivot_id).unwrap();

    // The row previously containing "West" should now be "Grand Total" and the old tail row should
    // be cleared.
    assert_eq!(
        engine.get_cell_value("Sheet1", "D3"),
        Value::Text("Grand Total".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "D4"), Value::Blank);
    assert_eq!(engine.get_cell_value("Sheet1", "E4"), Value::Blank);
}

#[test]
fn pivot_refresh_after_shrink_does_not_clear_cells_outside_latest_output() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine);

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B3")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();
    engine.set_cell_value("Sheet1", "A3", "East").unwrap();
    engine.refresh_pivot_table(pivot_id).unwrap();

    // Output shrank to 3 rows (D1:E3). User edits D4, which was previously part of the pivot output.
    engine.set_cell_value("Sheet1", "D4", "Note").unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", "D4"),
        Value::Text("Note".to_string())
    );

    // Another refresh should not clear D4 because it's outside the most recently rendered output.
    engine.set_cell_value("Sheet1", "B2", 110.0).unwrap();
    engine.refresh_pivot_table(pivot_id).unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", "D4"),
        Value::Text("Note".to_string())
    );
}

#[test]
fn pivot_refresh_failure_does_not_clear_previous_output() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine);

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B3")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", "D3"),
        Value::Text("West".to_string())
    );

    // Move the pivot destination anchor beyond the engine's supported row bounds so refresh fails.
    {
        let pivot = engine.pivot_table_mut(pivot_id).unwrap();
        pivot.destination.cell.row = i32::MAX as u32;
        pivot.needs_refresh = true;
    }

    let err = engine.refresh_pivot_table(pivot_id).unwrap_err();
    assert!(matches!(
        err,
        formula_engine::pivot::PivotRefreshError::OutputOutOfBounds
    ));

    // The previous pivot output should remain intact because the refresh never applied.
    assert_eq!(
        engine.get_cell_value("Sheet1", "D3"),
        Value::Text("West".to_string())
    );
}

#[test]
fn pivot_refresh_registers_pivot_for_getpivotdata() {
    let mut engine = Engine::new();
    seed_sales_data(&mut engine);

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B3")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: sum_sales_by_region_config(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();

    // Pivot refresh should register metadata for `GETPIVOTDATA` automatically.
    engine
        .set_cell_formula(
            "Sheet1",
            "G1",
            "=GETPIVOTDATA(\"Sum of Sales\", D1, \"Region\", \"East\")",
        )
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(100.0));
}
