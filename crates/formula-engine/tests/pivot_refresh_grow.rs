use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotDestination, PivotField, PivotFieldRef,
    PivotSource, PivotTableDefinition, SubtotalPosition, ValueField,
};
use formula_engine::{Engine, Value};
use formula_model::{CellRef, Font, Range, Style};
use pretty_assertions::assert_eq;

fn cell(a1: &str) -> CellRef {
    CellRef::from_a1(a1).unwrap()
}

fn range(a1: &str) -> Range {
    Range::from_a1(a1).unwrap()
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

fn sum_sales_by_region_and_product_config() -> PivotConfig {
    PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![PivotField::new("Product")],
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
fn pivot_growing_clears_formatting_in_new_output_area() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", "Region").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Sales").unwrap();
    engine.set_cell_value("Sheet1", "A2", "East").unwrap();
    engine.set_cell_value("Sheet1", "B2", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", "East").unwrap();
    engine.set_cell_value("Sheet1", "B3", 200.0).unwrap();

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

    // Initial output has a single row item (East) + grand total => 3 rows (D1:E3).
    engine.refresh_pivot_table(pivot_id).unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", "D3"),
        Value::Text("Grand Total".to_string())
    );

    // Pre-seed a non-default style outside the current output footprint but inside the next one.
    let bold_style_id = engine.intern_style(Style {
        font: Some(Font {
            bold: true,
            ..Font::default()
        }),
        ..Style::default()
    });
    engine
        .set_cell_style_id("Sheet1", "D4", bold_style_id)
        .unwrap();
    engine
        .set_cell_style_id("Sheet1", "E4", bold_style_id)
        .unwrap();
    engine.set_cell_value("Sheet1", "D4", "Note").unwrap();

    // Split the second record into a new region to grow the output by one row.
    engine.set_cell_value("Sheet1", "A3", "West").unwrap();
    engine.refresh_pivot_table(pivot_id).unwrap();

    // The newly covered row should be cleared before applying pivot writes so it does not retain
    // the pre-existing style.
    assert_eq!(
        engine.get_cell_value("Sheet1", "D4"),
        Value::Text("Grand Total".to_string())
    );
    assert_eq!(engine.get_cell_style_id("Sheet1", "D4").unwrap(), Some(0));
    assert_eq!(engine.get_cell_style_id("Sheet1", "E4").unwrap(), Some(0));
}

#[test]
fn pivot_refresh_does_not_clear_cells_outside_old_or_new_output_when_shape_changes() {
    let mut engine = Engine::new();

    // Initial state: 1 region (rows) × 3 products (columns) -> wide-but-short pivot output.
    engine.set_cell_value("Sheet1", "A1", "Region").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Product").unwrap();
    engine.set_cell_value("Sheet1", "C1", "Sales").unwrap();
    engine.set_cell_value("Sheet1", "A2", "East").unwrap();
    engine.set_cell_value("Sheet1", "B2", "A").unwrap();
    engine.set_cell_value("Sheet1", "C2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", "East").unwrap();
    engine.set_cell_value("Sheet1", "B3", "B").unwrap();
    engine.set_cell_value("Sheet1", "C3", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", "East").unwrap();
    engine.set_cell_value("Sheet1", "B4", "C").unwrap();
    engine.set_cell_value("Sheet1", "C4", 30.0).unwrap();

    let source_range = range("A1:C4");
    let destination = cell("D1");
    let config = sum_sales_by_region_and_product_config();

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Region and Product".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(source_range),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: destination,
        },
        config: config.clone(),
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    let prev_output_range = engine
        .refresh_pivot_table(pivot_id)
        .unwrap()
        .output_range
        .unwrap();

    // Mutate the source so the pivot becomes narrow-but-tall:
    // - 3 regions (rows)
    // - 1 product (column)
    engine.set_cell_value("Sheet1", "A3", "West").unwrap();
    engine.set_cell_value("Sheet1", "B3", "A").unwrap();
    engine.set_cell_value("Sheet1", "A4", "North").unwrap();
    engine.set_cell_value("Sheet1", "B4", "A").unwrap();

    // Compute the expected new output footprint without applying it to the sheet. We use this to
    // choose a cell that lies inside the *bounding box* of the old/new rectangles but outside both
    // rectangles (the "bottom-right corner" of the L-shaped union).
    let new_result = engine
        .calculate_pivot_from_range("Sheet1", source_range, &config)
        .unwrap();
    let rows = new_result.data.len() as u32;
    let cols = new_result.data.first().map(|r| r.len()).unwrap_or(0) as u32;
    assert!(rows > 0 && cols > 0);
    let new_output_range = Range::new(
        destination,
        CellRef::new(destination.row + rows - 1, destination.col + cols - 1),
    );

    let bbox = prev_output_range.bounding_box(&new_output_range);
    let note_cell = bbox.end.to_a1();

    // Sanity: `note_cell` should be outside both output ranges, otherwise this test isn't
    // exercising the shape-change case.
    assert!(!prev_output_range.contains(bbox.end));
    assert!(!new_output_range.contains(bbox.end));

    engine.set_cell_value("Sheet1", &note_cell, "Keep").unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", &note_cell),
        Value::Text("Keep".to_string())
    );

    // Refresh should only clear cells that were previously output by the pivot or will be output
    // by the pivot now. Cells outside both footprints must remain untouched.
    engine.refresh_pivot_table(pivot_id).unwrap();
    assert_eq!(
        engine.get_cell_value("Sheet1", &note_cell),
        Value::Text("Keep".to_string())
    );
}
