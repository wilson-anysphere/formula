use pretty_assertions::assert_eq;

use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotValue, SubtotalPosition,
    ValueField,
};
use formula_engine::Engine;

#[test]
fn engine_can_calculate_pivot_from_live_range_values() {
    let mut engine = Engine::new();

    // Header row
    engine.set_cell_value("Sheet1", "A1", "Region").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Product").unwrap();
    engine.set_cell_value("Sheet1", "C1", "Sales").unwrap();

    // Data rows (include a formula to ensure we read calculated values)
    engine.set_cell_value("Sheet1", "A2", "East").unwrap();
    engine.set_cell_value("Sheet1", "B2", "A").unwrap();
    engine.set_cell_value("Sheet1", "C2", 100).unwrap();

    engine.set_cell_value("Sheet1", "A3", "East").unwrap();
    engine.set_cell_value("Sheet1", "B3", "B").unwrap();
    engine.set_cell_formula("Sheet1", "C3", "=C2+50").unwrap();

    engine.set_cell_value("Sheet1", "A4", "West").unwrap();
    engine.set_cell_value("Sheet1", "B4", "A").unwrap();
    engine.set_cell_value("Sheet1", "C4", 200).unwrap();

    engine.set_cell_value("Sheet1", "A5", "West").unwrap();
    engine.set_cell_value("Sheet1", "B5", "B").unwrap();
    engine.set_cell_value("Sheet1", "C5", 250).unwrap();

    engine.recalculate();

    let range = formula_model::Range::from_a1("A1:C5").unwrap();
    let cfg = PivotConfig {
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

    let result = engine
        .calculate_pivot_from_range("Sheet1", range, &cfg)
        .unwrap();

    assert_eq!(
        result.data,
        vec![
            vec!["Region".into(), "Sum of Sales".into()],
            vec!["East".into(), 250.into()],
            vec!["West".into(), 450.into()],
            vec!["Grand Total".into(), 700.into()],
        ]
    );

    // Sanity: the formula cell should have contributed (C3 = 150).
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), formula_engine::Value::Number(150.0));

    // Ensure the pivot output values are typed as expected.
    assert_eq!(result.data[2][1], PivotValue::Number(450.0));
}
