use pretty_assertions::assert_eq;

use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotValue, SubtotalPosition,
    ValueField,
};
use formula_model::Style;
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

#[test]
fn engine_pivot_infers_dates_from_cell_number_formats() {
    use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};

    let mut engine = Engine::new();

    // Source data: a date serial column + a numeric measure.
    engine.set_cell_value("Sheet1", "A1", "Date").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Amount").unwrap();

    let date_style_id = engine.intern_style(Style {
        number_format: Some("m/d/yyyy".to_string()),
        ..Style::default()
    });

    let date1_serial =
        ymd_to_serial(ExcelDate::new(2024, 1, 15), ExcelDateSystem::EXCEL_1900).unwrap() as f64;
    let date2_serial =
        ymd_to_serial(ExcelDate::new(2024, 2, 1), ExcelDateSystem::EXCEL_1900).unwrap() as f64;

    engine.set_cell_value("Sheet1", "A2", date1_serial).unwrap();
    engine
        .set_cell_style_id("Sheet1", "A2", date_style_id)
        .unwrap();
    engine.set_cell_value("Sheet1", "B2", 10).unwrap();

    engine.set_cell_value("Sheet1", "A3", date2_serial).unwrap();
    engine
        .set_cell_style_id("Sheet1", "A3", date_style_id)
        .unwrap();
    engine.set_cell_value("Sheet1", "B3", 20).unwrap();

    engine.recalculate();

    let range = formula_model::Range::from_a1("A1:B3").unwrap();
    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Date")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: "Amount".to_string(),
            name: "Sum of Amount".to_string(),
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

    // Row labels should be dates (ISO strings from PivotValue::Date display), not raw serial numbers.
    assert_eq!(
        result.data,
        vec![
            vec!["Date".into(), "Sum of Amount".into()],
            vec!["2024-01-15".into(), 10.into()],
            vec!["2024-02-01".into(), 20.into()],
            vec!["Grand Total".into(), 30.into()],
        ]
    );
}
