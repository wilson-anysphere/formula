use pretty_assertions::assert_eq;

use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotValue, SubtotalPosition,
    ValueField,
};
use formula_engine::Engine;
use formula_model::{Font, Style};

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
            source_field: "Sales".into(),
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
    assert_eq!(
        engine.get_cell_value("Sheet1", "C3"),
        formula_engine::Value::Number(150.0)
    );

    // Ensure the pivot output values are typed as expected.
    assert_eq!(result.data[2][1], PivotValue::Number(450.0));
}

#[test]
fn engine_pivot_infers_dates_from_cell_number_formats() {
    use chrono::NaiveDate;
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
            source_field: "Amount".into(),
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

    // Row labels should be typed dates, not raw serial numbers or pre-formatted strings.
    assert_eq!(
        result.data,
        vec![
            vec!["Date".into(), "Sum of Amount".into()],
            vec![
                NaiveDate::from_ymd_opt(2024, 1, 15).unwrap().into(),
                10.into()
            ],
            vec![
                NaiveDate::from_ymd_opt(2024, 2, 1).unwrap().into(),
                20.into()
            ],
            vec!["Grand Total".into(), 30.into()],
        ]
    );
}

#[test]
fn engine_pivot_infers_dates_from_column_number_formats_when_cell_styles_inherit_num_fmt() {
    use chrono::NaiveDate;
    use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};

    let mut engine = Engine::new();

    // Source data: a date serial column.
    engine.set_cell_value("Sheet1", "A1", "Date").unwrap();

    // Apply the date number format via the column default style.
    let date_style_id = engine.intern_style(Style {
        number_format: Some("m/d/yyyy".to_string()),
        ..Style::default()
    });
    engine.set_col_style_id("Sheet1", 0, Some(date_style_id));

    // Simulate an additional cell-level style (e.g. bold) that does *not* set a number format.
    // Pivot typing should still treat the serial as a date by inheriting the column's number format.
    let bold_style_id = engine.intern_style(Style {
        font: Some(Font {
            bold: true,
            ..Font::default()
        }),
        ..Style::default()
    });

    let date1_serial =
        ymd_to_serial(ExcelDate::new(2024, 1, 15), ExcelDateSystem::EXCEL_1900).unwrap() as f64;
    let date2_serial =
        ymd_to_serial(ExcelDate::new(2024, 2, 1), ExcelDateSystem::EXCEL_1900).unwrap() as f64;

    engine.set_cell_value("Sheet1", "A2", date1_serial).unwrap();
    engine
        .set_cell_style_id("Sheet1", "A2", bold_style_id)
        .unwrap();

    engine.set_cell_value("Sheet1", "A3", date2_serial).unwrap();
    engine
        .set_cell_style_id("Sheet1", "A3", bold_style_id)
        .unwrap();

    engine.recalculate();

    let range = formula_model::Range::from_a1("A1:A3").unwrap();
    let cache = engine.pivot_cache_from_range("Sheet1", range).unwrap();
    assert_eq!(cache.records.len(), 2);
    assert_eq!(
        cache.records[0][0],
        PivotValue::Date(NaiveDate::from_ymd_opt(2024, 1, 15).unwrap())
    );
    assert_eq!(
        cache.records[1][0],
        PivotValue::Date(NaiveDate::from_ymd_opt(2024, 2, 1).unwrap())
    );
}

#[test]
fn engine_pivot_infers_dates_from_range_run_number_formats_when_cell_styles_inherit_num_fmt() {
    use chrono::NaiveDate;
    use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
    use formula_engine::metadata::FormatRun;

    let mut engine = Engine::new();

    // Source data: a date serial column.
    engine.set_cell_value("Sheet1", "A1", "Date").unwrap();

    // Apply the date number format via range-run formatting for the data rows.
    let date_style_id = engine.intern_style(Style {
        number_format: Some("m/d/yyyy".to_string()),
        ..Style::default()
    });
    engine
        .set_format_runs_by_col(
            "Sheet1",
            0, // col A
            vec![FormatRun {
                start_row: 1,         // row 2
                end_row_exclusive: 3, // row 4 exclusive (covers rows 2-3)
                style_id: date_style_id,
            }],
        )
        .unwrap();

    // Simulate an additional cell-level style (e.g. bold) that does *not* set a number format.
    // Pivot typing should still treat the serial as a date by inheriting the range-run's number
    // format.
    let bold_style_id = engine.intern_style(Style {
        font: Some(Font {
            bold: true,
            ..Font::default()
        }),
        ..Style::default()
    });

    let date1_serial =
        ymd_to_serial(ExcelDate::new(2024, 1, 15), ExcelDateSystem::EXCEL_1900).unwrap() as f64;
    let date2_serial =
        ymd_to_serial(ExcelDate::new(2024, 2, 1), ExcelDateSystem::EXCEL_1900).unwrap() as f64;

    engine.set_cell_value("Sheet1", "A2", date1_serial).unwrap();
    engine
        .set_cell_style_id("Sheet1", "A2", bold_style_id)
        .unwrap();

    engine.set_cell_value("Sheet1", "A3", date2_serial).unwrap();
    engine
        .set_cell_style_id("Sheet1", "A3", bold_style_id)
        .unwrap();

    engine.recalculate();

    let range = formula_model::Range::from_a1("A1:A3").unwrap();
    let cache = engine.pivot_cache_from_range("Sheet1", range).unwrap();
    assert_eq!(cache.records.len(), 2);
    assert_eq!(
        cache.records[0][0],
        PivotValue::Date(NaiveDate::from_ymd_opt(2024, 1, 15).unwrap())
    );
    assert_eq!(
        cache.records[1][0],
        PivotValue::Date(NaiveDate::from_ymd_opt(2024, 2, 1).unwrap())
    );
}

#[test]
fn engine_pivot_prefers_row_number_format_over_column_date_format() {
    use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};

    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", "Date").unwrap();

    let serial =
        ymd_to_serial(ExcelDate::new(2024, 1, 15), ExcelDateSystem::EXCEL_1900).unwrap() as f64;
    engine.set_cell_value("Sheet1", "A2", serial).unwrap();

    // Set a date number format on the column.
    let col_date_style = engine.intern_style(Style {
        number_format: Some("m/d/yyyy".to_string()),
        ..Style::default()
    });
    engine.set_col_style_id("Sheet1", 0, Some(col_date_style));

    // Override the row with a non-date numeric format. Row formatting should win over column
    // formatting (`sheet < col < row < cell`), so the serial should be treated as a number.
    let row_number_style = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });
    engine.set_row_style_id("Sheet1", 1, Some(row_number_style)); // row 2

    engine.recalculate();

    let range = formula_model::Range::from_a1("A1:A2").unwrap();
    let cache = engine.pivot_cache_from_range("Sheet1", range).unwrap();
    assert_eq!(cache.records.len(), 1);
    assert_eq!(cache.records[0][0], PivotValue::Number(serial));
}

#[test]
fn engine_pivot_coerces_non_finite_numbers_to_num_error_text() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", "Category").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Amount").unwrap();

    engine.set_cell_value("Sheet1", "A2", "A").unwrap();
    engine.set_cell_value("Sheet1", "B2", f64::NAN).unwrap();

    engine.recalculate();

    let range = formula_model::Range::from_a1("A1:B2").unwrap();
    let cache = engine.pivot_cache_from_range("Sheet1", range).unwrap();

    assert_eq!(cache.records.len(), 1);
    assert_eq!(cache.records[0][0], PivotValue::Text("A".to_string()));
    assert_eq!(cache.records[0][1], PivotValue::Text("#NUM!".to_string()));

    // Ensure schema sampling remains JSON-friendly (no NaNs).
    let schema = cache.schema(10);
    assert_eq!(schema.fields[1].sample_values, vec![PivotValue::Text("#NUM!".to_string())]);
}
