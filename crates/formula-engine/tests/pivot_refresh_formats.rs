use chrono::{Datelike, NaiveDate};
use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotDestination, PivotField, PivotFieldRef,
    PivotSource, PivotTableDefinition, ShowAsType, SubtotalPosition, ValueField,
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

fn date_style_id(engine: &mut Engine, fmt: &str) -> u32 {
    engine.intern_style(Style {
        number_format: Some(fmt.to_string()),
        ..Style::default()
    })
}

#[test]
fn refresh_pivot_writes_dates_as_serial_numbers_and_sets_date_format() {
    let date = NaiveDate::from_ymd_opt(1904, 1, 1).unwrap();

    for system in [
        ExcelDateSystem::Excel1900 { lotus_compat: true },
        ExcelDateSystem::Excel1904,
    ] {
        let mut engine = Engine::new();
        engine.set_date_system(system);

        // Seed source data. Dates are stored as numbers + date number format (Excel semantics).
        engine.set_cell_value("Sheet1", "A1", "Date").unwrap();
        engine.set_cell_value("Sheet1", "B1", "Sales").unwrap();

        let serial = ymd_to_serial(
            ExcelDate::new(date.year(), date.month() as u8, date.day() as u8),
            system,
        )
        .unwrap() as f64;
        engine.set_cell_value("Sheet1", "A2", serial).unwrap();
        engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();

        let date_style = date_style_id(&mut engine, "m/d/yyyy");
        engine
            .set_cell_style_id("Sheet1", "A2", date_style)
            .unwrap();

        let pivot_id = engine.add_pivot_table(PivotTableDefinition {
            id: 0,
            name: "Sales by Date".to_string(),
            source: PivotSource::Range {
                sheet: "Sheet1".to_string(),
                range: Some(range("A1:B2")),
            },
            destination: PivotDestination {
                sheet: "Sheet1".to_string(),
                cell: cell("D1"),
            },
            config: PivotConfig {
                row_fields: vec![PivotField::new("Date")],
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
                    rows: false,
                    columns: false,
                },
            },
            apply_number_formats: true,
            last_output_range: None,
            needs_refresh: true,
        });

        engine.refresh_pivot_table(pivot_id).unwrap();

        // The date label cell is written as a number (serial) + date format.
        assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(serial));
        let style_id = engine
            .get_cell_style_id("Sheet1", "D2")
            .unwrap()
            .unwrap_or(0);
        assert_ne!(style_id, 0);
        let style = engine.style_table().get(style_id).unwrap();
        assert_eq!(style.number_format.as_deref(), Some("m/d/yyyy"));
    }
}

#[test]
fn refresh_pivot_writes_dates_as_serial_numbers_and_sets_date_format_in_compact_layout() {
    let date = NaiveDate::from_ymd_opt(1904, 1, 1).unwrap();

    for system in [
        ExcelDateSystem::Excel1900 { lotus_compat: true },
        ExcelDateSystem::Excel1904,
    ] {
        let mut engine = Engine::new();
        engine.set_date_system(system);

        // Seed source data. Dates are stored as numbers + date number format (Excel semantics).
        engine.set_cell_value("Sheet1", "A1", "Date").unwrap();
        engine.set_cell_value("Sheet1", "B1", "Sales").unwrap();

        let serial = ymd_to_serial(
            ExcelDate::new(date.year(), date.month() as u8, date.day() as u8),
            system,
        )
        .unwrap() as f64;
        engine.set_cell_value("Sheet1", "A2", serial).unwrap();
        engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();

        let date_style = date_style_id(&mut engine, "m/d/yyyy");
        engine
            .set_cell_style_id("Sheet1", "A2", date_style)
            .unwrap();

        let pivot_id = engine.add_pivot_table(PivotTableDefinition {
            id: 0,
            name: "Sales by Date".to_string(),
            source: PivotSource::Range {
                sheet: "Sheet1".to_string(),
                range: Some(range("A1:B2")),
            },
            destination: PivotDestination {
                sheet: "Sheet1".to_string(),
                cell: cell("D1"),
            },
            config: PivotConfig {
                row_fields: vec![PivotField::new("Date")],
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
                layout: Layout::Compact,
                subtotals: SubtotalPosition::None,
                grand_totals: GrandTotals {
                    rows: false,
                    columns: false,
                },
            },
            apply_number_formats: true,
            last_output_range: None,
            needs_refresh: true,
        });

        engine.refresh_pivot_table(pivot_id).unwrap();

        // The date label cell is written as a number (serial) + date format.
        assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(serial));
        let style_id = engine
            .get_cell_style_id("Sheet1", "D2")
            .unwrap()
            .unwrap_or(0);
        assert_ne!(style_id, 0);
        let style = engine.style_table().get(style_id).unwrap();
        assert_eq!(style.number_format.as_deref(), Some("m/d/yyyy"));
    }
}

#[test]
fn refresh_pivot_infers_dates_from_column_number_formats_when_cell_styles_inherit_num_fmt() {
    let mut engine = Engine::new();

    // Seed source data. Dates are stored as numbers + date number format (Excel semantics).
    engine.set_cell_value("Sheet1", "A1", "Date").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Sales").unwrap();

    let serial = ymd_to_serial(
        ExcelDate::new(2024, 1, 15),
        ExcelDateSystem::Excel1900 { lotus_compat: true },
    )
    .unwrap() as f64;
    engine.set_cell_value("Sheet1", "A2", serial).unwrap();
    engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();

    // Apply the date number format via the column default style.
    let date_style = date_style_id(&mut engine, "m/d/yyyy");
    engine.set_col_style_id("Sheet1", 0, Some(date_style));

    // Simulate a separate cell-level style that does not specify a number format.
    let bold_style = engine.intern_style(Style {
        font: Some(Font {
            bold: true,
            ..Font::default()
        }),
        ..Style::default()
    });
    engine.set_cell_style_id("Sheet1", "A2", bold_style).unwrap();

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Date".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B2")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: PivotConfig {
            row_fields: vec![PivotField::new("Date")],
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
                rows: false,
                columns: false,
            },
        },
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();

    // The date label cell is written as a number (serial) + date format.
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(serial));
    let style_id = engine
        .get_cell_style_id("Sheet1", "D2")
        .unwrap()
        .unwrap_or(0);
    assert_ne!(style_id, 0);
    let style = engine.style_table().get(style_id).unwrap();
    assert_eq!(style.number_format.as_deref(), Some("m/d/yyyy"));
}

#[test]
fn refresh_pivot_infers_dates_from_range_run_number_formats_when_cell_styles_inherit_num_fmt() {
    use formula_engine::metadata::FormatRun;

    let mut engine = Engine::new();

    // Seed source data. Dates are stored as numbers + date number format (Excel semantics).
    engine.set_cell_value("Sheet1", "A1", "Date").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Sales").unwrap();

    let serial = ymd_to_serial(
        ExcelDate::new(2024, 1, 15),
        ExcelDateSystem::Excel1900 { lotus_compat: true },
    )
    .unwrap() as f64;
    engine.set_cell_value("Sheet1", "A2", serial).unwrap();
    engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();

    // Apply the date number format via range-run formatting for the data row.
    let date_style = date_style_id(&mut engine, "m/d/yyyy");
    engine
        .set_format_runs_by_col(
            "Sheet1",
            0, // col A
            vec![FormatRun {
                start_row: 1,         // row 2
                end_row_exclusive: 2, // row 3 exclusive (covers row 2)
                style_id: date_style,
            }],
        )
        .unwrap();

    // Simulate a separate cell-level style that does not specify a number format.
    let bold_style = engine.intern_style(Style {
        font: Some(Font {
            bold: true,
            ..Font::default()
        }),
        ..Style::default()
    });
    engine.set_cell_style_id("Sheet1", "A2", bold_style).unwrap();

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Date".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B2")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: PivotConfig {
            row_fields: vec![PivotField::new("Date")],
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
                rows: false,
                columns: false,
            },
        },
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();

    // The date label cell is written as a number (serial) + date format.
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(serial));
    let style_id = engine
        .get_cell_style_id("Sheet1", "D2")
        .unwrap()
        .unwrap_or(0);
    assert_ne!(style_id, 0);
    let style = engine.style_table().get(style_id).unwrap();
    assert_eq!(style.number_format.as_deref(), Some("m/d/yyyy"));
}

#[test]
fn refresh_pivot_prefers_row_number_format_over_column_date_format() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", "Date").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Sales").unwrap();

    let serial = ymd_to_serial(
        ExcelDate::new(2024, 1, 15),
        ExcelDateSystem::Excel1900 { lotus_compat: true },
    )
    .unwrap() as f64;
    engine.set_cell_value("Sheet1", "A2", serial).unwrap();
    engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();

    // Column says "date"...
    let date_style = date_style_id(&mut engine, "m/d/yyyy");
    engine.set_col_style_id("Sheet1", 0, Some(date_style));

    // ...but the row overrides it with a non-date number format. Row should win over column, so
    // the pivot should treat the serial as a number (no date output formatting).
    let number_style = engine.intern_style(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });
    engine.set_row_style_id("Sheet1", 1, Some(number_style)); // row 2

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Sales by Date".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B2")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: PivotConfig {
            row_fields: vec![PivotField::new("Date")],
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
                rows: false,
                columns: false,
            },
        },
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();

    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(serial));
    let style_id = engine
        .get_cell_style_id("Sheet1", "D2")
        .unwrap()
        .unwrap_or(0);
    assert_eq!(style_id, 0);
}

#[test]
fn refresh_pivot_applies_value_field_number_format() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "Region").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Sales").unwrap();
    engine.set_cell_value("Sheet1", "A2", "East").unwrap();
    engine.set_cell_value("Sheet1", "B2", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", "East").unwrap();
    engine.set_cell_value("Sheet1", "B3", 150.0).unwrap();

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
        config: PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: PivotFieldRef::CacheFieldName("Sales".to_string()),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: Some("$#,##0.00".to_string()),
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
                rows: false,
                columns: false,
            },
        },
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();

    let style_id = engine
        .get_cell_style_id("Sheet1", "E2")
        .unwrap()
        .unwrap_or(0);
    assert_ne!(style_id, 0);
    let style = engine.style_table().get(style_id).unwrap();
    assert_eq!(style.number_format.as_deref(), Some("$#,##0.00"));
}

#[test]
fn refresh_pivot_applies_percent_format_for_percent_show_as_when_no_explicit_format() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "Region").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Sales").unwrap();
    engine.set_cell_value("Sheet1", "A2", "East").unwrap();
    engine.set_cell_value("Sheet1", "B2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", "West").unwrap();
    engine.set_cell_value("Sheet1", "B3", 3.0).unwrap();

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Pct of Total".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B3")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: PivotFieldRef::CacheFieldName("Sales".to_string()),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::PercentOfGrandTotal),
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        },
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();

    // Value cells should have a percent number format applied when no explicit format is set.
    let style_id = engine
        .get_cell_style_id("Sheet1", "E2")
        .unwrap()
        .unwrap_or(0);
    assert_ne!(style_id, 0);
    let style = engine.style_table().get(style_id).unwrap();
    assert_eq!(style.number_format.as_deref(), Some("0.00%"));
}

#[test]
fn refresh_pivot_does_not_apply_value_field_number_format_when_apply_number_formats_false() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "Region").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Sales").unwrap();
    engine.set_cell_value("Sheet1", "A2", "East").unwrap();
    engine.set_cell_value("Sheet1", "B2", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", "East").unwrap();
    engine.set_cell_value("Sheet1", "B3", 150.0).unwrap();

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
        config: PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: PivotFieldRef::CacheFieldName("Sales".to_string()),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: Some("$#,##0.00".to_string()),
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
                rows: false,
                columns: false,
            },
        },
        apply_number_formats: false,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();

    let style_id = engine
        .get_cell_style_id("Sheet1", "E2")
        .unwrap()
        .unwrap_or(0);
    assert_eq!(style_id, 0);
}

#[test]
fn refresh_pivot_clears_value_field_number_format_when_apply_number_formats_toggled_off() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "Region").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Sales").unwrap();
    engine.set_cell_value("Sheet1", "A2", "East").unwrap();
    engine.set_cell_value("Sheet1", "B2", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", "East").unwrap();
    engine.set_cell_value("Sheet1", "B3", 150.0).unwrap();

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
        config: PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: PivotFieldRef::CacheFieldName("Sales".to_string()),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: Some("$#,##0.00".to_string()),
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
                rows: false,
                columns: false,
            },
        },
        apply_number_formats: true,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();
    let style_id = engine
        .get_cell_style_id("Sheet1", "E2")
        .unwrap()
        .unwrap_or(0);
    assert_ne!(style_id, 0);
    let style = engine.style_table().get(style_id).unwrap();
    assert_eq!(style.number_format.as_deref(), Some("$#,##0.00"));

    // Toggle off number-format application; refresh should clear the old value-field format.
    {
        let pivot = engine.pivot_table_mut(pivot_id).unwrap();
        pivot.apply_number_formats = false;
        pivot.needs_refresh = true;
    }
    engine.refresh_pivot_table(pivot_id).unwrap();

    let style_id = engine
        .get_cell_style_id("Sheet1", "E2")
        .unwrap()
        .unwrap_or(0);
    assert_eq!(style_id, 0);
}

#[test]
fn refresh_pivot_does_not_apply_percent_format_when_apply_number_formats_false() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "Region").unwrap();
    engine.set_cell_value("Sheet1", "B1", "Sales").unwrap();
    engine.set_cell_value("Sheet1", "A2", "East").unwrap();
    engine.set_cell_value("Sheet1", "B2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", "West").unwrap();
    engine.set_cell_value("Sheet1", "B3", 3.0).unwrap();

    let pivot_id = engine.add_pivot_table(PivotTableDefinition {
        id: 0,
        name: "Pct of Total".to_string(),
        source: PivotSource::Range {
            sheet: "Sheet1".to_string(),
            range: Some(range("A1:B3")),
        },
        destination: PivotDestination {
            sheet: "Sheet1".to_string(),
            cell: cell("D1"),
        },
        config: PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: vec![],
            value_fields: vec![ValueField {
                source_field: PivotFieldRef::CacheFieldName("Sales".to_string()),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: Some(ShowAsType::PercentOfGrandTotal),
                base_field: None,
                base_item: None,
            }],
            filter_fields: vec![],
            calculated_fields: vec![],
            calculated_items: vec![],
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: false,
                columns: false,
            },
        },
        apply_number_formats: false,
        last_output_range: None,
        needs_refresh: true,
    });

    engine.refresh_pivot_table(pivot_id).unwrap();

    let style_id = engine
        .get_cell_style_id("Sheet1", "E2")
        .unwrap()
        .unwrap_or(0);
    assert_eq!(style_id, 0);
}
