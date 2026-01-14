use chrono::{Datelike, NaiveDate};
use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::pivot::{
    AggregationType, GrandTotals, Layout, PivotConfig, PivotDestination, PivotField, PivotFieldRef,
    PivotSource, PivotTableDefinition, ShowAsType, SubtotalPosition, ValueField,
};
use formula_engine::{Engine, Value};
use formula_model::{CellRef, Range, Style};
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
