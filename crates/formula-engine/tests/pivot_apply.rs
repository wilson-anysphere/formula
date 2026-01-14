use chrono::NaiveDate;
use formula_engine::pivot::{
    apply_pivot_result_to_worksheet, AggregationType, GrandTotals, Layout, PivotApplyOptions,
    PivotConfig, PivotField, PivotFieldRef, PivotResult, PivotValue, SubtotalPosition, ValueField,
};
use formula_model::{CellRef, CellValue, DateSystem, Workbook};

#[test]
fn applies_date_values_as_serial_numbers_with_date_format_for_1900_and_1904_systems() {
    let date = NaiveDate::from_ymd_opt(1904, 1, 1).unwrap();

    let result = PivotResult {
        data: vec![
            vec![PivotValue::Text("Date".to_string())],
            vec![PivotValue::Date(date)],
        ],
    };
    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Date")],
        column_fields: vec![],
        value_fields: vec![],
        filter_fields: vec![],
        calculated_fields: vec![],
        calculated_items: vec![],
        layout: Layout::Tabular,
        subtotals: SubtotalPosition::None,
        grand_totals: GrandTotals {
            rows: false,
            columns: false,
        },
    };

    // 1900 system: 1904-01-01 is serial 1462 (Excel's offset between systems).
    let mut wb_1900 = Workbook::new();
    wb_1900.date_system = DateSystem::Excel1900;
    let sheet_id = wb_1900.add_sheet("Sheet1").unwrap();
    apply_pivot_result_to_worksheet(
        &mut wb_1900,
        sheet_id,
        CellRef::new(0, 0),
        &result,
        &cfg,
        PivotApplyOptions::default(),
    )
    .unwrap();

    let cell = wb_1900
        .sheet(sheet_id)
        .unwrap()
        .cell(CellRef::new(1, 0))
        .unwrap();
    assert_eq!(cell.value, CellValue::Number(1462.0));
    let style = wb_1900.styles.get(cell.style_id).unwrap();
    assert_eq!(style.number_format.as_deref(), Some("m/d/yyyy"));

    // 1904 system: 1904-01-01 is serial 0.
    let mut wb_1904 = Workbook::new();
    wb_1904.date_system = DateSystem::Excel1904;
    let sheet_id = wb_1904.add_sheet("Sheet1").unwrap();
    apply_pivot_result_to_worksheet(
        &mut wb_1904,
        sheet_id,
        CellRef::new(0, 0),
        &result,
        &cfg,
        PivotApplyOptions::default(),
    )
    .unwrap();

    let cell = wb_1904
        .sheet(sheet_id)
        .unwrap()
        .cell(CellRef::new(1, 0))
        .unwrap();
    assert_eq!(cell.value, CellValue::Number(0.0));
    let style = wb_1904.styles.get(cell.style_id).unwrap();
    assert_eq!(style.number_format.as_deref(), Some("m/d/yyyy"));
}

#[test]
fn applies_value_field_number_format_when_present() {
    let result = PivotResult {
        data: vec![
            vec![
                PivotValue::Text("Region".to_string()),
                PivotValue::Text("Sum of Sales".to_string()),
            ],
            vec![
                PivotValue::Text("East".to_string()),
                PivotValue::Number(250.0),
            ],
        ],
    };

    let cfg = PivotConfig {
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
    };

    let mut wb = Workbook::new();
    let sheet_id = wb.add_sheet("Sheet1").unwrap();
    apply_pivot_result_to_worksheet(
        &mut wb,
        sheet_id,
        CellRef::new(0, 0),
        &result,
        &cfg,
        PivotApplyOptions::default(),
    )
    .unwrap();

    let cell = wb
        .sheet(sheet_id)
        .unwrap()
        .cell(CellRef::new(1, 1))
        .unwrap();
    assert_eq!(cell.value, CellValue::Number(250.0));
    let style = wb.styles.get(cell.style_id).unwrap();
    assert_eq!(style.number_format.as_deref(), Some("$#,##0.00"));
}

#[test]
fn applies_percent_format_for_percent_show_as_when_no_explicit_format() {
    let result = PivotResult {
        data: vec![
            vec![
                PivotValue::Text("Region".to_string()),
                PivotValue::Text("Sum of Sales".to_string()),
            ],
            vec![
                PivotValue::Text("East".to_string()),
                PivotValue::Number(0.5),
            ],
        ],
    };

    let cfg = PivotConfig {
        row_fields: vec![PivotField::new("Region")],
        column_fields: vec![],
        value_fields: vec![ValueField {
            source_field: PivotFieldRef::CacheFieldName("Sales".to_string()),
            name: "Sum of Sales".to_string(),
            aggregation: AggregationType::Sum,
            number_format: None,
            show_as: Some(formula_engine::pivot::ShowAsType::PercentOfGrandTotal),
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
    };

    let mut wb = Workbook::new();
    let sheet_id = wb.add_sheet("Sheet1").unwrap();
    apply_pivot_result_to_worksheet(
        &mut wb,
        sheet_id,
        CellRef::new(0, 0),
        &result,
        &cfg,
        PivotApplyOptions::default(),
    )
    .unwrap();

    let cell = wb
        .sheet(sheet_id)
        .unwrap()
        .cell(CellRef::new(1, 1))
        .unwrap();
    assert_eq!(cell.value, CellValue::Number(0.5));
    let style = wb.styles.get(cell.style_id).unwrap();
    assert_eq!(style.number_format.as_deref(), Some("0.00%"));
}
