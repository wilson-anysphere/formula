use formula_model::{
    CellRef, CellValue, DataValidation, DataValidationKind, Range, StyleTable, Worksheet,
    EXCEL_MAX_ROWS,
};

#[test]
fn default_sheet_row_count_is_excel_max() {
    let sheet = Worksheet::new(1, "Sheet1");
    let styles = StyleTable::new();

    assert_eq!(sheet.row_count, EXCEL_MAX_ROWS);
    assert!(sheet.is_cell_editable(CellRef::new(EXCEL_MAX_ROWS - 1, 0), &styles));
    assert!(!sheet.is_cell_editable(CellRef::new(EXCEL_MAX_ROWS, 0), &styles));
}

#[test]
fn worksheet_can_edit_cells_beyond_excel_max_when_row_count_is_larger() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    let styles = StyleTable::new();
    let beyond_excel = CellRef::new(EXCEL_MAX_ROWS, 0);

    // Default sheet bounds are Excel-like.
    assert!(!sheet.is_cell_editable(beyond_excel, &styles));

    // Growing the sheet's dimensions beyond Excel's limit makes higher rows editable.
    sheet.set_row_height(EXCEL_MAX_ROWS, Some(12.0));
    assert!(sheet.row_count > EXCEL_MAX_ROWS);
    assert!(sheet.is_cell_editable(beyond_excel, &styles));

    sheet.set_value(beyond_excel, CellValue::Number(123.0));
    assert_eq!(sheet.value(beyond_excel), CellValue::Number(123.0));
}

#[test]
fn inserting_cells_beyond_excel_max_rows_grows_row_count_and_updates_editability() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    let styles = StyleTable::new();
    let cell = CellRef::new(EXCEL_MAX_ROWS + 5, 0);

    assert!(!sheet.is_cell_editable(cell, &styles));

    sheet.set_value(cell, CellValue::Number(1.0));
    assert!(sheet.row_count > EXCEL_MAX_ROWS);
    assert!(sheet.is_cell_editable(cell, &styles));
    assert_eq!(sheet.value(cell), CellValue::Number(1.0));
}

#[test]
fn data_validations_apply_beyond_excel_max_rows_when_sheet_row_count_is_larger() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    let cell = CellRef::new(EXCEL_MAX_ROWS, 0);

    // Grow the sheet so the row is inside the worksheet's dimensions.
    sheet.set_row_height(EXCEL_MAX_ROWS, Some(12.0));

    let id = sheet.add_data_validation(
        vec![Range::new(cell, cell)],
        DataValidation {
            kind: DataValidationKind::List,
            operator: None,
            formula1: "\"a,b\"".into(),
            formula2: None,
            allow_blank: false,
            show_input_message: false,
            show_error_message: false,
            show_drop_down: false,
            input_message: None,
            error_alert: None,
        },
    );

    let dvs = sheet.data_validations_for_cell(cell);
    assert_eq!(dvs.len(), 1);
    assert_eq!(dvs[0].id, id);
}

#[test]
fn worksheet_deserialize_allows_row_count_above_excel_max_and_updates_editability() {
    let sheet: Worksheet = serde_json::from_value(serde_json::json!({
        "id": 1,
        "name": "Sheet1",
        "row_count": EXCEL_MAX_ROWS + 10,
        "col_count": 1,
    }))
    .unwrap();

    let styles = StyleTable::new();
    assert!(sheet.is_cell_editable(CellRef::new(EXCEL_MAX_ROWS, 0), &styles));
    assert!(!sheet.is_cell_editable(CellRef::new(EXCEL_MAX_ROWS + 10, 0), &styles));
}
