use formula_model::{CellRef, CellValue, StyleTable, Worksheet, EXCEL_MAX_ROWS};

#[test]
fn default_sheet_row_count_is_excel_max() {
    let sheet = Worksheet::new(1, "Sheet1");
    let styles = StyleTable::new();

    assert_eq!(sheet.row_count, EXCEL_MAX_ROWS);
    assert!(sheet.is_cell_editable(
        CellRef::new(EXCEL_MAX_ROWS - 1, 0),
        &styles
    ));
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

