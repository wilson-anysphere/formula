use formula_model::{
    CellKey, CellRef, CellValue, ErrorValue, Range, Table, TableColumn, Workbook, Worksheet,
    EXCEL_MAX_COLS,
    EXCEL_MAX_ROWS, SCHEMA_VERSION,
};

#[test]
fn sparse_storage_is_proportional_to_stored_cells() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    // Empty writes should not allocate a stored record.
    sheet.set_value(CellRef::new(10, 10), CellValue::Empty);
    assert_eq!(sheet.cell_count(), 0);

    sheet.set_value(CellRef::new(0, 0), CellValue::Number(1.0));
    sheet.set_value(CellRef::new(999_999, 999), CellValue::Number(2.0));
    assert_eq!(sheet.cell_count(), 2);

    // Clearing removes storage again.
    sheet.clear_cell(CellRef::new(0, 0));
    assert_eq!(sheet.cell_count(), 1);
}

#[test]
fn used_range_updates_on_set_and_clear() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    assert_eq!(sheet.used_range(), None);

    let a = CellRef::new(5, 2);
    let b = CellRef::new(1, 10);

    sheet.set_value(a, CellValue::Number(1.0));
    assert_eq!(sheet.used_range(), Some(Range::new(a, a)));

    sheet.set_value(b, CellValue::Boolean(true));
    assert_eq!(
        sheet.used_range(),
        Some(Range::new(CellRef::new(1, 2), CellRef::new(5, 10)))
    );

    // Removing a non-boundary cell should not shrink the used range.
    sheet.set_value(CellRef::new(3, 5), CellValue::String("x".into()));
    sheet.clear_cell(CellRef::new(3, 5));
    assert_eq!(
        sheet.used_range(),
        Some(Range::new(CellRef::new(1, 2), CellRef::new(5, 10)))
    );

    // Removing a boundary cell triggers recomputation.
    sheet.clear_cell(b);
    assert_eq!(sheet.used_range(), Some(Range::new(a, a)));

    sheet.clear_cell(a);
    assert_eq!(sheet.used_range(), None);
}

#[test]
fn used_range_is_recomputed_on_deserialize() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    let a = CellRef::new(5, 2);
    let b = CellRef::new(1, 10);
    sheet.set_value(a, CellValue::Number(1.0));
    sheet.set_value(b, CellValue::Boolean(true));

    let mut json = serde_json::to_value(&sheet).unwrap();
    let obj = json.as_object_mut().unwrap();
    obj.remove("used_range");

    let deserialized: Worksheet = serde_json::from_value(json).unwrap();
    assert_eq!(deserialized.used_range(), sheet.used_range());
}

#[test]
fn worksheet_a1_helpers_work() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    sheet.set_value_a1("B2", CellValue::Number(3.0)).unwrap();

    assert_eq!(sheet.value(CellRef::new(1, 1)), CellValue::Number(3.0));
    assert_eq!(sheet.value_a1("B2").unwrap(), CellValue::Number(3.0));
    assert_eq!(sheet.cell_a1("B2").unwrap().is_some(), true);

    sheet.clear_cell_a1("$B$2").unwrap();
    assert_eq!(sheet.value_a1("B2").unwrap(), CellValue::Empty);
}

#[test]
fn formula_and_style_keep_empty_cells_in_sparse_store() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    // Formula-only cell is stored.
    sheet
        .set_formula_a1("A1", Some("=1+1".to_string()))
        .unwrap();
    assert_eq!(sheet.cell_count(), 1);
    assert_eq!(sheet.formula(CellRef::new(0, 0)), Some("=1+1"));

    // Clearing formula removes cell again (since it's otherwise empty/default).
    sheet.set_formula(CellRef::new(0, 0), None);
    assert_eq!(sheet.cell_count(), 0);

    // Styled empty cell is stored.
    sheet.set_style_id_a1("B2", 42).unwrap();
    assert_eq!(sheet.cell_count(), 1);
    assert_eq!(sheet.cell_a1("B2").unwrap().unwrap().style_id, 42);

    // Resetting to default style drops the cell.
    sheet.set_style_id(CellRef::new(1, 1), 0);
    assert_eq!(sheet.cell_count(), 0);
}

#[test]
fn row_and_col_properties_are_deduped() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    assert!(sheet.row_properties(10).is_none());
    sheet.set_row_height(10, Some(12.5));
    assert_eq!(sheet.row_properties(10).unwrap().height, Some(12.5));

    // Clearing the override removes the entry.
    sheet.set_row_height(10, None);
    assert!(sheet.row_properties(10).is_none());

    sheet.set_row_hidden(10, true);
    assert_eq!(sheet.row_properties(10).unwrap().hidden, true);

    sheet.set_row_hidden(10, false);
    assert!(sheet.row_properties(10).is_none());

    assert!(sheet.col_properties(3).is_none());
    sheet.set_col_width(3, Some(8.0));
    assert_eq!(sheet.col_properties(3).unwrap().width, Some(8.0));
    sheet.set_col_width(3, None);
    assert!(sheet.col_properties(3).is_none());

    sheet.set_col_hidden(3, true);
    assert!(sheet.col_properties(3).unwrap().hidden);
    sheet.set_col_hidden(3, false);
    assert!(sheet.col_properties(3).is_none());
}

#[test]
fn clear_range_removes_cells_and_updates_used_range() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    // Set 3 cells in a 3x3 block.
    sheet.set_value_a1("A1", CellValue::Number(1.0)).unwrap();
    sheet.set_value_a1("B2", CellValue::Number(2.0)).unwrap();
    sheet.set_value_a1("C3", CellValue::Number(3.0)).unwrap();

    assert_eq!(sheet.used_range(), Some(Range::from_a1("A1:C3").unwrap()));

    // Clear a subrange that removes the bottom-right corner.
    sheet.clear_range(Range::from_a1("B2:C3").unwrap());
    assert_eq!(sheet.cell_count(), 1);
    assert_eq!(sheet.value_a1("A1").unwrap(), CellValue::Number(1.0));
    assert_eq!(sheet.used_range(), Some(Range::from_a1("A1").unwrap()));
}

#[test]
fn iter_cells_in_range_filters_sparse_cells() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    sheet.set_value_a1("A1", CellValue::Number(1.0)).unwrap();
    sheet.set_value_a1("D4", CellValue::Number(2.0)).unwrap();
    sheet.set_value_a1("F1", CellValue::Number(3.0)).unwrap();

    let cells: Vec<_> = sheet
        .iter_cells_in_range(Range::from_a1("A1:E4").unwrap())
        .map(|(r, _)| r.to_a1())
        .collect();
    assert_eq!(cells.len(), 2);
    assert!(cells.contains(&"A1".to_string()));
    assert!(cells.contains(&"D4".to_string()));
}

#[test]
fn worksheet_visibility_serializes_only_when_non_default() {
    let mut sheet = Worksheet::new(1, "Sheet1");
    let json = serde_json::to_value(&sheet).unwrap();
    assert!(json.get("visibility").is_none());

    sheet.visibility = formula_model::SheetVisibility::Hidden;
    let json = serde_json::to_value(&sheet).unwrap();
    assert_eq!(json.get("visibility").unwrap(), "hidden");
}

#[test]
fn cell_key_encoding_round_trips() {
    for &(row, col) in &[
        (0, 0),
        (1, 1),
        (123, 456),
        (1_048_575, 16_383), // Excel max (0-indexed)
    ] {
        let key = CellKey::new(row, col);
        assert_eq!(key.row(), row);
        assert_eq!(key.col(), col);
        assert_eq!(CellKey::from_ref(key.to_ref()).as_u64(), key.as_u64());
    }
}

#[test]
fn workbook_sheet_by_name_is_case_insensitive() {
    let mut workbook = Workbook::new();
    workbook.add_sheet("Sheet1");
    assert!(workbook.sheet_by_name("sheet1").is_some());
    assert!(workbook.sheet_by_name("SHEET1").is_some());
}

#[test]
fn workbook_find_table_is_case_insensitive() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1");
    let sheet = workbook.sheet_mut(sheet_id).unwrap();
    sheet.tables.push(Table {
        id: 1,
        name: "Table1".into(),
        display_name: "Table1".into(),
        range: Range::from_a1("A1:A2").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![TableColumn {
            id: 1,
            name: "Col".into(),
            formula: None,
            totals_formula: None,
        }],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    });

    assert!(workbook.find_table("table1").is_some());
    assert!(workbook.find_table("TABLE1").is_some());
}

#[test]
fn error_strings_match_excel_spellings() {
    assert_eq!(ErrorValue::Div0.to_string(), "#DIV/0!");
    assert_eq!(ErrorValue::Name.to_string(), "#NAME?");
    assert_eq!("#SPILL!".parse::<ErrorValue>().unwrap(), ErrorValue::Spill);
    assert_eq!(
        "#GETTING_DATA".parse::<ErrorValue>().unwrap(),
        ErrorValue::GettingData
    );
}

#[test]
fn serde_schema_for_cell_values_is_stable() {
    let v = serde_json::to_value(CellValue::Number(1.5)).unwrap();
    assert_eq!(v, serde_json::json!({ "type": "number", "value": 1.5 }));

    let v = serde_json::to_value(CellValue::Error(ErrorValue::Div0)).unwrap();
    assert_eq!(
        v,
        serde_json::json!({ "type": "error", "value": "#DIV/0!" })
    );
}

#[test]
fn workbook_schema_version_is_enforced() {
    let wb: Workbook = serde_json::from_value(serde_json::json!({})).unwrap();
    assert_eq!(wb.schema_version, SCHEMA_VERSION);

    let err = serde_json::from_value::<Workbook>(serde_json::json!({
        "schema_version": SCHEMA_VERSION + 1
    }))
    .unwrap_err();
    assert!(err.to_string().contains("unsupported schema_version"));
}

#[test]
fn worksheet_deserialize_validates_dimensions() {
    let err = serde_json::from_value::<Worksheet>(serde_json::json!({
        "id": 1,
        "name": "Sheet1",
        "row_count": 0,
        "col_count": 1
    }))
    .unwrap_err();
    assert!(err.to_string().contains("row_count"));

    let err = serde_json::from_value::<Worksheet>(serde_json::json!({
        "id": 1,
        "name": "Sheet1",
        "row_count": EXCEL_MAX_ROWS + 1,
        "col_count": EXCEL_MAX_COLS + 1
    }))
    .unwrap_err();
    assert!(err.to_string().contains("out of Excel bounds"));
}
