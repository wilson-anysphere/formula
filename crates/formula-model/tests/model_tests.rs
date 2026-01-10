use formula_model::{CellKey, CellRef, CellValue, ErrorValue, Range, Worksheet};

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
