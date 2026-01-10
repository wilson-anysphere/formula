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
