use formula_model::{
    Cell, CellKey, CellRef, CellValue, DataValidation, DataValidationAssignment,
    DataValidationKind, ErrorValue, Hyperlink, HyperlinkTarget, Range, Table, TableColumn,
    Workbook, Worksheet, EXCEL_MAX_COLS, EXCEL_MAX_ROWS, SCHEMA_VERSION,
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
    sheet.set_formula_a1("A1", Some("1+1".to_string())).unwrap();
    assert_eq!(sheet.cell_count(), 1);
    assert_eq!(sheet.formula(CellRef::new(0, 0)), Some("1+1"));

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
fn cell_phonetic_roundtrips_through_workbook_json() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").unwrap();
    let sheet = workbook.sheet_mut(sheet_id).unwrap();

    let mut cell = Cell::new(CellValue::String("漢字".to_string()));
    cell.phonetic = Some("かんじ".to_string());
    sheet.set_cell(CellRef::new(0, 0), cell);

    let json = serde_json::to_string(&workbook).unwrap();
    assert!(
        json.contains("\"phonetic\""),
        "expected serialized workbook to include phonetic metadata"
    );

    let decoded: Workbook = serde_json::from_str(&json).unwrap();
    let decoded_sheet = decoded.sheet_by_name("Sheet1").unwrap();
    let decoded_cell = decoded_sheet.cell(CellRef::new(0, 0)).unwrap();
    assert_eq!(decoded_cell.phonetic_text(), Some("かんじ"));

    // Back-compat: older payloads without the `phonetic` key must deserialize.
    let legacy: Cell = serde_json::from_str("{}").unwrap();
    assert_eq!(legacy.phonetic_text(), None);
    let legacy_json = serde_json::to_string(&legacy).unwrap();
    assert!(
        !legacy_json.contains("\"phonetic\""),
        "expected phonetic key to be omitted when None"
    );
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
fn row_and_col_style_id_overrides_are_deduped_and_persist() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    // Row style creates/removes entries.
    assert!(sheet.row_properties(5).is_none());
    sheet.set_row_style_id(5, Some(42));
    let props = sheet.row_properties(5).unwrap();
    assert_eq!(props.style_id, Some(42));
    assert_eq!(props.height, None);
    assert!(!props.hidden);

    sheet.set_row_style_id(5, None);
    assert!(sheet.row_properties(5).is_none());

    // Column style creates/removes entries.
    assert!(sheet.col_properties(7).is_none());
    sheet.set_col_style_id(7, Some(24));
    let props = sheet.col_properties(7).unwrap();
    assert_eq!(props.style_id, Some(24));
    assert_eq!(props.width, None);
    assert!(!props.hidden);

    sheet.set_col_style_id(7, None);
    assert!(sheet.col_properties(7).is_none());

    // Style overrides should not be dropped when other overrides are cleared.
    let row = 10;
    sheet.set_row_style_id(row, Some(1));
    sheet.set_row_height(row, Some(12.5));
    sheet.set_row_height(row, None);
    assert_eq!(sheet.row_properties(row).unwrap().style_id, Some(1));

    sheet.set_row_hidden(row, true);
    sheet.set_row_hidden(row, false);
    assert_eq!(sheet.row_properties(row).unwrap().style_id, Some(1));

    sheet.set_row_style_id(row, None);
    assert!(sheet.row_properties(row).is_none());

    let col = 3;
    sheet.set_col_style_id(col, Some(2));
    sheet.set_col_width(col, Some(8.0));
    sheet.set_col_width(col, None);
    assert_eq!(sheet.col_properties(col).unwrap().style_id, Some(2));

    sheet.set_col_hidden(col, true);
    sheet.set_col_hidden(col, false);
    assert_eq!(sheet.col_properties(col).unwrap().style_id, Some(2));

    sheet.set_col_style_id(col, None);
    assert!(sheet.col_properties(col).is_none());
}

#[test]
fn row_and_col_style_id_overrides_treat_zero_as_clear() {
    let mut sheet = Worksheet::new(1, "Sheet1");

    sheet.set_row_style_id(5, Some(42));
    assert_eq!(sheet.row_properties(5).unwrap().style_id, Some(42));
    sheet.set_row_style_id(5, Some(0));
    assert!(
        sheet.row_properties(5).is_none(),
        "expected Some(0) to clear row style"
    );

    sheet.set_col_style_id(7, Some(24));
    assert_eq!(sheet.col_properties(7).unwrap().style_id, Some(24));
    sheet.set_col_style_id(7, Some(0));
    assert!(
        sheet.col_properties(7).is_none(),
        "expected Some(0) to clear col style"
    );
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
    workbook.add_sheet("Sheet1").unwrap();
    assert!(workbook.sheet_by_name("sheet1").is_some());
    assert!(workbook.sheet_by_name("SHEET1").is_some());
}

#[test]
fn workbook_sheet_by_name_is_unicode_case_insensitive() {
    let mut workbook = Workbook::new();
    workbook.add_sheet("Äbc").unwrap();
    assert!(workbook.sheet_by_name("äbc").is_some());
}

#[test]
fn workbook_find_table_is_case_insensitive() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").unwrap();
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
fn table_column_index_is_case_insensitive() {
    let table = Table {
        id: 1,
        name: "Table1".into(),
        display_name: "Table1".into(),
        range: Range::from_a1("A1:B2").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Col".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 2,
                name: "Other".into(),
                formula: None,
                totals_formula: None,
            },
        ],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    };

    assert_eq!(table.column_index("col"), Some(0));
    assert_eq!(table.column_index("COL"), Some(0));
    assert_eq!(table.column_index("oThEr"), Some(1));
}

#[test]
fn rename_sheet_rewrites_non_cell_formulas_and_hyperlinks() {
    let mut workbook = Workbook::new();
    let renamed_id = workbook.add_sheet("Sheet1").unwrap();
    let other_id = workbook.add_sheet("Other").unwrap();

    let other = workbook.sheet_mut(other_id).unwrap();
    other.tables.push(Table {
        id: 1,
        name: "Table1".into(),
        display_name: "Table1".into(),
        range: Range::from_a1("A1:B2").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![TableColumn {
            id: 1,
            name: "Col".into(),
            formula: Some("Sheet1!A1".into()),
            totals_formula: Some("Sheet1!B1".into()),
        }],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    });
    other.data_validations.push(DataValidationAssignment {
        id: 1,
        ranges: vec![Range::from_a1("A1").unwrap()],
        validation: DataValidation {
            kind: DataValidationKind::List,
            operator: None,
            formula1: "Sheet1!A1:A3".into(),
            formula2: None,
            allow_blank: false,
            show_input_message: false,
            show_error_message: false,
            show_drop_down: false,
            input_message: None,
            error_alert: None,
        },
    });
    other.hyperlinks.push(Hyperlink::for_cell(
        CellRef::new(0, 0),
        HyperlinkTarget::Internal {
            sheet: "Sheet1".into(),
            cell: CellRef::new(0, 0),
        },
    ));

    workbook.rename_sheet(renamed_id, "Data").unwrap();

    let other = workbook.sheet(other_id).unwrap();
    let col = &other.tables[0].columns[0];
    assert_eq!(col.formula.as_deref(), Some("Data!A1"));
    assert_eq!(col.totals_formula.as_deref(), Some("Data!B1"));
    assert_eq!(other.data_validations[0].validation.formula1, "Data!A1:A3");
    match &other.hyperlinks[0].target {
        HyperlinkTarget::Internal { sheet, cell } => {
            assert_eq!(sheet, "Data");
            assert_eq!(*cell, CellRef::new(0, 0));
        }
        _ => panic!("expected internal hyperlink"),
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
