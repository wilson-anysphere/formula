use formula_model::{
    validate_sheet_name, CellRef, DuplicateSheetError, Range, SheetNameError, Table, TableColumn,
    Workbook,
};

#[test]
fn duplicate_sheet_rewrites_explicit_self_references() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();

    wb.sheet_mut(sheet1)
        .unwrap()
        .set_formula_a1("B2", Some("=Sheet1!A1".to_string()))
        .unwrap();

    let copied = wb.duplicate_sheet(sheet1, None).unwrap();

    assert_eq!(wb.sheets.len(), 2);
    assert_eq!(wb.sheets[0].id, sheet1);
    assert_eq!(wb.sheets[1].id, copied);

    let copied_sheet = wb.sheet(copied).unwrap();
    assert_eq!(copied_sheet.name, "Sheet1 (2)");
    assert_eq!(
        copied_sheet.formula(CellRef::from_a1("B2").unwrap()),
        Some("'Sheet1 (2)'!A1")
    );

    // The source sheet is unchanged.
    let source_sheet = wb.sheet(sheet1).unwrap();
    assert_eq!(
        source_sheet.formula(CellRef::from_a1("B2").unwrap()),
        Some("Sheet1!A1")
    );
}

#[test]
fn duplicate_sheet_renames_tables_and_updates_structured_refs() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();

    let table = Table {
        id: 1,
        name: "Table1".to_string(),
        display_name: "Table1".to_string(),
        range: Range::from_a1("A1:B3").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Col1".to_string(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 2,
                name: "Col2".to_string(),
                formula: None,
                totals_formula: None,
            },
        ],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    };

    {
        let sheet = wb.sheet_mut(sheet1).unwrap();
        sheet.tables.push(table);
        sheet
            .set_formula_a1("C1", Some("=SUM(Table1[Col1])".to_string()))
            .unwrap();
    }

    let copied = wb.duplicate_sheet(sheet1, None).unwrap();
    let copied_sheet = wb.sheet(copied).unwrap();

    assert_eq!(copied_sheet.tables.len(), 1);
    assert_eq!(copied_sheet.tables[0].name, "Table1_1");
    assert_ne!(copied_sheet.tables[0].id, 1);

    assert_eq!(
        copied_sheet.formula(CellRef::from_a1("C1").unwrap()),
        Some("SUM(Table1_1[Col1])")
    );

    // The source sheet's table name and formula should be unchanged.
    let source_sheet = wb.sheet(sheet1).unwrap();
    assert_eq!(source_sheet.tables[0].name, "Table1");
    assert_eq!(
        source_sheet.formula(CellRef::from_a1("C1").unwrap()),
        Some("SUM(Table1[Col1])")
    );
}

#[test]
fn duplicate_sheet_name_collisions_match_excel_style_suffixes() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();

    let copy2 = wb.duplicate_sheet(sheet1, None).unwrap();
    assert_eq!(wb.sheet(copy2).unwrap().name, "Sheet1 (2)");

    let copy3 = wb.duplicate_sheet(sheet1, None).unwrap();
    assert_eq!(wb.sheet(copy3).unwrap().name, "Sheet1 (3)");
}

#[test]
fn duplicate_sheet_name_generation_respects_utf16_length_limit() {
    let mut wb = Workbook::new();
    let base = format!("{}A", "😀".repeat(15)); // 15 emoji (30 UTF-16) + 'A' = 31
    let sheet = wb.add_sheet(base).unwrap();

    let copied = wb.duplicate_sheet(sheet, None).unwrap();
    let copy_name = wb.sheet(copied).unwrap().name.clone();

    // Should not exceed Excel's 31 UTF-16 code unit limit.
    validate_sheet_name(&copy_name).unwrap();
    assert_eq!(copy_name.encode_utf16().count(), 30); // 13 emoji (26) + " (2)" (4)
    assert_eq!(copy_name, format!("{} (2)", "😀".repeat(13)));
}

#[test]
fn duplicate_sheet_rejects_duplicate_target_name() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();
    let _ = wb.add_sheet("Other").unwrap();

    let err = wb.duplicate_sheet(sheet1, Some("Other")).unwrap_err();
    assert_eq!(
        err,
        DuplicateSheetError::InvalidName(SheetNameError::DuplicateName)
    );
}
