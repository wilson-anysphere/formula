use formula_model::{validate_table_name, Range, Table, TableColumn, TableError, Workbook};

fn table_fixture(id: u32, name: &str) -> Table {
    Table {
        id,
        name: name.to_string(),
        display_name: name.to_string(),
        range: Range::from_a1("A1:C3").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Col1".into(),
                formula: Some("SUM(Table1[Col1])".into()),
                totals_formula: None,
            },
            TableColumn {
                id: 2,
                name: "Col2".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 3,
                name: "Col3".into(),
                formula: None,
                totals_formula: None,
            },
        ],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    }
}

#[test]
fn validates_table_names() {
    assert_eq!(validate_table_name("").unwrap_err(), TableError::EmptyName);
    assert_eq!(
        validate_table_name("1Table").unwrap_err(),
        TableError::InvalidStartChar
    );
    assert!(matches!(
        validate_table_name("Table 1"),
        Err(TableError::InvalidChar { .. })
    ));
    assert_eq!(
        validate_table_name("A1").unwrap_err(),
        TableError::ConflictsWithCellReference
    );
    assert_eq!(
        validate_table_name("R1C1").unwrap_err(),
        TableError::ConflictsWithCellReference
    );
    assert_eq!(
        validate_table_name("C").unwrap_err(),
        TableError::ReservedName
    );
    assert_eq!(
        validate_table_name("TRUE").unwrap_err(),
        TableError::ReservedName
    );
    validate_table_name("Table1").unwrap();
}

#[test]
fn workbook_enforces_unique_table_names_case_insensitive() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();
    let sheet2 = wb.add_sheet("Sheet2").unwrap();

    wb.add_table(sheet1, table_fixture(1, "Table1")).unwrap();
    assert_eq!(
        wb.add_table(sheet2, table_fixture(2, "table1"))
            .unwrap_err(),
        TableError::DuplicateName
    );
}

#[test]
fn rename_table_rewrites_structured_refs_in_formulas() {
    let mut wb = Workbook::new();
    let sheet1 = wb.add_sheet("Sheet1").unwrap();

    wb.add_table(sheet1, table_fixture(1, "Table1")).unwrap();

    {
        let sheet = wb.sheet_mut(sheet1).unwrap();
        sheet
            .set_formula_a1("A1", Some("=SUM(table1[Col1])".into()))
            .unwrap();
        sheet
            .set_formula_a1("A2", Some("=SUM(Sheet1!TABLE1[Col1])".into()))
            .unwrap();
        sheet
            .set_formula_a1("A3", Some("=\"table1[Col1]\"".into()))
            .unwrap();
    }

    wb.rename_table("Table1", "Sales").unwrap();

    let sheet = wb.sheet(sheet1).unwrap();
    assert_eq!(sheet.formula_a1("A1").unwrap(), Some("SUM(Sales[Col1])"));
    assert_eq!(
        sheet.formula_a1("A2").unwrap(),
        Some("SUM(Sheet1!Sales[Col1])")
    );
    assert_eq!(sheet.formula_a1("A3").unwrap(), Some("\"table1[Col1]\""));

    let (_sheet, table) = wb.find_table_case_insensitive("sales").unwrap();
    assert_eq!(table.name, "Sales");
    assert_eq!(table.display_name, "Sales");
    assert_eq!(
        table.columns[0].formula.as_deref(),
        Some("SUM(Sales[Col1])")
    );
}

#[test]
fn table_column_index_is_case_insensitive_for_unicode_text() {
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
                name: "Maß".to_string(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 2,
                name: "Other".to_string(),
                formula: None,
                totals_formula: None,
            },
        ],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    };

    // Uses Unicode-aware uppercasing: ß -> SS.
    assert_eq!(table.column_index("MASS"), Some(0));
}

#[test]
fn table_set_range_resizes_columns() {
    let mut table = table_fixture(1, "Table1");

    table.set_range(Range::from_a1("A1:D3").unwrap()).unwrap();
    assert_eq!(table.range.to_string(), "A1:D3");
    assert_eq!(table.columns.len(), 4);
    assert_eq!(table.columns[3].id, 4);

    table.set_range(Range::from_a1("A1:B3").unwrap()).unwrap();
    assert_eq!(table.range.to_string(), "A1:B3");
    assert_eq!(table.columns.len(), 2);
    assert_eq!(table.columns[1].id, 2);
}

#[test]
fn table_set_range_validates_header_totals_rows() {
    let mut table = table_fixture(1, "Table1");
    table.header_row_count = 1;
    table.totals_row_count = 1;

    assert_eq!(
        table
            .set_range(Range::from_a1("A1:C1").unwrap())
            .unwrap_err(),
        TableError::InvalidRange
    );
}
