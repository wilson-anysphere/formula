use formula_model::autofilter::{DateComparison, FilterColumn, FilterCriterion, FilterJoin, NumberComparison, SortCondition, SortState};
use formula_model::table::{Table, TableColumn};
use formula_model::{Cell, CellRef, CellValue, Range, SheetAutoFilter, Workbook};
use formula_xlsx::{read_workbook, write_workbook};
use tempfile::tempdir;

#[test]
fn table_autofilter_advanced_roundtrips_through_read_write_workbook() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").unwrap();
    let sheet = workbook.sheet_mut(sheet_id).unwrap();

    // Create a minimal grid for the table range.
    sheet.set_cell(CellRef::from_a1("A1").unwrap(), Cell::new(CellValue::String("A".into())));
    sheet.set_cell(CellRef::from_a1("B1").unwrap(), Cell::new(CellValue::String("B".into())));
    sheet.set_cell(CellRef::from_a1("C1").unwrap(), Cell::new(CellValue::String("C".into())));
    sheet.set_cell(CellRef::from_a1("A2").unwrap(), Cell::new(CellValue::Number(1.0)));
    sheet.set_cell(CellRef::from_a1("B2").unwrap(), Cell::new(CellValue::Number(2.0)));
    sheet.set_cell(CellRef::from_a1("C2").unwrap(), Cell::new(CellValue::Number(3.0)));

    let table_range = Range::from_a1("A1:C10").unwrap();
    let filter = SheetAutoFilter {
        range: table_range,
        filter_columns: vec![
            FilterColumn {
                col_id: 0,
                join: FilterJoin::All,
                criteria: vec![FilterCriterion::Number(NumberComparison::GreaterThan(10.0))],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            FilterColumn {
                col_id: 1,
                join: FilterJoin::Any,
                criteria: vec![FilterCriterion::Date(DateComparison::Today)],
                values: Vec::new(),
                raw_xml: Vec::new(),
            },
            // Unsupported filter criteria stored as raw XML should survive read/write.
            FilterColumn {
                col_id: 2,
                join: FilterJoin::Any,
                criteria: Vec::new(),
                values: Vec::new(),
                raw_xml: vec![r#"<top10 val="5" percent="1"/>"#.to_string()],
            },
        ],
        sort_state: Some(SortState {
            conditions: vec![SortCondition {
                range: Range::from_a1("A2:A10").unwrap(),
                descending: true,
            }],
        }),
        raw_xml: vec![r#"<extLst><ext uri="{00000000-0000-0000-0000-000000000000}"/></extLst>"#
            .to_string()],
    };

    let table = Table {
        id: 1,
        name: "TableAdvanced".to_string(),
        display_name: "TableAdvanced".to_string(),
        range: table_range,
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![
            TableColumn {
                id: 1,
                name: "A".to_string(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 2,
                name: "B".to_string(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 3,
                name: "C".to_string(),
                formula: None,
                totals_formula: None,
            },
        ],
        style: None,
        auto_filter: Some(filter),
        // Match the writer's default values so a full model equality check holds across
        // read/write.
        relationship_id: Some("rId1".to_string()),
        part_path: Some("xl/tables/table1.xml".to_string()),
    };

    workbook.add_table(sheet_id, table).unwrap();

    let dir = tempdir().unwrap();
    let out_path = dir.path().join("table-autofilter-advanced.xlsx");
    write_workbook(&workbook, &out_path).unwrap();

    let loaded = read_workbook(&out_path).unwrap();
    let loaded_sheet = &loaded.sheets[0];
    assert_eq!(loaded_sheet.tables, workbook.sheets[0].tables);
}

