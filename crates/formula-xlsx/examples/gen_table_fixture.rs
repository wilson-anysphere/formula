use formula_model::table::{AutoFilter, FilterColumn, Table, TableColumn, TableStyleInfo};
use formula_model::{FilterCriterion, FilterJoin, FilterValue};
use formula_model::{Cell, CellRef, CellValue, Range, Workbook};

fn main() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet("Sheet1").unwrap();
    let sheet = workbook.sheet_mut(sheet_id).unwrap();

    // Header row.
    sheet.set_cell(
        CellRef::from_a1("A1").unwrap(),
        Cell {
            value: CellValue::String("Item".into()),
            formula: None,
            phonetic: None,
            style_id: 0,
            phonetic: None,
        },
    );
    sheet.set_cell(
        CellRef::from_a1("B1").unwrap(),
        Cell {
            value: CellValue::String("Qty".into()),
            formula: None,
            phonetic: None,
            style_id: 0,
            phonetic: None,
        },
    );
    sheet.set_cell(
        CellRef::from_a1("C1").unwrap(),
        Cell {
            value: CellValue::String("Price".into()),
            formula: None,
            phonetic: None,
            style_id: 0,
            phonetic: None,
        },
    );
    sheet.set_cell(
        CellRef::from_a1("D1").unwrap(),
        Cell {
            value: CellValue::String("Total".into()),
            formula: None,
            phonetic: None,
            style_id: 0,
            phonetic: None,
        },
    );

    // Data rows.
    let rows = [
        ("Apple", 2.0, 3.0, 6.0),
        ("Banana", 1.0, 4.0, 4.0),
        ("Cherry", 5.0, 2.0, 10.0),
    ];
    for (idx, (item, qty, price, total)) in rows.iter().enumerate() {
        let row = (idx as u32) + 2;
        sheet.set_cell(
            CellRef::from_a1(&format!("A{row}")).unwrap(),
            Cell {
                value: CellValue::String((*item).into()),
                formula: None,
                phonetic: None,
                style_id: 0,
                phonetic: None,
            },
        );
        sheet.set_cell(
            CellRef::from_a1(&format!("B{row}")).unwrap(),
            Cell {
                value: CellValue::Number(*qty),
                formula: None,
                phonetic: None,
                style_id: 0,
                phonetic: None,
            },
        );
        sheet.set_cell(
            CellRef::from_a1(&format!("C{row}")).unwrap(),
            Cell {
                value: CellValue::Number(*price),
                formula: None,
                phonetic: None,
                style_id: 0,
                phonetic: None,
            },
        );
        sheet.set_cell(
            CellRef::from_a1(&format!("D{row}")).unwrap(),
            Cell {
                value: CellValue::Number(*total),
                formula: Some("[@Qty]*[@Price]".into()),
                phonetic: None,
                style_id: 0,
                phonetic: None,
            },
        );
    }

    // Structured reference formulas outside the table.
    sheet.set_cell(
        CellRef::from_a1("E1").unwrap(),
        Cell {
            value: CellValue::Number(20.0),
            formula: Some("SUM(Table1[Total])".into()),
            phonetic: None,
            style_id: 0,
            phonetic: None,
        },
    );

    sheet.set_cell(
        CellRef::from_a1("F1").unwrap(),
        Cell {
            value: CellValue::String("Qty".into()),
            formula: Some("Table1[[#Headers],[Qty]]".into()),
            phonetic: None,
            style_id: 0,
            phonetic: None,
        },
    );

    let table = Table {
        id: 1,
        name: "Table1".into(),
        display_name: "Table1".into(),
        range: Range::from_a1("A1:D4").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Item".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 2,
                name: "Qty".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 3,
                name: "Price".into(),
                formula: None,
                totals_formula: None,
            },
            TableColumn {
                id: 4,
                name: "Total".into(),
                formula: Some("[@Qty]*[@Price]".into()),
                totals_formula: None,
            },
        ],
        style: Some(TableStyleInfo {
            name: "TableStyleMedium2".into(),
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
        }),
        auto_filter: Some(AutoFilter {
            range: Range::from_a1("A1:D4").unwrap(),
            filter_columns: vec![FilterColumn {
                col_id: 0,
                join: FilterJoin::Any,
                criteria: vec![
                    FilterCriterion::Equals(FilterValue::Text("Apple".into())),
                    FilterCriterion::Equals(FilterValue::Text("Cherry".into())),
                ],
                values: vec!["Apple".into(), "Cherry".into()],
                raw_xml: Vec::new(),
            }],
            sort_state: None,
            raw_xml: Vec::new(),
        }),
        relationship_id: Some("rId1".into()),
        part_path: None,
    };
    sheet.tables.push(table);

    let out_path =
        std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/table.xlsx");
    std::fs::create_dir_all(out_path.parent().unwrap()).unwrap();
    formula_xlsx::write_workbook(&workbook, &out_path).unwrap();
}
