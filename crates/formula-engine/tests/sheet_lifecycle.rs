use formula_engine::Engine;
use formula_model::{CellRef, Range, Table, TableColumn};
use pretty_assertions::assert_eq;

fn cell(a1: &str) -> CellRef {
    CellRef::from_a1(a1).unwrap()
}

fn range(a1: &str) -> Range {
    Range::from_a1(a1).unwrap()
}

fn table_with_sheet_refs() -> Table {
    let formula = "Sheet1!A1+'[Book.xlsx]Sheet1'!A1".to_string();
    Table {
        id: 1,
        name: "Table1".to_string(),
        display_name: "Table1".to_string(),
        range: range("A1:B3"),
        header_row_count: 1,
        totals_row_count: 1,
        columns: vec![
            TableColumn {
                id: 1,
                name: "Col1".to_string(),
                formula: Some(formula.clone()),
                totals_formula: Some(formula),
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
    }
}

#[test]
fn delete_sheet_rewrites_local_refs_but_not_external_workbook_refs() {
    let mut engine = Engine::new();

    // Ensure both sheets exist so `Sheet1!A1` is compiled as an internal sheet reference.
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");

    engine
        .set_cell_formula("Sheet2", "A1", "=[Book.xlsx]Sheet1!A1+Sheet1!A1")
        .unwrap();

    engine.delete_sheet("Sheet1").unwrap();

    // Local references to the deleted sheet should become `#REF!`, but external workbook refs with
    // the same sheet name must remain intact.
    assert_eq!(
        engine.get_cell_formula("Sheet2", "A1").unwrap(),
        "=[Book.xlsx]Sheet1!A1+#REF!"
    );
}

#[test]
fn rename_sheet_rewrites_table_column_formulas_but_not_external_refs() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.set_sheet_tables("Sheet2", vec![table_with_sheet_refs()]);

    assert!(engine.rename_sheet("Sheet1", "Renamed"));

    let tables = engine.sheet_tables("Sheet2").unwrap();
    assert_eq!(tables.len(), 1);
    let table = &tables[0];

    assert_eq!(
        table.columns[0].formula.as_deref(),
        Some("Renamed!A1+'[Book.xlsx]Sheet1'!A1")
    );
    assert_eq!(
        table.columns[0].totals_formula.as_deref(),
        Some("Renamed!A1+'[Book.xlsx]Sheet1'!A1")
    );
}

#[test]
fn delete_sheet_invalidates_table_column_formulas_but_not_external_refs() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.set_sheet_tables("Sheet2", vec![table_with_sheet_refs()]);

    engine.delete_sheet("Sheet1").unwrap();

    let tables = engine.sheet_tables("Sheet2").unwrap();
    assert_eq!(tables.len(), 1);
    let table = &tables[0];

    assert_eq!(
        table.columns[0].formula.as_deref(),
        Some("#REF!+'[Book.xlsx]Sheet1'!A1")
    );
    assert_eq!(
        table.columns[0].totals_formula.as_deref(),
        Some("#REF!+'[Book.xlsx]Sheet1'!A1")
    );
}

#[test]
fn delete_sheet_adjusts_table_3d_boundary() {
    let mut engine = Engine::new();
    engine.ensure_sheet("Sheet1");
    engine.ensure_sheet("Sheet2");
    engine.ensure_sheet("Sheet3");

    let formula = "SUM(Sheet1:Sheet3!A1)".to_string();
    let table = Table {
        id: 1,
        name: "Table1".to_string(),
        display_name: "Table1".to_string(),
        range: Range::new(cell("A1"), cell("A3")),
        header_row_count: 1,
        totals_row_count: 1,
        columns: vec![TableColumn {
            id: 1,
            name: "Col1".to_string(),
            formula: Some(formula.clone()),
            totals_formula: Some(formula),
        }],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    };

    engine.set_sheet_tables("Sheet3", vec![table]);

    engine.delete_sheet("Sheet1").unwrap();

    let tables = engine.sheet_tables("Sheet3").unwrap();
    assert_eq!(tables.len(), 1);
    assert_eq!(
        tables[0].columns[0].formula.as_deref(),
        Some("SUM(Sheet2:Sheet3!A1)")
    );
}
