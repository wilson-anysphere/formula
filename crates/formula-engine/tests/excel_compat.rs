use formula_engine::{EditError, EditOp, Engine, ErrorKind, Value};
use formula_model::{Range, Table, TableColumn};

fn table_fixture() -> Table {
    Table {
        id: 1,
        name: "Table1".into(),
        display_name: "Table1".into(),
        range: Range::from_a1("A1:A3").unwrap(),
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
    }
}

fn unicode_table_fixture() -> Table {
    Table {
        id: 1,
        name: "Übersicht".into(),
        display_name: "Übersicht".into(),
        range: Range::from_a1("A1:A3").unwrap(),
        header_row_count: 1,
        totals_row_count: 0,
        columns: vec![TableColumn {
            id: 1,
            name: "Cöl".into(),
            formula: None,
            totals_formula: None,
        }],
        style: None,
        auto_filter: None,
        relationship_id: None,
        part_path: None,
    }
}

#[test]
fn sheet_references_are_case_insensitive() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 10.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=sheet1!A1")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(10.0));

    // Engine APIs should also resolve sheets case-insensitively.
    assert_eq!(engine.get_cell_value("sHeEt1", "A1"), Value::Number(10.0));
}

#[test]
fn sheet_references_are_case_insensitive_unicode() {
    let mut engine = Engine::new();
    engine.set_cell_value("Übersicht", "A1", 10.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "='übersicht'!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(10.0));
    assert_eq!(
        engine.get_cell_value("übersicht", "A1"),
        Value::Number(10.0)
    );
}

#[test]
fn missing_sheet_reference_evaluates_to_ref_error() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=NoSuchSheet!A1")
        .unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Ref)
    );

    // Referencing a missing sheet should not implicitly create it.
    let err = engine
        .apply_operation(EditOp::InsertRows {
            sheet: "NoSuchSheet".to_string(),
            row: 0,
            count: 1,
        })
        .unwrap_err();
    assert_eq!(err, EditError::SheetNotFound("NoSuchSheet".to_string()));
}

#[test]
fn structured_references_match_tables_and_columns_case_insensitively() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_sheet_tables("Sheet1", vec![table_fixture()]);

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(tAbLe1[cOl])")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
}

#[test]
fn structured_references_match_tables_and_columns_case_insensitively_unicode() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_sheet_tables("Sheet1", vec![unicode_table_fixture()]);

    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(übersicht[cÖL])")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
}

#[test]
fn text_comparisons_are_case_insensitive_unicode() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"="Straße"="STRASSE""#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", r#"="Straße"<>"STRASSE""#)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", r#"=IF("Straße"="STRASSE", 1, 0)"#)
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(1.0));
}
