use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn sheet_reports_current_and_referenced_sheet_numbers() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();

    engine.set_cell_formula("Sheet1", "B1", "=SHEET()").unwrap();
    engine.set_cell_formula("Sheet2", "B1", "=SHEET()").unwrap();
    engine.set_cell_formula("Sheet2", "B2", "=SHEET(A1)").unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=SHEET(Sheet2!A1)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet2", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet2", "B2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(2.0));
}

#[test]
fn sheets_reports_workbook_sheet_count_and_3d_reference_span() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet3", "A1", 3.0).unwrap();

    engine.set_cell_formula("Sheet1", "B1", "=SHEETS()").unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=SHEETS(Sheet1:Sheet3!A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=SHEET(Sheet1:Sheet3!A1)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(1.0));
}

#[test]
fn formulatext_and_isformula_reflect_cell_formula_presence() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=1+1")
        .unwrap();
    engine.set_cell_value("Sheet1", "A2", 5.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=FORMULATEXT(A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=FORMULATEXT(A2)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=ISFORMULA(A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=ISFORMULA(A2)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("=1+1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Error(ErrorKind::NA)
    );

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Bool(false));
}

#[test]
fn normalize_formula_text_does_not_duplicate_equals_for_leading_whitespace_formulas() {
    assert_eq!(
        formula_engine::functions::information::workbook::normalize_formula_text(" =1+1"),
        " =1+1".to_string()
    )
}
