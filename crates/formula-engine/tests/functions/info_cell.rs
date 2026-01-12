use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn cell_address_row_and_col() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=CELL(\"address\",A1)"),
        Value::Text("$A$1".to_string())
    );
    assert_number(&sheet.eval("=CELL(\"row\",A10)"), 10.0);
    assert_number(&sheet.eval("=CELL(\"col\",C1)"), 3.0);
}

#[test]
fn cell_type_codes_match_excel() {
    let mut sheet = TestSheet::new();

    // Blank.
    sheet.set("A1", Value::Blank);
    assert_eq!(sheet.eval("=CELL(\"type\",A1)"), Value::Text("b".to_string()));

    // Number.
    sheet.set("A1", 1.0);
    assert_eq!(sheet.eval("=CELL(\"type\",A1)"), Value::Text("v".to_string()));

    // Text.
    sheet.set("A1", "x");
    assert_eq!(sheet.eval("=CELL(\"type\",A1)"), Value::Text("l".to_string()));
}

#[test]
fn cell_contents_returns_formula_text_or_value() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", 5.0);
    assert_number(&sheet.eval("=CELL(\"contents\",A1)"), 5.0);

    sheet.set_formula("A1", "=1+1");
    assert_eq!(
        sheet.eval("=CELL(\"contents\",A1)"),
        Value::Text("=1+1".to_string())
    );
}

#[test]
fn info_recalc_and_unknown_keys() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=INFO(\"recalc\")"),
        // The engine defaults to manual calculation mode; callers can opt into Excel-like
        // automatic calculation by setting `CalcSettings.calculation_mode`.
        Value::Text("Manual".to_string())
    );
    assert_eq!(sheet.eval("=INFO(\"no_such_key\")"), Value::Error(ErrorKind::Value));
}

#[test]
fn info_numfile_counts_sheets() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet2", "A1", 1.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=INFO(\"numfile\")")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_number(&engine.get_cell_value("Sheet1", "B1"), 2.0);
}

#[test]
fn cell_errors_for_unknown_info_types() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=CELL(\"no_such_info_type\",A1)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn cell_filename_is_empty_for_unsaved_workbooks() {
    let mut sheet = TestSheet::new();

    // Excel returns "" until the workbook has been saved.
    assert_eq!(sheet.eval("=CELL(\"filename\")"), Value::Text(String::new()));
}

#[test]
fn cell_implicit_reference_does_not_create_dynamic_dependency_cycles() {
    let mut sheet = TestSheet::new();

    // Including INDIRECT marks the formula as dynamic-deps even though the IF short-circuits
    // and the INDIRECT branch is never evaluated.
    //
    // CELL("contents") with no explicit reference should not record a self-reference as a
    // dynamic precedent; otherwise the engine's dynamic dependency update can introduce a
    // self-edge and force the cell into circular-reference handling.
    let formula = "=IF(FALSE,INDIRECT(\"A1\"),CELL(\"contents\"))";
    assert_eq!(sheet.eval(formula), Value::Text(formula.to_string()));

    // Same idea, but for CELL("type") which also consults the referenced cell.
    assert_eq!(
        sheet.eval("=IF(FALSE,INDIRECT(\"A1\"),CELL(\"type\"))"),
        Value::Text("v".to_string())
    );
}

#[test]
fn cell_address_quotes_sheet_names_when_needed() {
    use formula_engine::Engine;

    let mut engine = Engine::new();
    engine.set_cell_value("My Sheet", "A1", 1.0).unwrap();
    engine.set_cell_value("A1", "A1", 1.0).unwrap();
    engine.set_cell_value("O'Brien", "A1", 1.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", "=CELL(\"address\",'My Sheet'!A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=CELL(\"address\",'A1'!A1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B3", "=CELL(\"address\",'O''Brien'!A1)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("'My Sheet'!$A$1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Text("'A1'!$A$1".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B3"),
        Value::Text("'O''Brien'!$A$1".to_string())
    );
}
