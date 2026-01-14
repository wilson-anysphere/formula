use formula_engine::eval::CellAddr;
use formula_engine::{Engine, ErrorKind, ExternalValueProvider, Value};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};

#[derive(Default)]
struct TestExternalProvider {
    sheet_order: Mutex<HashMap<String, Vec<String>>>,
}

impl TestExternalProvider {
    fn set_sheet_order(&self, workbook: &str, order: impl Into<Vec<String>>) {
        self.sheet_order
            .lock()
            .expect("lock poisoned")
            .insert(workbook.to_string(), order.into());
    }
}

impl ExternalValueProvider for TestExternalProvider {
    fn get(&self, _sheet: &str, _addr: CellAddr) -> Option<Value> {
        None
    }

    fn sheet_order(&self, workbook: &str) -> Option<Vec<String>> {
        self.sheet_order
            .lock()
            .expect("lock poisoned")
            .get(workbook)
            .cloned()
    }
}

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
fn sheet_reports_external_sheet_number_when_order_available() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ],
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SHEET([Book.xlsx]Sheet2!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
}

#[test]
fn sheet_returns_na_for_external_sheet_when_order_unavailable() {
    let provider = Arc::new(TestExternalProvider::default());

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SHEET([Book.xlsx]Sheet2!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn sheet_reports_external_sheet_number_for_3d_span_argument() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ],
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SHEET([Book.xlsx]Sheet2:Sheet3!A1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
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
