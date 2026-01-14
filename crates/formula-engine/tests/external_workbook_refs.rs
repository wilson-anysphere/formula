use formula_engine::eval::CellAddr;
use formula_engine::{Engine, ExternalValueProvider, PrecedentNode, Value};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};

#[derive(Default)]
struct TestExternalProvider {
    values: Mutex<HashMap<(String, CellAddr), Value>>,
    sheet_order: Mutex<HashMap<String, Vec<String>>>,
}

impl TestExternalProvider {
    fn set(&self, sheet: &str, addr: CellAddr, value: impl Into<Value>) {
        self.values
            .lock()
            .expect("lock poisoned")
            .insert((sheet.to_string(), addr), value.into());
    }

    fn set_sheet_order(&self, workbook: &str, order: impl Into<Vec<String>>) {
        self.sheet_order
            .lock()
            .expect("lock poisoned")
            .insert(workbook.to_string(), order.into());
    }
}

impl ExternalValueProvider for TestExternalProvider {
    fn get(&self, sheet: &str, addr: CellAddr) -> Option<Value> {
        self.values
            .lock()
            .expect("lock poisoned")
            .get(&(sheet.to_string(), addr))
            .cloned()
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
fn external_cell_ref_resolves_via_provider() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        41.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
}

#[test]
fn external_cell_ref_participates_in_arithmetic() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        41.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1+1")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(42.0));
}

#[test]
fn sum_over_external_range_uses_reference_semantics() {
    // Excel quirk: SUM over references ignores logicals/text stored in cells.
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        1.0,
    );
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 1, col: 0 },
        Value::Text("2".to_string()),
    );
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 2, col: 0 },
        Value::Bool(true),
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet1!A1:A3)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
}

#[test]
fn external_range_spills_to_grid() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        1.0,
    );
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 1 },
        2.0,
    );
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 1, col: 0 },
        3.0,
    );
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 1, col: 1 },
        4.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "C1", "=[Book.xlsx]Sheet1!A1:B2")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.spill_range("Sheet1", "C1"),
        Some((
            CellAddr { row: 0, col: 2 },
            CellAddr { row: 1, col: 3 }
        ))
    );
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(4.0));
}

#[test]
fn missing_external_value_is_ref_error() {
    let provider = Arc::new(TestExternalProvider::default());

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );
}

#[test]
fn precedents_include_external_refs() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "B1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();

    assert_eq!(
        engine.precedents("Sheet1", "B1").unwrap(),
        vec![PrecedentNode::ExternalCell {
            sheet: "[Book.xlsx]Sheet1".to_string(),
            addr: CellAddr { row: 0, col: 0 },
        }]
    );
}

#[test]
fn external_refs_are_volatile() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        1.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider.clone()));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1!A1")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));

    // Mutate the provider without marking any cells dirty. Volatile external references should be
    // included in subsequent recalculation passes.
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        2.0,
    );
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
}

#[test]
fn degenerate_external_3d_sheet_range_ref_resolves_via_provider() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Book.xlsx]Sheet1",
        CellAddr { row: 0, col: 0 },
        41.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1:Sheet1!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(41.0));
}

#[test]
fn external_3d_sheet_span_with_quoted_sheet_names_expands_via_provider_sheet_order() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            "Sheet 1".to_string(),
            "Sheet 2".to_string(),
            "Sheet 3".to_string(),
        ],
    );
    for (sheet, value) in [("Sheet 1", 1.0), ("Sheet 2", 2.0), ("Sheet 3", 3.0)] {
        provider.set(
            &format!("[Book.xlsx]{sheet}"),
            CellAddr { row: 0, col: 0 },
            value,
        );
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]'Sheet 1':'Sheet 3'!A1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(6.0));
}

#[test]
fn external_3d_sheet_span_matches_endpoints_case_insensitively() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set_sheet_order(
        "Book.xlsx",
        vec![
            "Sheet1".to_string(),
            "Sheet2".to_string(),
            "Sheet3".to_string(),
        ],
    );
    for (sheet, value) in [("Sheet1", 1.0), ("Sheet2", 2.0), ("Sheet3", 3.0)] {
        provider.set(
            &format!("[Book.xlsx]{sheet}"),
            CellAddr { row: 0, col: 0 },
            value,
        );
    }

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]sheet1:sheet3!A1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(6.0));
}

#[test]
fn external_3d_sheet_span_with_missing_endpoint_is_ref_error() {
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
        .set_cell_formula("Sheet1", "A1", "=SUM([Book.xlsx]Sheet1:Sheet4!A1)")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );
}

#[test]
fn external_3d_sheet_range_refs_are_ref_error_even_if_provider_has_value() {
    let provider = Arc::new(TestExternalProvider::default());
    provider.set(
        "[Book.xlsx]Sheet1:Sheet3",
        CellAddr { row: 0, col: 0 },
        41.0,
    );

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));
    engine
        .set_cell_formula("Sheet1", "A1", "=[Book.xlsx]Sheet1:Sheet3!A1")
        .unwrap();
    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Ref)
    );
}

#[test]
fn database_functions_support_computed_criteria_over_external_database() {
    let provider = Arc::new(TestExternalProvider::default());

    // External database: [Book.xlsx]Sheet1!A1:D4 (header + 3 records).
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 0 }, "Name");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 1 }, "Dept");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 2 }, "Age");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 0, col: 3 }, "Salary");

    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 0 }, "Alice");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 1 }, "Sales");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 2 }, 30.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 1, col: 3 }, 1000.0);

    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 0 }, "Bob");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 1 }, "Sales");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 2 }, 35.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 2, col: 3 }, 1500.0);

    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 0 }, "Carol");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 1 }, "HR");
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 2 }, 28.0);
    provider.set("[Book.xlsx]Sheet1", CellAddr { row: 3, col: 3 }, 1200.0);

    let mut engine = Engine::new();
    engine.set_external_value_provider(Some(provider));

    // Computed criteria: blank header + formula referencing the first database record row.
    engine
        .set_cell_formula("Sheet1", "F2", "=C2>30")
        .unwrap();

    // Evaluate a representative set of D* functions over the external database range.
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=DSUM([Book.xlsx]Sheet1!A1:D4,\"Salary\",F1:F2)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=DGET([Book.xlsx]Sheet1!A1:D4,\"Salary\",F1:F2)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            "=DCOUNT([Book.xlsx]Sheet1!A1:D4,\"Salary\",F1:F2)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            "=DVARP([Book.xlsx]Sheet1!A1:D4,\"Salary\",F1:F2)",
        )
        .unwrap();
    engine.recalculate();

    // Age > 30 matches only Bob => sum salary = 1500.
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1500.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(1500.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Number(0.0));
}
