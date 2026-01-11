use formula_engine::eval::CellAddr;
use formula_engine::{Engine, ExternalValueProvider, PrecedentNode, Value};
use std::collections::HashMap;
use std::sync::{Arc, Mutex};

#[derive(Default)]
struct TestExternalProvider {
    values: Mutex<HashMap<(String, CellAddr), Value>>,
}

impl TestExternalProvider {
    fn set(&self, sheet: &str, addr: CellAddr, value: impl Into<Value>) {
        self.values
            .lock()
            .expect("lock poisoned")
            .insert((sheet.to_string(), addr), value.into());
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
