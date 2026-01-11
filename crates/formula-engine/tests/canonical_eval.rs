use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn percent_postfix_evaluates() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=1%").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.01));
}

#[test]
fn concat_operator_evaluates() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"="a"&"b""#)
        .unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("ab".into())
    );
}

#[test]
fn sum_range_regression() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=SUM(A1:A3)")
        .unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(6.0));
}

#[test]
fn formulas_with_union_intersect_and_array_literals_ingest() {
    let mut engine = Engine::new();

    // Union (comma outside a function call).
    engine.set_cell_formula("Sheet1", "A1", "=B1,C1").unwrap();
    // Intersection (whitespace).
    engine.set_cell_formula("Sheet1", "A2", "=B1 C1").unwrap();
    // Array literal.
    engine
        .set_cell_formula("Sheet1", "A3", "={1,2;3,4}")
        .unwrap();

    engine.recalculate();

    // Semantics for these constructs may be partial; evaluation should be deterministic and
    // non-panicking.
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::Null)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B4"), Value::Number(4.0));
}
