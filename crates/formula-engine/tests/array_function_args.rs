use formula_engine::{Engine, Value};

#[test]
fn functions_accept_array_results_from_operator_expressions() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", -2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 30.0).unwrap();

    // MIN should accept an array produced by range arithmetic.
    engine
        .set_cell_formula("Sheet1", "D1", "=MIN(A1:A3*10)")
        .unwrap();

    // AND should accept an array produced by a comparison.
    engine
        .set_cell_formula("Sheet1", "D2", "=AND(A1:A3>0)")
        .unwrap();

    // SUMPRODUCT should accept array + range inputs, enabling common filtering patterns.
    engine
        .set_cell_formula("Sheet1", "D3", "=SUMPRODUCT(A1:A3>0,B1:B3)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(-20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(40.0));
}

#[test]
fn concat_flattens_array_results_from_operators() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=CONCAT(A1:A3*10)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Text("102030".to_string())
    );
}
