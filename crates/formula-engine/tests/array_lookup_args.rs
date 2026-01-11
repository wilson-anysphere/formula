use formula_engine::{Engine, Value};

#[test]
fn lookup_functions_accept_array_arguments_from_operator_expressions() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=MATCH(20, A1:A3*10, 0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=INDEX(SEQUENCE(2,2)*10, 2, 1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C3", "=XMATCH(2, SEQUENCE(3))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C4", "=XLOOKUP(2, SEQUENCE(3), SEQUENCE(3)*10)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C5", "=VLOOKUP(3, SEQUENCE(3,2,1,1), 2, FALSE)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C6", "=HLOOKUP(2, SEQUENCE(2,3,1,1), 2, FALSE)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(30.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C4"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C5"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C6"), Value::Number(5.0));
}
