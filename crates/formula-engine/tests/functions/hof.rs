use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn map_applies_lambda_over_range() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=MAP(A1:A3,LAMBDA(x,x*2))")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(6.0));
}

#[test]
fn map_over_two_ranges_sums_elements() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 30.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=MAP(A1:A3,B1:B3,LAMBDA(x,y,x+y))")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(11.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(22.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(33.0));
}

#[test]
fn map_shape_mismatch_returns_value_error() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "E1", "=MAP(A1:A3,B1:B2,LAMBDA(x,y,x+y))")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Error(ErrorKind::Value));
}

#[test]
fn reduce_sums_over_range_with_and_without_initial() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C5", "=REDUCE(A1:A3,LAMBDA(acc,v,acc+v))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D5", "=REDUCE(0,A1:A3,LAMBDA(acc,v,acc+v))")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "C5"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D5"), Value::Number(6.0));
}

#[test]
fn scan_spills_running_accumulations_over_range() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "F1", "=SCAN(0,A1:A3,LAMBDA(acc,v,acc+v))")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "F1").expect("spill range");
    assert_eq!(start, parse_a1("F1").unwrap());
    assert_eq!(end, parse_a1("F3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F3"), Value::Number(6.0));
}

#[test]
fn reduce_can_accumulate_dynamic_arrays() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();

    // REDUCE should allow the lambda to return a (spillable) dynamic array and continue reducing
    // over it. This enables common patterns like building a result via VSTACK/HSTACK.
    engine
        .set_cell_formula(
            "Sheet1",
            "C1",
            "=REDUCE(1,A1:A2,LAMBDA(acc,v,VSTACK(acc,v)))",
        )
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));
}
