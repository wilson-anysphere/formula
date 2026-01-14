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

    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        Value::Error(ErrorKind::Value)
    );
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
fn reduce_without_initial_uses_first_array_element() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();

    // If initial_value is omitted, Excel uses the first element of the array as the initial
    // accumulator (rather than starting from blank/zero).
    engine
        .set_cell_formula("Sheet1", "C7", "=REDUCE(A1:A2,LAMBDA(acc,v,acc*v))")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "C7"), Value::Number(6.0));
}

#[test]
fn reduce_can_recover_from_error_accumulator() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=1/0")
        .expect("set formula");
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    // When initial_value is omitted, REDUCE seeds the accumulator from the first array element.
    // If that element is an error, Excel still evaluates the lambda for subsequent elements, so
    // the lambda can choose to recover.
    engine
        .set_cell_formula(
            "Sheet1",
            "C9",
            "=REDUCE(A1:A3,LAMBDA(acc,v,IFERROR(acc,0)+v))",
        )
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "C9"), Value::Number(5.0));
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
fn scan_without_initial_uses_first_array_element() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 4.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "G1", "=SCAN(A1:A3,LAMBDA(acc,v,acc*v))")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "G1").expect("spill range");
    assert_eq!(start, parse_a1("G1").unwrap());
    assert_eq!(end, parse_a1("G3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G3"), Value::Number(24.0));
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
