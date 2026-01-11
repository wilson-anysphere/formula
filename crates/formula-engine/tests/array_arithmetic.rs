use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn range_times_scalar_spills_elementwise() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1:A3*10").unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(30.0));
}

#[test]
fn range_plus_range_spills_elementwise_sums() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 30.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=(A1:A3)+(B1:B3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "D1").expect("spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("D3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(11.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(22.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(33.0));
}

#[test]
fn mismatched_array_shapes_return_value_error() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(1,2)+SEQUENCE(2,1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Error(ErrorKind::Value));
    assert!(engine.spill_range("Sheet1", "A1").is_none());
}

