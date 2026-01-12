use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn randarray_default_returns_number_between_0_and_1() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY()")
        .unwrap();
    engine.recalculate_single_threaded();

    let Value::Number(n) = engine.get_cell_value("Sheet1", "A1") else {
        panic!("expected RANDARRAY() to return a number");
    };
    assert!((0.0..1.0).contains(&n), "expected 0 <= {n} < 1");
}

#[test]
fn randarray_spills_2_by_3() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(2,3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("C2").unwrap());

    for row in 1..=2 {
        for col in ['A', 'B', 'C'] {
            let addr = format!("{col}{row}");
            let Value::Number(n) = engine.get_cell_value("Sheet1", &addr) else {
                panic!("expected {addr} to be a number");
            };
            assert!((0.0..1.0).contains(&n), "expected 0 <= {n} < 1");
        }
    }
}

#[test]
fn randarray_whole_numbers_respect_min_max() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(3,1,10,20,TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("A3").unwrap());

    for row in 1..=3 {
        let addr = format!("A{row}");
        let Value::Number(n) = engine.get_cell_value("Sheet1", &addr) else {
            panic!("expected {addr} to be a number");
        };
        assert!((10.0..=20.0).contains(&n), "expected 10 <= {n} <= 20");
        assert_eq!(n.fract(), 0.0, "expected {n} to be integer-valued");
    }
}

#[test]
fn randarray_rejects_non_positive_dimensions() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(0,1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn randarray_rejects_min_greater_than_max() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(1,1,5,4)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Value)
    );
}
