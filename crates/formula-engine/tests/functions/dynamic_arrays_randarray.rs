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
fn randarray_spill_blocking_produces_spill_error_and_resolves_after_clear() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(1,2)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B1").unwrap());

    // Block the spill output cell with a user value.
    engine.set_cell_value("Sheet1", "B1", 99.0).unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Spill)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(99.0));
    assert!(engine.spill_range("Sheet1", "A1").is_none());

    // Clearing the blocker should allow the volatile formula to spill again.
    engine.set_cell_value("Sheet1", "B1", Value::Blank).unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B1").unwrap());

    for addr in ["A1", "B1"] {
        let Value::Number(n) = engine.get_cell_value("Sheet1", addr) else {
            panic!("expected {addr} to be a number after spill resolves");
        };
        assert!((0.0..1.0).contains(&n), "expected 0 <= {n} < 1");
    }
}

#[test]
fn randarray_spill_too_big_returns_spill_error() {
    let mut engine = Engine::new();

    // Last column in Excel is XFD. Spilling two columns from there would exceed the sheet bounds.
    engine
        .set_cell_formula("Sheet1", "XFD1", "=RANDARRAY(1,2)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "XFD1"),
        Value::Error(ErrorKind::Spill)
    );
    assert!(engine.spill_range("Sheet1", "XFD1").is_none());

    // Similarly, spilling down from the last row should fail.
    engine
        .set_cell_formula("Sheet1", "A1048576", "=RANDARRAY(2,1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1048576"),
        Value::Error(ErrorKind::Spill)
    );
    assert!(engine.spill_range("Sheet1", "A1048576").is_none());
}

#[test]
fn randarray_missing_rows_uses_default() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(,3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("C1").unwrap());
}

#[test]
fn randarray_missing_cols_uses_default() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(2,)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("A2").unwrap());
}

#[test]
fn randarray_accepts_xlfn_prefix() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=_xlfn.RANDARRAY(,3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("C1").unwrap());
}

#[test]
fn randarray_handles_large_min_max_without_overflowing_span() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(1,1,-1E308,1E308)")
        .unwrap();
    engine.recalculate_single_threaded();

    let Value::Number(n) = engine.get_cell_value("Sheet1", "A1") else {
        panic!("expected RANDARRAY to return a number");
    };
    assert!(n.is_finite(), "expected result to be finite, got {n}");
    assert!(
        (-1e308..1e308).contains(&n),
        "expected result to be within [-1e308, 1e308], got {n}"
    );
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

#[test]
fn randarray_min_equals_max_returns_constant_array() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(2,2,5,5)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B2").unwrap());

    for addr in ["A1", "B1", "A2", "B2"] {
        assert_eq!(
            engine.get_cell_value("Sheet1", addr),
            Value::Number(5.0),
            "expected {addr} to be the constant min/max value"
        );
    }
}

#[test]
fn randarray_blank_min_uses_default() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(1,1,,2)")
        .unwrap();
    engine.recalculate_single_threaded();

    let Value::Number(n) = engine.get_cell_value("Sheet1", "A1") else {
        panic!("expected RANDARRAY to return a number");
    };
    assert!((0.0..2.0).contains(&n), "expected 0 <= {n} < 2");
}

#[test]
fn randarray_missing_max_uses_default() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(1,1,1)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
}

#[test]
fn randarray_whole_number_requires_non_empty_integer_interval() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(1,1,0.2,0.8,TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn randarray_rejects_non_finite_bounds() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=RANDARRAY(1,1,0,1E308*1E308)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Num)
    );
}
