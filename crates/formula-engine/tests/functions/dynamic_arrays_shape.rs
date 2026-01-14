use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn take_selects_rows_and_cols_and_handles_missing_args() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 4.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 6.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 7.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 8.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 9.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "E1", "=TAKE(A1:C3,1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "E2", "=TAKE(A1:C3,-1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "E4", "=TAKE(A1:C3,,2)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "E1").expect("spill range");
    assert_eq!(start, parse_a1("E1").unwrap());
    assert_eq!(end, parse_a1("G1").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(3.0));

    let (start, end) = engine.spill_range("Sheet1", "E2").expect("spill range");
    assert_eq!(start, parse_a1("E2").unwrap());
    assert_eq!(end, parse_a1("G2").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(7.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(8.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(9.0));

    let (start, end) = engine.spill_range("Sheet1", "E4").expect("spill range");
    assert_eq!(start, parse_a1("E4").unwrap());
    assert_eq!(end, parse_a1("F6").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "E4"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F4"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E5"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F5"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E6"), Value::Number(7.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F6"), Value::Number(8.0));
}

#[test]
fn take_accepts_spilled_input() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(3,2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=TAKE(A1#,2,1)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "D1").expect("spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("D2").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(3.0));
}

#[test]
fn drop_drops_rows_and_returns_calc_when_empty() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=DROP(A1:A3,1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=DROP(A1:A3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "E1", "=DROP(A1:A3,3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C2").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(3.0));

    let (start, end) = engine.spill_range("Sheet1", "D1").expect("spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("D3").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(3.0));

    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        Value::Error(ErrorKind::Calc)
    );
}

#[test]
fn choosecols_and_chooserows_support_negative_and_duplicate_indices() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 4.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 6.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 7.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 8.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 9.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "E1", "=CHOOSECOLS(A1:C3,1,3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "G1", "=CHOOSECOLS(A1:C3,-1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "H1", "=CHOOSECOLS(A1:C3,2,2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "J1", "=CHOOSECOLS(A1:C3,0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "K1", "=CHOOSECOLS(A1:C3,4)")
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "E5", "=CHOOSEROWS(A1:C3,1,3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "E7", "=CHOOSEROWS(A1:C3,-1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "E8", "=CHOOSEROWS(A1:C3,2,2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "E10", "=CHOOSEROWS(A1:C3,0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "F10", "=CHOOSEROWS(A1:C3,4)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(7.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F3"), Value::Number(9.0));

    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G3"), Value::Number(9.0));

    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H2"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I2"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H3"), Value::Number(8.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I3"), Value::Number(8.0));

    assert_eq!(
        engine.get_cell_value("Sheet1", "J1"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "K1"),
        Value::Error(ErrorKind::Value)
    );

    assert_eq!(engine.get_cell_value("Sheet1", "E5"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F5"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G5"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E6"), Value::Number(7.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F6"), Value::Number(8.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G6"), Value::Number(9.0));

    assert_eq!(engine.get_cell_value("Sheet1", "E7"), Value::Number(7.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F7"), Value::Number(8.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G7"), Value::Number(9.0));

    assert_eq!(engine.get_cell_value("Sheet1", "E8"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F8"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G8"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E9"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F9"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G9"), Value::Number(6.0));

    assert_eq!(
        engine.get_cell_value("Sheet1", "E10"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "F10"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn expand_expands_with_default_and_custom_padding() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 4.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=EXPAND(A1:B2,3,4)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D5", "=EXPAND(A1:B2,3,4,0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D9", "=EXPAND(A1:B2,3,4,NA())")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D13", "=EXPAND(A1:B2,1,2)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(2.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "F1"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "G1"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(4.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "F2"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "G2"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "D3"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "G3"),
        Value::Error(ErrorKind::NA)
    );

    assert_eq!(engine.get_cell_value("Sheet1", "D5"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E5"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F5"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G5"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D6"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E6"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F6"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G6"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D7"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G7"), Value::Number(0.0));

    assert_eq!(engine.get_cell_value("Sheet1", "D9"), Value::Number(1.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "F9"),
        Value::Error(ErrorKind::NA)
    );

    assert_eq!(
        engine.get_cell_value("Sheet1", "D13"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn take_and_drop_return_calc_when_result_is_empty() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 4.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=TAKE(A1:B2,0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D2", "=TAKE(A1:B2,1,0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "E1", "=DROP(A1:B2,,2)")
        .unwrap();

    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Error(ErrorKind::Calc)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "D2"),
        Value::Error(ErrorKind::Calc)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        Value::Error(ErrorKind::Calc)
    );
}

#[test]
fn shape_functions_accept_xlfn_prefix() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 4.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 6.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 7.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 8.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 9.0).unwrap();

    // TAKE/DROP are stored with `_xlfn.` in file formats for forward compatibility, but the engine
    // should accept the prefixed names for evaluation.
    engine
        .set_cell_formula("Sheet1", "E1", "=_xlfn.TAKE(A1:C3,1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=_xlfn.DROP(A1:A3)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "H1", "=_xlfn.CHOOSECOLS(A1:C3,2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "H5", "=_xlfn.CHOOSEROWS(A1:C3,2)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "H7", "=_xlfn.EXPAND(A1:B2,3,2,0)")
        .unwrap();

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "E1").expect("spill range");
    assert_eq!(start, parse_a1("E1").unwrap());
    assert_eq!(end, parse_a1("G1").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(3.0));

    let (start, end) = engine.spill_range("Sheet1", "D1").expect("spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("D3").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(7.0));

    let (start, end) = engine.spill_range("Sheet1", "H1").expect("spill range");
    assert_eq!(start, parse_a1("H1").unwrap());
    assert_eq!(end, parse_a1("H3").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H2"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H3"), Value::Number(8.0));

    let (start, end) = engine.spill_range("Sheet1", "H5").expect("spill range");
    assert_eq!(start, parse_a1("H5").unwrap());
    assert_eq!(end, parse_a1("J5").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "H5"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I5"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J5"), Value::Number(6.0));

    let (start, end) = engine.spill_range("Sheet1", "H7").expect("spill range");
    assert_eq!(start, parse_a1("H7").unwrap());
    assert_eq!(end, parse_a1("I9").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "H7"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I7"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H8"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I8"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H9"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I9"), Value::Number(0.0));
}
