use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn filter_filters_rows_from_range() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", true).unwrap();
    engine.set_cell_value("Sheet1", "B2", false).unwrap();
    engine.set_cell_value("Sheet1", "B3", true).unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=FILTER(A1:A3,B1:B3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "D1").expect("spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("D2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(3.0));
}

#[test]
fn filter_filters_columns_from_range() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 4.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 5.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 6.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "E1", "=FILTER(A1:C2,{1,0,1})")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "E1").expect("spill range");
    assert_eq!(start, parse_a1("E1").unwrap());
    assert_eq!(end, parse_a1("F2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(6.0));
}

#[test]
fn filter_if_empty_branch_and_calc_error() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", false).unwrap();
    engine.set_cell_value("Sheet1", "B2", false).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=FILTER(A1:A2,B1:B2,\"none\")")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Text("none".to_string())
    );

    engine
        .set_cell_formula("Sheet1", "E1", "=FILTER(A1:A2,B1:B2)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "E1"),
        Value::Error(ErrorKind::Calc)
    );
}

#[test]
fn sort_sorts_numbers_and_text() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=SORT(A1:A3)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));

    engine.set_cell_value("Sheet1", "B1", "b").unwrap();
    engine.set_cell_value("Sheet1", "B2", "A").unwrap();
    engine.set_cell_value("Sheet1", "B3", "c").unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=SORT(B1:B3)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Text("A".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "D2"),
        Value::Text("b".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "D3"),
        Value::Text("c".to_string())
    );
}

#[test]
fn sort_by_col_sorts_columns() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 20.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "E1", "=SORT(A1:C2,1,1,TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(30.0));
}

#[test]
fn unique_by_row_and_column_and_exactly_once() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A4", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A5", 2.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=UNIQUE(A1:A5)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));

    engine
        .set_cell_formula("Sheet1", "D1", "=UNIQUE(A1:A5,,TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(3.0));

    engine.set_cell_value("Sheet1", "E1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "F1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "G1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "E2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "F2", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "G2", 4.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "I1", "=UNIQUE(E1:G2,TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "I1").expect("spill range");
    assert_eq!(start, parse_a1("I1").unwrap());
    assert_eq!(end, parse_a1("J2").unwrap());

    // First two columns are duplicates; UNIQUE by_col keeps the first occurrence.
    assert_eq!(engine.get_cell_value("Sheet1", "I1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J2"), Value::Number(4.0));
}

#[test]
fn spill_blocked_spill_and_footprint_change_recalculate_dependents() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", true).unwrap();
    engine.set_cell_value("Sheet1", "B2", true).unwrap();
    engine.set_cell_value("Sheet1", "B3", true).unwrap();
    engine
        .set_cell_formula("Sheet1", "D1", "=FILTER(A1:A3,B1:B3)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(3.0));

    // Block the spill.
    engine.set_cell_value("Sheet1", "D2", 99.0).unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Error(ErrorKind::Spill)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(99.0));

    // Clear the blocker and shrink/expand the spill while ensuring dependents recalculate.
    engine.set_cell_value("Sheet1", "D2", Value::Blank).unwrap();
    engine.set_cell_value("Sheet1", "B2", false).unwrap();
    engine.set_cell_value("Sheet1", "B3", false).unwrap();
    engine.set_cell_formula("Sheet1", "E1", "=D3").unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Blank);

    // Expanding the footprint should update E1 via the new spill cell dependency.
    engine.set_cell_value("Sheet1", "B2", true).unwrap();
    engine.set_cell_value("Sheet1", "B3", true).unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(3.0));
}

#[test]
fn sort_accepts_arrays_produced_by_expressions() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", true).unwrap();
    engine.set_cell_value("Sheet1", "B2", false).unwrap();
    engine.set_cell_value("Sheet1", "B3", true).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=SORT(FILTER(A1:A3,B1:B3))")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(3.0));
}

#[test]
fn sort_rejects_invalid_order_and_index() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=SORT(A1:A2,0)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Error(ErrorKind::Value)
    );

    engine
        .set_cell_formula("Sheet1", "D1", "=SORT(A1:A2,1,0)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Error(ErrorKind::Value)
    );
}
