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

#[test]
fn sort_treats_blank_optional_args_as_defaults() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 2.0).unwrap();

    // Leave sort_order blank to reach by_col argument; Excel treats this as omitted.
    engine
        .set_cell_formula("Sheet1", "C1", "=SORT(A1:A3,1,,FALSE)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));
}

#[test]
fn sortby_sorts_rows_and_columns() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 30.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 20.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=SORTBY(A1:A3,B1:B3)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D3"), Value::Number(1.0));

    // Column sort: sort the columns of A1:C1 based on a 1x3 key vector.
    engine.set_cell_value("Sheet1", "F1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "G1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "H1", 3.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "F2", "=SORTBY(F1:H1,{3,1,2})")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H2"), Value::Number(1.0));
}

#[test]
fn sortby_disambiguates_optional_sort_order_args() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 100.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 200.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 300.0).unwrap();

    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B3", 2.0).unwrap();

    engine.set_cell_value("Sheet1", "C1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 0.0).unwrap();

    // Excel treats the 3rd argument (C1:C3) as the second by_array (not sort_order1) because it is a range.
    engine
        .set_cell_formula("Sheet1", "E1", "=SORTBY(A1:A3,B1:B3,C1:C3)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(200.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(100.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(300.0));
}

#[test]
fn take_and_drop_slice_arrays() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "={1,2,3;4,5,6;7,8,9}")
        .unwrap();
    engine.recalculate_single_threaded();

    engine
        .set_cell_formula("Sheet1", "E1", "=TAKE(A1:C3,2,-2)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(6.0));

    engine
        .set_cell_formula("Sheet1", "G1", "=TAKE(A1:C3,-1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(7.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(8.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I1"), Value::Number(9.0));

    engine
        .set_cell_formula("Sheet1", "E4", "=DROP(A1:C3,1,1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E4"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F4"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E5"), Value::Number(8.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F5"), Value::Number(9.0));

    engine
        .set_cell_formula("Sheet1", "G4", "=DROP(A1:C3,-1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "G4"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H4"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I4"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G5"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H5"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I5"), Value::Number(6.0));
}

#[test]
fn choosecols_and_chooserows_support_negative_indices() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "={1,2,3;4,5,6}")
        .unwrap();
    engine.recalculate_single_threaded();

    engine
        .set_cell_formula("Sheet1", "E1", "=CHOOSECOLS(A1:C2,3,1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(4.0));

    engine
        .set_cell_formula("Sheet1", "G1", "=CHOOSECOLS(A1:C2,-1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(6.0));

    engine
        .set_cell_formula("Sheet1", "I1", "=CHOOSEROWS(A1:C2,-1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "I1"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J1"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "K1"), Value::Number(6.0));

    engine
        .set_cell_formula("Sheet1", "L1", "=CHOOSECOLS(A1:C2,0)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "L1"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn hstack_and_vstack_fill_missing_with_na() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=HSTACK({1;2},{3;4;5})")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(4.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "A3"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(5.0));

    engine
        .set_cell_formula("Sheet1", "D1", "=VSTACK({1,2},{3})")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(3.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "E2"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn tocol_and_torow_flatten_arrays_with_ignore_modes() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TOCOL({1,2;3,4})")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Number(4.0));

    engine
        .set_cell_formula("Sheet1", "C1", "=TOCOL({1,,2},1)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));

    engine
        .set_cell_formula("Sheet1", "E1", "=TOCOL({1,#VALUE!,2},2)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(2.0));

    engine
        .set_cell_formula("Sheet1", "G1", "=TOROW({1,2;3,4},0,TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();
    // Scan by column when the third argument is TRUE.
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "I1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "J1"), Value::Number(4.0));
}

#[test]
fn wraprows_and_wrapcols_wrap_vectors_with_padding() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=WRAPROWS({1;2;3;4;5},2)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(5.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "B3"),
        Value::Error(ErrorKind::NA)
    );

    engine
        .set_cell_formula("Sheet1", "D1", "=WRAPCOLS({1;2;3;4;5},2)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(4.0));
    assert_eq!(
        engine.get_cell_value("Sheet1", "F2"),
        Value::Error(ErrorKind::NA)
    );

    engine
        .set_cell_formula("Sheet1", "H1", "=WRAPROWS({1;2;3},2,0)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "I2"), Value::Number(0.0));
}

#[test]
fn multithreaded_and_singlethreaded_recalc_match_for_dynamic_arrays() {
    fn setup(engine: &mut Engine) {
        // Data for SORT stability: sort by column B ascending, with ties.
        engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
        engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
        engine.set_cell_value("Sheet1", "A4", 4.0).unwrap();

        engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
        engine.set_cell_value("Sheet1", "B2", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "B3", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "B4", 3.0).unwrap();

        engine
            .set_cell_formula("Sheet1", "D1", "=SORT(A1:B4,2,1,FALSE)")
            .unwrap();

        // Data for UNIQUE.
        engine.set_cell_value("Sheet1", "F1", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "F2", 1.0).unwrap();
        engine.set_cell_value("Sheet1", "F3", 2.0).unwrap();
        engine.set_cell_value("Sheet1", "F4", 3.0).unwrap();
        engine.set_cell_value("Sheet1", "F5", 2.0).unwrap();
        engine
            .set_cell_formula("Sheet1", "H1", "=UNIQUE(F1:F5)")
            .unwrap();
    }

    let mut single = Engine::new();
    setup(&mut single);
    single.recalculate_single_threaded();

    let mut multi = Engine::new();
    setup(&mut multi);
    multi.recalculate_multi_threaded();

    // SORT results should match across modes.
    for addr in ["D1", "E1", "D2", "E2", "D3", "E3", "D4", "E4"] {
        assert_eq!(
            multi.get_cell_value("Sheet1", addr),
            single.get_cell_value("Sheet1", addr),
            "mismatch for {addr}"
        );
    }

    // UNIQUE results should match across modes.
    for addr in ["H1", "H2", "H3"] {
        assert_eq!(
            multi.get_cell_value("Sheet1", addr),
            single.get_cell_value("Sheet1", addr),
            "mismatch for {addr}"
        );
    }
}
