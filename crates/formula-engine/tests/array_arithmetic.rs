use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn range_times_scalar_spills_elementwise() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=A1:A3*10")
        .unwrap();
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
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(2,2)+SEQUENCE(3,2)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Value)
    );
    assert!(engine.spill_range("Sheet1", "A1").is_none());
}

#[test]
fn outer_broadcasting_spills_2d_arrays() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(2,1)+SEQUENCE(1,3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("C2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(5.0));
}

#[test]
fn outer_broadcasting_spills_2d_arrays_for_multiplication() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(2,1)*SEQUENCE(1,3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("C2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(6.0));
}

#[test]
fn outer_broadcasting_spills_2d_arrays_for_exponentiation() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(2,1)^SEQUENCE(1,3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("C2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(8.0));
}

#[test]
fn outer_broadcasting_spills_2d_arrays_for_division() {
    // Use a 2x1 / 1x2 example so all results are exactly representable in IEEE754.
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(2,1)/SEQUENCE(1,2)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(0.5));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(1.0));
}

#[test]
fn outer_broadcasting_respects_operand_positions_for_subtraction() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(2,1)-SEQUENCE(1,3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("C2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(-1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(-2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(0.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(-1.0));
}

#[test]
fn outer_broadcasting_over_ranges_spills_2d_arrays() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine.set_cell_value("Sheet1", "B1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "D1", 30.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "F1", "=A1:A3+B1:D1")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "F1").expect("spill range");
    assert_eq!(start, parse_a1("F1").unwrap());
    assert_eq!(end, parse_a1("H3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(11.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G1"), Value::Number(21.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H1"), Value::Number(31.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F2"), Value::Number(12.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G2"), Value::Number(22.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H2"), Value::Number(32.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F3"), Value::Number(13.0));
    assert_eq!(engine.get_cell_value("Sheet1", "G3"), Value::Number(23.0));
    assert_eq!(engine.get_cell_value("Sheet1", "H3"), Value::Number(33.0));
}

#[test]
fn outer_broadcasting_over_array_literals_spills_2d_arrays() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "={1;2}+{10,20,30}")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("C2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(11.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(21.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(31.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(12.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(22.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(32.0));
}

#[test]
fn outer_broadcasting_over_spilled_ranges_spills_2d_arrays() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(3,1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=SEQUENCE(1,4)")
        .unwrap();
    engine.set_cell_formula("Sheet1", "A5", "=A1#+C1#").unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A5").expect("spill range");
    assert_eq!(start, parse_a1("A5").unwrap());
    assert_eq!(end, parse_a1("D7").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A5"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B5"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C5"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D5"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A6"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B6"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C6"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D6"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A7"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B7"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C7"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D7"), Value::Number(7.0));
}

#[test]
fn row_broadcasting_preserves_column_count() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(3,2)+SEQUENCE(1,2)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(8.0));
}

#[test]
fn column_broadcasting_preserves_row_count() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(3,2)+SEQUENCE(3,1)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(6.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::Number(8.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Number(9.0));
}

#[test]
fn comparisons_spill_boolean_arrays() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", -2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 0.0).unwrap();

    engine.set_cell_formula("Sheet1", "B1", "=A1:A3>0").unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "B1").expect("spill range");
    assert_eq!(start, parse_a1("B1").unwrap());
    assert_eq!(end, parse_a1("B3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Bool(false));
}

#[test]
fn comparison_broadcasting_spills_2d_boolean_arrays() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(2,1)=SEQUENCE(1,3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("C2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Bool(false));
}

#[test]
fn comparison_broadcasting_respects_operator_direction() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(2,1)<SEQUENCE(1,3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("C2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Bool(true));
}

#[test]
fn unary_minus_spills_over_ranges() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", -3.0).unwrap();

    engine.set_cell_formula("Sheet1", "C1", "=-A1:A3").unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(-1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(-2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));
}

#[test]
fn concat_broadcasts_scalars_over_arrays() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "a").unwrap();
    engine.set_cell_value("Sheet1", "A2", "b").unwrap();
    engine.set_cell_value("Sheet1", "A3", "").unwrap();

    engine
        .set_cell_formula("Sheet1", "D1", "=A1:A3&\"x\"")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "D1").expect("spill range");
    assert_eq!(start, parse_a1("D1").unwrap());
    assert_eq!(end, parse_a1("D3").unwrap());

    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Text("ax".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "D2"),
        Value::Text("bx".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "D3"),
        Value::Text("x".to_string())
    );
}

#[test]
fn concat_broadcasts_outer_arrays() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(2,1)&SEQUENCE(1,3)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("C2").unwrap());

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("11".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Text("12".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Text("13".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Text("21".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Text("22".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "C2"),
        Value::Text("23".to_string())
    );
}

#[test]
fn spill_range_operator_participates_in_elementwise_ops() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SEQUENCE(3)")
        .unwrap();
    engine.set_cell_formula("Sheet1", "E1", "=A1#*10").unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "E1").expect("spill range");
    assert_eq!(start, parse_a1("E1").unwrap());
    assert_eq!(end, parse_a1("E3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(10.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(20.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(30.0));
}
