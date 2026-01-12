use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn bytecode_backend_spills_range_reference() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();

    // Ensure we're exercising the bytecode backend.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));
}

#[test]
fn bytecode_backend_spills_range_reference_with_mixed_types() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", "a").unwrap();
    engine.set_cell_value("Sheet1", "A2", true).unwrap();
    engine
        .set_cell_value("Sheet1", "A3", Value::Error(ErrorKind::Div0))
        .unwrap();

    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C3").unwrap());

    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Text("a".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Bool(true));
    assert_eq!(
        engine.get_cell_value("Sheet1", "C3"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn bytecode_backend_spills_range_plus_scalar() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine.set_cell_formula("Sheet1", "C1", "=A1:A3+1").unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(4.0));
}

#[test]
fn bytecode_backend_broadcasts_row_and_column_ranges() {
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
    assert_eq!(engine.bytecode_program_count(), 1);

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
fn bytecode_backend_spills_comparison_results() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", -2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 0.0).unwrap();

    engine.set_cell_formula("Sheet1", "B1", "=A1:A3>0").unwrap();
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "B1").expect("spill range");
    assert_eq!(start, parse_a1("B1").unwrap());
    assert_eq!(end, parse_a1("B3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Bool(true));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::Bool(false));
}

#[test]
fn bytecode_backend_spills_array_literal() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "={1,2;3,4}")
        .unwrap();

    // Ensure we're exercising the bytecode backend.
    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("D2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(4.0));
}

#[test]
fn bytecode_backend_spills_array_literal_plus_scalar() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "={1,2;3,4}+1")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("D2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(5.0));
}

#[test]
fn bytecode_backend_spills_let_bound_array_literal() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "=LET(a,{1,2;3,4},a+1)")
        .unwrap();

    assert_eq!(engine.bytecode_program_count(), 1);

    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("D2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D2"), Value::Number(5.0));
}
