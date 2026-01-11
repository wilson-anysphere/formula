use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn reference_spill_spills_values() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "C1").expect("spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("C3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Number(3.0));

    let (origin_sheet, origin_addr) = engine.spill_origin("Sheet1", "C2").expect("spill origin");
    assert_eq!(origin_sheet, 0);
    assert_eq!(origin_addr, parse_a1("C1").unwrap());
}

#[test]
fn spill_blocking_produces_spill_error() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();
    engine.recalculate_single_threaded();

    // Block the middle spill cell with a user value.
    engine.set_cell_value("Sheet1", "C2", 99.0).unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Error(ErrorKind::Spill));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(99.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C3"), Value::Blank);
    assert!(engine.spill_range("Sheet1", "C1").is_none());
}

#[test]
fn transpose_spills_down() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 3.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "E1", "=TRANSPOSE(A1:C1)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "E1").expect("spill range");
    assert_eq!(start, parse_a1("E1").unwrap());
    assert_eq!(end, parse_a1("E3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E2"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "E3"), Value::Number(3.0));
}

#[test]
fn sequence_spills_matrix() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "C1", "=SEQUENCE(2,2,1,1)")
        .unwrap();
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
fn dependents_of_spill_cells_recalculate() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1:A3").unwrap();
    engine.set_cell_formula("Sheet1", "D1", "=C2*10").unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(20.0));

    engine.set_cell_value("Sheet1", "A2", 5.0).unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(5.0));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Number(50.0));
}

