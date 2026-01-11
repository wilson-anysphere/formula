use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn engine_evaluates_r1c1_relative_cell_reference_equivalent_to_a1() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 42.0).unwrap();

    engine.set_cell_formula("Sheet1", "C5", "=A1").unwrap();
    engine.recalculate();
    let a1_value = engine.get_cell_value("Sheet1", "C5");

    engine
        .set_cell_formula_r1c1("Sheet1", "C5", "=R[-4]C[-2]")
        .unwrap();
    engine.recalculate();
    let r1c1_value = engine.get_cell_value("Sheet1", "C5");

    assert_eq!(a1_value, r1c1_value);
    assert_eq!(r1c1_value, Value::Number(42.0));
}

#[test]
fn engine_evaluates_r1c1_ranges_equivalent_to_a1() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "A3", 3.0).unwrap();

    engine.set_cell_formula("Sheet1", "C5", "=A1:A3").unwrap();
    engine.recalculate();
    let a1_value = engine.get_cell_value("Sheet1", "C5");

    engine
        .set_cell_formula_r1c1("Sheet1", "C5", "=R1C1:R3C1")
        .unwrap();
    engine.recalculate();
    let r1c1_value = engine.get_cell_value("Sheet1", "C5");

    assert_eq!(a1_value, r1c1_value);
    assert_eq!(r1c1_value, Value::Error(ErrorKind::Spill));
}

