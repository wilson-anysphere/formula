use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn array_literal_spills_rectangular_block() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "={1,2;3,4}")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(2.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(3.0));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::Number(4.0));
}

#[test]
fn array_literal_can_be_used_as_a_function_argument() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=SUM({1,2;3,4})")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(10.0));
}

#[test]
fn array_literal_preserves_scalar_types_and_errors() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "={\"a\",TRUE;#VALUE!,\"b\"}")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B2").unwrap());

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("a".to_string())
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Bool(true));
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Text("b".to_string())
    );
}

#[test]
fn array_literal_rejects_ragged_rows() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "={1,2;3}")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Value)
    );
    assert!(engine.spill_range("Sheet1", "A1").is_none());
}

#[test]
fn array_literal_coerces_nested_arrays_to_value_errors() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "={SEQUENCE(2),1}")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B1").unwrap());

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(1.0));
}
