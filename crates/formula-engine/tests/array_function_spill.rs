use formula_engine::eval::parse_a1;
use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn scalar_functions_lift_over_array_arguments() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=ABS({-1;-2})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C1", "=NOT({TRUE,FALSE})")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "E1", "=IF({TRUE,FALSE}, {1,2}, 0)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=LEFT({\"hello\";\"world\"},2)")
        .unwrap();

    engine.recalculate_single_threaded();

    // ABS({-1;-2}) -> {1;2}
    let (start, end) = engine.spill_range("Sheet1", "A1").expect("ABS spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("A2").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Number(2.0));

    // NOT({TRUE,FALSE}) -> {FALSE,TRUE}
    let (start, end) = engine.spill_range("Sheet1", "C1").expect("NOT spill range");
    assert_eq!(start, parse_a1("C1").unwrap());
    assert_eq!(end, parse_a1("D1").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Bool(false));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::Bool(true));

    // IF({TRUE,FALSE}, {1,2}, 0) -> {1,0}
    let (start, end) = engine.spill_range("Sheet1", "E1").expect("IF spill range");
    assert_eq!(start, parse_a1("E1").unwrap());
    assert_eq!(end, parse_a1("F1").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "E1"), Value::Number(1.0));
    assert_eq!(engine.get_cell_value("Sheet1", "F1"), Value::Number(0.0));

    // LEFT({"hello";"world"}, 2) -> {"he";"wo"}
    let (start, end) = engine
        .spill_range("Sheet1", "A4")
        .expect("LEFT spill range");
    assert_eq!(start, parse_a1("A4").unwrap());
    assert_eq!(end, parse_a1("A5").unwrap());
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Text("he".to_string())
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A5"),
        Value::Text("wo".to_string())
    );
}

#[test]
fn scalar_function_lifting_returns_value_for_shape_mismatches() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=IF({TRUE;FALSE}, {1,2}, 0)")
        .unwrap();
    engine.recalculate_single_threaded();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Value)
    );
    assert!(engine.spill_range("Sheet1", "A1").is_none());
}
