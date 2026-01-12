use formula_engine::eval::parse_a1;
use formula_engine::{Engine, Value};

#[test]
fn textsplit_basic_columns() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"a,b,c\",\",\")")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("C1").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::from("a"));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::from("b"));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::from("c"));
}

#[test]
fn textsplit_rows_and_columns() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"a,b;c,d\",\",\",\";\")")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::from("a"));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::from("b"));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::from("c"));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::from("d"));
}

#[test]
fn textsplit_ignore_empty() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"a,,b\",\",\",,TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B1").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::from("a"));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::from("b"));
}

#[test]
fn textsplit_pad_with() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=TEXTSPLIT(\"a,b;c\",\",\",\";\",FALSE,0,\"x\")",
        )
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::from("a"));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::from("b"));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::from("c"));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::from("x"));
}

#[test]
fn textsplit_match_mode_case_insensitive() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"aXb\",\"x\",,FALSE,1)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B1").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::from("a"));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::from("b"));
}

