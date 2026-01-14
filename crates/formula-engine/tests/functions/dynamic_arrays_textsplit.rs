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
fn textsplit_column_delimiters_can_be_array_literals() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"a,b;c,d\",{\",\",\";\"})")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("D1").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::from("a"));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::from("b"));
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::from("c"));
    assert_eq!(engine.get_cell_value("Sheet1", "D1"), Value::from("d"));
}

#[test]
fn textsplit_row_delimiters_can_be_array_literals() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=TEXTSPLIT(\"a,b|c,d;e,f\",\",\",{\"|\",\";\"})",
        )
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B3").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::from("a"));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::from("b"));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::from("c"));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::from("d"));
    assert_eq!(engine.get_cell_value("Sheet1", "A3"), Value::from("e"));
    assert_eq!(engine.get_cell_value("Sheet1", "B3"), Value::from("f"));
}

#[test]
fn textsplit_blank_row_delimiter_is_no_row_split() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"a;b\",\",\",\"\")")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("A1").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::from("a;b"));
}

#[test]
fn textsplit_row_delimiter_array_literal_rejects_empty_elements() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"a;b\",\",\",{\";\",\"\"})")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Value)
    );
}

#[test]
fn textsplit_delimiter_array_literals_propagate_errors() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"a,b\",{\",\",#DIV/0!})")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Div0)
    );
}

#[test]
fn textsplit_ignore_empty_applies_to_row_splitting() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"a;;b\",\",\",\";\",TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("A2").unwrap());
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::from("a"));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::from("b"));
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
fn textsplit_blank_pad_with_defaults_to_na() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"a,b;c\",\",\",\";\",FALSE,0,)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::from("a"));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::from("b"));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::from("c"));
    assert_eq!(
        engine.get_cell_value("Sheet1", "B2"),
        Value::Error(formula_engine::ErrorKind::NA)
    );
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

#[test]
fn textsplit_match_mode_case_insensitive_is_unicode_aware() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"aMaßb\",\"MASS\",,FALSE,1)")
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B1").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::from("a"));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::from("b"));
}

#[test]
fn textsplit_keeps_rows_that_become_empty_after_column_split() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=TEXTSPLIT(\"a,b;,,\",\",\",\";\",TRUE,0,\"x\")",
        )
        .unwrap();
    engine.recalculate_single_threaded();

    let (start, end) = engine.spill_range("Sheet1", "A1").expect("spill range");
    assert_eq!(start, parse_a1("A1").unwrap());
    assert_eq!(end, parse_a1("B2").unwrap());

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::from("a"));
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::from("b"));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::from("x"));
    assert_eq!(engine.get_cell_value("Sheet1", "B2"), Value::from("x"));
}

#[test]
fn textsplit_rejects_empty_column_delimiter() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"a,b\",\"\")")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Value)
    );
}

#[test]
fn textsplit_ignore_empty_can_return_calc_when_all_segments_removed() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\",,,\",\",\",,TRUE)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Calc)
    );
}

#[test]
fn textsplit_invalid_match_mode_is_value_error() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"aXb\",\"x\",,FALSE,2)")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Value)
    );
}

#[test]
fn textsplit_rejects_array_pad_with() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=TEXTSPLIT(\"a,b\",\",\",,FALSE,0,{\"x\"})")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(formula_engine::ErrorKind::Value)
    );
}
