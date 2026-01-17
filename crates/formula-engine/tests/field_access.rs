use formula_engine::{
    parse_formula, parse_formula_partial, Engine, ErrorKind, ParseOptions, SerializeOptions, Value,
};
use formula_engine::value::RecordValue;
use std::collections::HashMap;

#[test]
fn field_access_roundtrip_ident() {
    let formula = "=A1.Price";
    let ast = parse_formula(formula, ParseOptions::default()).expect("parse");
    let out = ast
        .to_string(SerializeOptions::default())
        .expect("serialize");
    assert_eq!(out, formula);
}

#[test]
fn field_access_roundtrip_bracketed() {
    let formula = r#"=A1.["Change%"]"#;
    let ast = parse_formula(formula, ParseOptions::default()).expect("parse");
    let out = ast
        .to_string(SerializeOptions::default())
        .expect("serialize");
    assert_eq!(out, formula);
}

#[test]
fn field_access_roundtrip_nested() {
    let formula = "=A1.Address.City";
    let ast = parse_formula(formula, ParseOptions::default()).expect("parse");
    let out = ast
        .to_string(SerializeOptions::default())
        .expect("serialize");
    assert_eq!(out, formula);
}

#[test]
fn field_access_partial_parse_trailing_dot() {
    let partial = parse_formula_partial("=A1.", ParseOptions::default());
    assert!(
        partial.error.is_some(),
        "expected partial parse to capture an error"
    );
}

#[test]
fn field_access_evaluates_to_value_error_for_non_rich_values() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "A1", 123.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=A1.Price")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn field_access_propagates_base_error() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=1/0").unwrap();
    engine
        .set_cell_formula("Sheet1", "B1", "=A1.Price")
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(
        engine.get_cell_value("Sheet1", "B1"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn field_access_bracket_selector_allows_outer_whitespace() {
    let mut engine = Engine::new();

    let mut fields = HashMap::new();
    fields.insert("Change%".to_string(), Value::Number(3.0));
    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Record(RecordValue::with_fields("Record", fields)),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", r#"=A1.[  "Change%"  ]"#)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
}

#[test]
fn field_access_bracket_selector_allows_whitespace_between_dot_and_bracket() {
    let mut engine = Engine::new();

    let mut fields = HashMap::new();
    fields.insert("Change%".to_string(), Value::Number(3.0));
    engine
        .set_cell_value(
            "Sheet1",
            "A1",
            Value::Record(RecordValue::with_fields("Record", fields)),
        )
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "B1", r#"=A1. ["Change%"]"#)
        .unwrap();
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "B1"), Value::Number(3.0));
}
