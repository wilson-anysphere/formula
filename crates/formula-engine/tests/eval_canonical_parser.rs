use formula_engine::{Engine, ErrorKind, Value};

#[test]
fn evaluates_pow_operator() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=2^3").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(8.0));
}

#[test]
fn evaluates_pow_precedence_over_add() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=2+3^2").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(11.0));
}

#[test]
fn evaluates_percent_postfix() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=50%").unwrap();
    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(0.5));
}

#[test]
fn evaluates_concat_operator() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", r#"="a"&"b""#)
        .unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Text("ab".to_string())
    );
}

#[test]
fn parses_error_literals() {
    let mut engine = Engine::new();
    engine.set_cell_formula("Sheet1", "A1", "=#NUM!").unwrap();
    engine.recalculate();
    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn parses_pow_formulas_from_oracle_corpus() {
    // `tests/compatibility/excel-oracle/cases.json` includes many `=A1^B1` cases; these previously
    // failed to parse via the legacy eval parser.
    let mut engine = Engine::new();

    // arith_pow_8a3211580106: 2^2
    engine.set_cell_value("Sheet1", "A1", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 2.0).unwrap();
    engine.set_cell_formula("Sheet1", "C1", "=A1^B1").unwrap();

    // arith_pow_7cd40eba4d06: 10^-1
    engine.set_cell_value("Sheet1", "A2", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", -1.0).unwrap();
    engine.set_cell_formula("Sheet1", "C2", "=A2^B2").unwrap();

    engine.recalculate();
    assert_eq!(engine.get_cell_value("Sheet1", "C1"), Value::Number(4.0));
    assert_eq!(engine.get_cell_value("Sheet1", "C2"), Value::Number(0.1));
}

