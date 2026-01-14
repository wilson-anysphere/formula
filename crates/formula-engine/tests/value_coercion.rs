use formula_engine::{Engine, ErrorKind, Value};

fn assert_number(value: &Value, expected: f64) {
    match value {
        Value::Number(n) => {
            assert!((*n - expected).abs() < 1e-9, "expected {expected}, got {n}");
        }
        other => panic!("expected number {expected}, got {other:?}"),
    }
}

fn eval(formula: &str) -> Value {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", formula)
        .expect("set formula");
    engine.recalculate();
    engine.get_cell_value("Sheet1", "A1")
}

#[test]
fn text_to_number_coercion_accepts_common_excel_forms() {
    assert_eq!(
        Value::Text(" 1,234.50 ".to_string())
            .coerce_to_number()
            .unwrap(),
        1234.5
    );
    assert_eq!(
        Value::Text("(1,000)".to_string())
            .coerce_to_number()
            .unwrap(),
        -1000.0
    );
    assert_eq!(
        Value::Text("$12,345.67".to_string())
            .coerce_to_number()
            .unwrap(),
        12345.67
    );
    assert_eq!(
        Value::Text("12%".to_string()).coerce_to_number().unwrap(),
        0.12
    );
    assert_eq!(
        Value::Text("1e3".to_string()).coerce_to_number().unwrap(),
        1000.0
    );

    // Excel normalizes negative zero to 0 in coercions.
    assert_eq!(
        Value::Text("-0".to_string()).coerce_to_number().unwrap(),
        0.0
    );
    assert_eq!(Value::Text("".to_string()).coerce_to_number().unwrap(), 0.0);
    assert_eq!(
        Value::Text("   ".to_string()).coerce_to_number().unwrap(),
        0.0
    );

    assert_eq!(
        Value::Text("INF".to_string())
            .coerce_to_number()
            .unwrap_err(),
        ErrorKind::Value
    );
    assert_eq!(
        Value::Text("1e309".to_string())
            .coerce_to_number()
            .unwrap_err(),
        ErrorKind::Num
    );
}

#[test]
fn text_to_bool_coercion_matches_numeric_coercion() {
    assert_eq!(
        Value::Text("TRUE".to_string()).coerce_to_bool().unwrap(),
        true
    );
    assert_eq!(
        Value::Text("FALSE".to_string()).coerce_to_bool().unwrap(),
        false
    );
    assert_eq!(
        Value::Text("  0  ".to_string()).coerce_to_bool().unwrap(),
        false
    );
    assert_eq!(
        Value::Text("  2  ".to_string()).coerce_to_bool().unwrap(),
        true
    );
}

#[test]
fn number_to_text_general_formatting_is_excel_like() {
    assert_eq!(Value::Number(0.0).coerce_to_string().unwrap(), "0");
    assert_eq!(Value::Number(-0.0).coerce_to_string().unwrap(), "0");
    assert_eq!(Value::Number(1.0).coerce_to_string().unwrap(), "1");
    assert_eq!(Value::Number(1.25).coerce_to_string().unwrap(), "1.25");
    // Preserve 15 significant digits even for small magnitudes (< 0.1).
    assert_eq!(
        Value::Number(0.0123456789012345)
            .coerce_to_string()
            .unwrap(),
        "0.0123456789012345"
    );
    assert_eq!(
        Value::Number(0.000123456789012345)
            .coerce_to_string()
            .unwrap(),
        "0.000123456789012345"
    );
    // General switches to scientific at 1e11 but not 1e10.
    assert_eq!(
        Value::Number(1e10).coerce_to_string().unwrap(),
        "10000000000"
    );
    assert_eq!(Value::Number(1e11).coerce_to_string().unwrap(), "1E+11");
    assert_eq!(Value::Number(1e20).coerce_to_string().unwrap(), "1E+20");
    assert_eq!(Value::Number(1e-10).coerce_to_string().unwrap(), "1E-10");
}

#[test]
fn engine_formulas_use_excel_like_implicit_coercions() {
    assert_number(&eval("=\"$1,200\"+1"), 1201.0);
    assert_number(&eval("=--\"12%\""), 0.12);
    assert_number(&eval("=IF(\"0\",1,2)"), 2.0);

    assert_eq!(eval("=\"x\"&1"), Value::Text("x1".to_string()));

    assert_number(&eval("=SUM(\"1\",\"2\")"), 3.0);

    let mut engine = Engine::new();
    engine
        .set_cell_value("Sheet1", "A1", Value::Text("1".to_string()))
        .unwrap();
    engine
        .set_cell_value("Sheet1", "A2", Value::Text("2".to_string()))
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=SUM(A1:A2)")
        .unwrap();
    engine.recalculate();
    assert_number(&engine.get_cell_value("Sheet1", "A3"), 0.0);
}
