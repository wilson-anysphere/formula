use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

fn assert_number_tol(value: &Value, expected: f64, tol: f64) {
    match value {
        Value::Number(n) => {
            assert!(
                (*n - expected).abs() <= tol,
                "expected {expected} ± {tol}, got {n}"
            );
        }
        other => panic!("expected number {expected}, got {other:?}"),
    }
}

#[test]
fn skew_symmetric_is_zero() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=SKEW({1,2,3,4,5})"), 0.0);
    assert_number(&sheet.eval("=SKEW.P({1,2,3,4,5})"), 0.0);
}

#[test]
fn skew_and_kurt_match_excel_example_values() {
    let mut sheet = TestSheet::new();

    // Examples from Excel docs:
    // - SKEW({3,4,5,2,3,4,5,6,4,7}) ≈ 0.3595430714
    // - KURT({3,4,5,2,3,4,5,6,4,7}) ≈ -0.1517996372
    assert_number_tol(
        &sheet.eval("=SKEW({3,4,5,2,3,4,5,6,4,7})"),
        0.3595430714067975,
        1e-9,
    );
    assert_number_tol(
        &sheet.eval("=SKEW.P({3,4,5,2,3,4,5,6,4,7})"),
        0.30319333935414394,
        1e-9,
    );
    assert_number_tol(
        &sheet.eval("=KURT({3,4,5,2,3,4,5,6,4,7})"),
        -0.15179963720841538,
        1e-9,
    );
}

#[test]
fn skew_and_kurt_domain_errors() {
    let mut sheet = TestSheet::new();

    assert_eq!(sheet.eval("=SKEW({1,2})"), Value::Error(ErrorKind::Div0));
    assert_eq!(sheet.eval("=KURT({1,2,3})"), Value::Error(ErrorKind::Div0));

    assert_eq!(sheet.eval("=SKEW({2,2,2})"), Value::Error(ErrorKind::Div0));
    assert_eq!(
        sheet.eval("=SKEW.P({2,2,2})"),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        sheet.eval("=KURT({2,2,2,2})"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn skew_ignores_text_and_logicals_in_references() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", Value::Text("2".to_string()));
    sheet.set("A3", true);
    sheet.set("A4", 3.0);

    // In references, text/bools are ignored, leaving {1,3} => insufficient for sample skewness.
    assert_eq!(sheet.eval("=SKEW(A1:A4)"), Value::Error(ErrorKind::Div0));

    // As direct scalar arguments, numeric text/bools are coerced.
    assert_number_tol(
        &sheet.eval(r#"=SKEW(1,"2",TRUE,3)"#),
        0.8545630383279712,
        1e-9,
    );
}

#[test]
fn moments_reject_lambda_values() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=SKEW({1,2,LAMBDA(x,x)})"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=KURT({1,2,3,LAMBDA(x,x)})"),
        Value::Error(ErrorKind::Value)
    );
}
