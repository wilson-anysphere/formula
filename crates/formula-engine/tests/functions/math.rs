use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn sum_ignores_text_in_ranges_but_coerces_scalar_text() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("5".to_string()));
    sheet.set("A2", 3.0);
    sheet.set("A3", 4.0);

    assert_number(&sheet.eval("=SUM(A1:A3)"), 7.0);
    assert_number(&sheet.eval(r#"=SUM("5", TRUE, 3)"#), 9.0);
}

#[test]
fn sum_propagates_errors() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", Value::Error(ErrorKind::Div0));
    assert_eq!(sheet.eval("=SUM(A1:A2)"), Value::Error(ErrorKind::Div0));
}

#[test]
fn average_div0_when_no_numeric_values() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("x".to_string()));
    sheet.set("A2", Value::Blank);
    assert_eq!(sheet.eval("=AVERAGE(A1:A2)"), Value::Error(ErrorKind::Div0));
}

#[test]
fn min_max_ignore_text_in_ranges() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Text("100".to_string()));
    sheet.set("A2", 3.0);
    sheet.set("A3", 4.0);

    assert_number(&sheet.eval("=MIN(A1:A3)"), 3.0);
    assert_number(&sheet.eval("=MAX(A1:A3)"), 4.0);
    assert_number(&sheet.eval(r#"=MIN("5", TRUE, 3)"#), 1.0);
}

#[test]
fn count_counta_countblank() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", Value::Text("x".to_string()));
    sheet.set("A3", true);
    sheet.set("A4", Value::Blank);
    sheet.set("A5", Value::Text("".to_string()));
    sheet.set("A6", Value::Error(ErrorKind::Div0));

    assert_number(&sheet.eval("=COUNT(A1:A6)"), 1.0);
    assert_number(&sheet.eval("=COUNTA(A1:A6)"), 5.0);
    assert_number(&sheet.eval("=COUNTBLANK(A1:A6)"), 2.0);
}

#[test]
fn round_variants() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=ROUND(2.5,0)"), 3.0);
    assert_number(&sheet.eval("=ROUND(-2.5,0)"), -3.0);
    assert_number(&sheet.eval("=ROUND(1234,-2)"), 1200.0);

    assert_number(&sheet.eval("=ROUNDDOWN(1.29,1)"), 1.2);
    assert_number(&sheet.eval("=ROUNDDOWN(-1.29,1)"), -1.2);
    assert_number(&sheet.eval("=ROUNDUP(1.21,1)"), 1.3);
    assert_number(&sheet.eval("=ROUNDUP(-1.21,1)"), -1.3);
}

#[test]
fn int_abs_mod() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=INT(2.9)"), 2.0);
    assert_number(&sheet.eval("=INT(-2.1)"), -3.0);

    assert_number(&sheet.eval("=ABS(-3)"), 3.0);

    assert_number(&sheet.eval("=MOD(5,2)"), 1.0);
    assert_number(&sheet.eval("=MOD(-3,2)"), 1.0);
    assert_number(&sheet.eval("=MOD(3,-2)"), -1.0);
    assert_eq!(sheet.eval("=MOD(5,0)"), Value::Error(ErrorKind::Div0));
}

#[test]
fn sign_returns_expected_signum() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=SIGN(-2)"), -1.0);
    assert_number(&sheet.eval("=SIGN(0)"), 0.0);
    assert_number(&sheet.eval("=SIGN(2)"), 1.0);
}
