use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn kurt_matches_known_value() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=KURT({1,2,3,4,5})"), -1.2);
}

#[test]
fn skew_and_skew_p_match_known_values() {
    let mut sheet = TestSheet::new();

    // Dataset with known exact results:
    // - SKEW = 9*sqrt(5)/16
    // - SKEW.P = 27/32
    assert_number(
        &sheet.eval("=SKEW({1,1,1,2,3})"),
        9.0 * 5.0_f64.sqrt() / 16.0,
    );
    assert_number(&sheet.eval("=SKEW.P({1,1,1,2,3})"), 27.0 / 32.0);
}

#[test]
fn skew_and_kurt_error_on_too_few_points() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=SKEW({1,2})"), Value::Error(ErrorKind::Div0));
    assert_eq!(sheet.eval("=SKEW.P({1})"), Value::Error(ErrorKind::Div0));
    assert_eq!(sheet.eval("=KURT({1,2,3})"), Value::Error(ErrorKind::Div0));
}

#[test]
fn skew_ignores_text_and_logicals_in_references() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", Value::Text("2".to_string()));
    sheet.set("A3", true);

    // In references, text/bools are ignored, leaving a single numeric value.
    assert_eq!(sheet.eval("=SKEW(A1:A3)"), Value::Error(ErrorKind::Div0));

    // As direct scalar arguments, numeric text/bools are coerced.
    assert_number(&sheet.eval(r#"=SKEW(1,"2",TRUE)"#), 3.0_f64.sqrt());
}

#[test]
fn frequency_spills_expected_counts_from_literals() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("Z1", "=FREQUENCY({79,85,78,85,50,81,95,88,97},{70,79,89})");
    sheet.recalc();

    assert_eq!(sheet.get("Z1"), Value::Number(1.0));
    assert_eq!(sheet.get("Z2"), Value::Number(2.0));
    assert_eq!(sheet.get("Z3"), Value::Number(4.0));
    assert_eq!(sheet.get("Z4"), Value::Number(2.0));

    // Verify spill shape (vertical).
    assert_eq!(sheet.get("AA1"), Value::Blank);
    assert_eq!(sheet.get("Z5"), Value::Blank);
}

#[test]
fn frequency_ignores_text_in_data_range_and_propagates_errors() {
    let mut sheet = TestSheet::new();

    // Seed a data range with a text value (ignored).
    let data = [79.0, 85.0, 78.0, 85.0, 0.0, 81.0, 95.0, 88.0, 97.0];
    for (idx, v) in data.iter().enumerate() {
        let addr = format!("A{}", idx + 1);
        if idx == 4 {
            sheet.set(&addr, Value::Text("x".to_string()));
        } else {
            sheet.set(&addr, *v);
        }
    }
    sheet.set("B1", 70.0);
    sheet.set("B2", 79.0);
    sheet.set("B3", 89.0);

    sheet.set_formula("Z1", "=FREQUENCY(A1:A9,B1:B3)");
    sheet.recalc();

    assert_eq!(sheet.get("Z1"), Value::Number(0.0));
    assert_eq!(sheet.get("Z2"), Value::Number(2.0));
    assert_eq!(sheet.get("Z3"), Value::Number(4.0));
    assert_eq!(sheet.get("Z4"), Value::Number(2.0));
    assert_eq!(sheet.get("Z5"), Value::Blank);

    // Errors propagate.
    sheet.set("A1", Value::Error(ErrorKind::Div0));
    assert_eq!(
        sheet.eval("=FREQUENCY(A1:A2,B1:B1)"),
        Value::Error(ErrorKind::Div0)
    );
}
