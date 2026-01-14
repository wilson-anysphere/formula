use formula_engine::{ErrorKind, Value};

use super::harness::TestSheet;

#[test]
fn frequency_basic_histogram_spills_vertical() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("Z1", "=FREQUENCY({1,2,3,4,5},{2,4})");
    sheet.recalc();

    assert_eq!(sheet.get("Z1"), Value::Number(2.0));
    assert_eq!(sheet.get("Z2"), Value::Number(2.0));
    assert_eq!(sheet.get("Z3"), Value::Number(1.0));
    assert_eq!(sheet.get("Z4"), Value::Blank);
}

#[test]
fn frequency_ignores_non_numeric_data_values() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1);
    sheet.set("A2", Value::Text("x".to_string()));
    sheet.set("A3", Value::Bool(true));
    sheet.set("A4", Value::Blank);
    sheet.set("A5", 2);
    sheet.set("A6", 3);

    sheet.set_formula("Z1", "=FREQUENCY(A1:A6,{2})");
    sheet.recalc();

    assert_eq!(sheet.get("Z1"), Value::Number(2.0));
    assert_eq!(sheet.get("Z2"), Value::Number(1.0));
    assert_eq!(sheet.get("Z3"), Value::Blank);
}

#[test]
fn frequency_propagates_errors_from_data_array() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1);
    sheet.set_formula("A2", "=NA()");
    sheet.set("A3", 2);

    assert_eq!(
        sheet.eval("=FREQUENCY(A1:A3,{2})"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn frequency_coerces_bin_text_values() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("Z1", r#"=FREQUENCY({1,2,3,4,5},{"2","4"})"#);
    sheet.recalc();

    assert_eq!(sheet.get("Z1"), Value::Number(2.0));
    assert_eq!(sheet.get("Z2"), Value::Number(2.0));
    assert_eq!(sheet.get("Z3"), Value::Number(1.0));
}

#[test]
fn frequency_propagates_errors_from_bins_array() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1);
    sheet.set("A2", 2);
    sheet.set("B1", 2);
    sheet.set_formula("B2", "=NA()");

    assert_eq!(
        sheet.eval("=FREQUENCY(A1:A2,B1:B2)"),
        Value::Error(ErrorKind::NA)
    );
}
