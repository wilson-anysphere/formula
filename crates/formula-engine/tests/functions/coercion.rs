use formula_engine::Value;

use super::harness::{assert_number, TestSheet};

#[test]
fn empty_string_coerces_to_zero_in_arithmetic() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=1+\"\""), 1.0);
    assert_number(&sheet.eval("=--\"\""), 0.0);
}

#[test]
fn empty_string_coerces_to_false_in_logical_contexts() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=NOT(\"\")"), Value::Bool(true));
    assert_number(&sheet.eval("=IF(\"\",10,20)"), 20.0);
}

#[test]
fn text_numbers_parse_like_excel_value() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=\"1,234\"+1"), 1235.0);
    assert_number(&sheet.eval("=\"10%\"*100"), 10.0);
}

#[test]
fn text_dates_coerce_to_serial_numbers() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=YEAR(\"2020-01-01\")"), 2020.0);
}
