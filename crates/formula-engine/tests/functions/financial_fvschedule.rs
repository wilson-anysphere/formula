use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn fvschedule_accepts_array_literal_schedule() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=FVSCHEDULE(100,{0.1,0.2})"), 132.0);
}

#[test]
fn fvschedule_accepts_range_schedule() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 0.1);
    sheet.set("A2", 0.2);
    sheet.set("A3", 0.3);

    assert_number(&sheet.eval("=FVSCHEDULE(100,A1:A3)"), 171.6);
}

#[test]
fn fvschedule_accepts_union_schedule() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 0.1);
    sheet.set("A2", 0.2);
    sheet.set("B1", 0.3);
    sheet.set("B2", 0.4);

    // Union references use the `,` operator; in a function argument, it must be parenthesized.
    assert_number(&sheet.eval("=FVSCHEDULE(100,(A1:A2,B1:B2))"), 240.24);
}

#[test]
fn fvschedule_propagates_schedule_errors() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 0.1);
    sheet.set_formula("A2", "=1/0");
    sheet.set("A3", 0.2);

    assert_eq!(
        sheet.eval("=FVSCHEDULE(100,A1:A3)"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn fvschedule_propagates_schedule_errors_from_union() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 0.1);
    sheet.set_formula("A2", "=1/0");
    sheet.set("B1", 0.2);

    assert_eq!(
        sheet.eval("=FVSCHEDULE(100,(A1:A2,B1))"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn fvschedule_coerces_blank_and_numeric_text_schedule_cells() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 0.1);
    sheet.set("A2", Value::Blank);
    sheet.set("A3", "0.2");

    // Blank is treated as 0%, and numeric text is parsed as a number.
    assert_number(&sheet.eval("=FVSCHEDULE(100,A1:A3)"), 132.0);
}

#[test]
fn fvschedule_returns_value_for_non_numeric_text_in_schedule() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 0.1);
    sheet.set("A2", "not a number");

    assert_eq!(
        sheet.eval("=FVSCHEDULE(100,A1:A2)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn fvschedule_coerces_bool_schedule_cells() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", true);

    assert_number(&sheet.eval("=FVSCHEDULE(100,TRUE)"), 200.0);
    assert_number(&sheet.eval("=FVSCHEDULE(100,A1)"), 200.0);
}
