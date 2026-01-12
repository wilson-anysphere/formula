use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn fvschedule_accepts_array_literals() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=FVSCHEDULE(100,{0.1,0.2})"), 132.0);
}

#[test]
fn fvschedule_accepts_ranges() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 0.1);
    sheet.set("A2", 0.2);
    assert_number(&sheet.eval("=FVSCHEDULE(100,A1:A2)"), 132.0);
}

#[test]
fn fvschedule_propagates_errors_from_schedule() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 0.1);
    sheet.set("A2", Value::Error(ErrorKind::Div0));
    assert_eq!(
        sheet.eval("=FVSCHEDULE(100,A1:A2)"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn fvschedule_text_in_schedule_is_zero_rate() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 0.1);
    sheet.set("A2", Value::Text("not a number".to_string()));
    // Excel: text inside a reference/range does not participate.
    // For FVSCHEDULE, treating it as a 0% rate yields the same behavior.
    assert_number(&sheet.eval("=FVSCHEDULE(100,A1:A2)"), 110.0);
}

#[test]
fn fvschedule_blank_in_schedule_is_zero_rate() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 0.1);
    sheet.set("A2", Value::Blank);
    sheet.set("A3", 0.2);
    assert_number(&sheet.eval("=FVSCHEDULE(100,A1:A3)"), 132.0);
}
