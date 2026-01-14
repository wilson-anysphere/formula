use formula_engine::value::ErrorKind;
use formula_engine::Value;

use super::harness::{assert_number, TestSheet};

#[test]
fn fvschedule_simple_array_literal() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=FVSCHEDULE(100,{0.1,0.2})"), 132.0);
}

#[test]
fn fvschedule_range_mixed_cell_types() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", 0.1);
    sheet.set("A2", true);
    sheet.set("A3", 0.2);
    sheet.set("A4", "not a number");
    // A5 left blank.

    assert_number(&sheet.eval("=FVSCHEDULE(100,A1:A5)"), 132.0);
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
fn fvschedule_propagates_errors_in_schedule() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", 0.1);
    sheet.set_formula("A2", "=1/0");

    assert_eq!(
        sheet.eval("=FVSCHEDULE(100,A1:A2)"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn fvschedule_propagates_errors_from_union_schedule() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 0.1);
    sheet.set_formula("A2", "=1/0");
    sheet.set("B1", 0.2);

    assert_eq!(
        sheet.eval("=FVSCHEDULE(100,(A1:A2,B1))"),
        Value::Error(ErrorKind::Div0)
    );
}
