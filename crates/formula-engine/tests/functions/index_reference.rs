use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn index_reference_enables_range_operator_dynamic_ranges() {
    let mut sheet = TestSheet::new();
    for i in 1..=10 {
        sheet.set(&format!("A{i}"), i as f64);
    }

    assert_number(&sheet.eval("=SUM(INDEX(A1:A10,2):INDEX(A1:A10,5))"), 14.0);
}

#[test]
fn index_reference_supports_row_and_column_zero_slices() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("B1", 2.0);
    sheet.set("C1", 3.0);

    sheet.set("A2", 10.0);
    sheet.set("B2", 20.0);
    sheet.set("C2", 30.0);

    sheet.set("A3", 100.0);
    sheet.set("B3", 200.0);
    sheet.set("C3", 300.0);

    assert_number(&sheet.eval("=SUM(INDEX(A1:C3,0,2))"), 222.0);
    assert_number(&sheet.eval("=SUM(INDEX(A1:C3,2,0))"), 60.0);
    assert_number(&sheet.eval("=SUM(INDEX(A1:C3,0,0))"), 666.0);
}

#[test]
fn index_reference_supports_multi_area_union_and_area_num() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);
    sheet.set("C1", 10.0);
    sheet.set("C2", 20.0);
    sheet.set("C3", 30.0);

    assert_eq!(
        sheet.eval("=INDEX((A1:A3,C1:C3),2,1,2)"),
        Value::Number(20.0)
    );
}

#[test]
fn index_reference_errors() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);

    // Excel requires column_num to be provided when row_num is 0.
    assert_eq!(sheet.eval("=INDEX(A1:A3,0)"), Value::Error(ErrorKind::Value));
    assert_eq!(sheet.eval("=INDEX(A1:A3,4)"), Value::Error(ErrorKind::Ref));
}

