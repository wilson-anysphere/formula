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

    // column_num defaults to 1, so row_num=0 returns the first column slice.
    assert_number(&sheet.eval("=SUM(INDEX(A1:A3,0))"), 6.0);
    assert_eq!(sheet.eval("=INDEX(A1:A3,4)"), Value::Error(ErrorKind::Ref));
}

#[test]
fn index_truncates_fractional_indices_toward_zero() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 10.0);
    sheet.set("A2", 20.0);
    sheet.set("A3", 30.0);
    sheet.set("B1", 1.0);
    sheet.set("C1", 2.0);

    // row_num/col_num are coerced via truncation toward 0.
    assert_eq!(sheet.eval("=INDEX(A1:A3,2.9)"), Value::Number(20.0));
    assert_eq!(sheet.eval("=INDEX(A1:C1,1,2.1)"), Value::Number(1.0));

    // Negative indices become #VALUE! after truncation.
    assert_eq!(sheet.eval("=INDEX(A1:A3,-1.1)"), Value::Error(ErrorKind::Value));
    assert_eq!(sheet.eval("=INDEX(A1:C1,1,-2.9)"), Value::Error(ErrorKind::Value));
}

#[test]
fn index_area_num_defaults_and_bounds() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 1.0);
    sheet.set("A2", 2.0);
    sheet.set("A3", 3.0);
    sheet.set("C1", 10.0);
    sheet.set("C2", 20.0);
    sheet.set("C3", 30.0);

    // Default area_num is 1.
    assert_eq!(
        sheet.eval("=INDEX((A1:A3,C1:C3),2,1)"),
        Value::Number(2.0)
    );

    // Out of bounds area_num -> #REF!
    assert_eq!(
        sheet.eval("=INDEX((A1:A3,C1:C3),2,1,0)"),
        Value::Error(ErrorKind::Ref)
    );
    assert_eq!(
        sheet.eval("=INDEX((A1:A3,C1:C3),2,1,3)"),
        Value::Error(ErrorKind::Ref)
    );

    // Passing area_num for a single-area reference behaves like Excel and errors unless it's 1.
    assert_eq!(
        sheet.eval("=INDEX(A1:A3,2,1,2)"),
        Value::Error(ErrorKind::Ref)
    );
}
