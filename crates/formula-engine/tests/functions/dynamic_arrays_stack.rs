use formula_engine::{ErrorKind, Value};

use super::harness::TestSheet;

#[test]
fn hstack_pads_missing_rows_with_na() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=HSTACK({1;2;3},{10;20})");
    sheet.recalc();

    assert_eq!(sheet.get("A1"), Value::Number(1.0));
    assert_eq!(sheet.get("B1"), Value::Number(10.0));
    assert_eq!(sheet.get("A2"), Value::Number(2.0));
    assert_eq!(sheet.get("B2"), Value::Number(20.0));
    assert_eq!(sheet.get("A3"), Value::Number(3.0));
    assert_eq!(sheet.get("B3"), Value::Error(ErrorKind::NA));
}

#[test]
fn vstack_pads_missing_cols_with_na() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=VSTACK({1,2},{3})");
    sheet.recalc();

    assert_eq!(sheet.get("A1"), Value::Number(1.0));
    assert_eq!(sheet.get("B1"), Value::Number(2.0));
    assert_eq!(sheet.get("A2"), Value::Number(3.0));
    assert_eq!(sheet.get("B2"), Value::Error(ErrorKind::NA));
}

#[test]
fn tocol_torow_ordering_and_ignore_blanks() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=TOCOL({1,2;3,4})");
    sheet.set_formula("C1", "=TOCOL({1,2;3,4},,TRUE)");
    sheet.set_formula("E1", "=TOCOL({1,,3},1)");
    sheet.set_formula("A6", "=TOROW({1,2;3,4})");
    sheet.set_formula("A8", "=TOROW({1,2;3,4},,TRUE)");
    sheet.recalc();

    // TOCOL row-major.
    assert_eq!(sheet.get("A1"), Value::Number(1.0));
    assert_eq!(sheet.get("A2"), Value::Number(2.0));
    assert_eq!(sheet.get("A3"), Value::Number(3.0));
    assert_eq!(sheet.get("A4"), Value::Number(4.0));

    // TOCOL column-major.
    assert_eq!(sheet.get("C1"), Value::Number(1.0));
    assert_eq!(sheet.get("C2"), Value::Number(3.0));
    assert_eq!(sheet.get("C3"), Value::Number(2.0));
    assert_eq!(sheet.get("C4"), Value::Number(4.0));

    // TOCOL ignore blanks.
    assert_eq!(sheet.get("E1"), Value::Number(1.0));
    assert_eq!(sheet.get("E2"), Value::Number(3.0));

    // TOROW row-major.
    assert_eq!(sheet.get("A6"), Value::Number(1.0));
    assert_eq!(sheet.get("B6"), Value::Number(2.0));
    assert_eq!(sheet.get("C6"), Value::Number(3.0));
    assert_eq!(sheet.get("D6"), Value::Number(4.0));

    // TOROW column-major.
    assert_eq!(sheet.get("A8"), Value::Number(1.0));
    assert_eq!(sheet.get("B8"), Value::Number(3.0));
    assert_eq!(sheet.get("C8"), Value::Number(2.0));
    assert_eq!(sheet.get("D8"), Value::Number(4.0));
}

#[test]
fn tocol_returns_calc_when_all_values_are_ignored() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", r#"=IFERROR(TOCOL({,},1),"fallback")"#);
    sheet.recalc();

    assert_eq!(
        sheet.get("A1"),
        Value::Text("fallback".to_string()),
        "TOCOL with ignore_blanks should produce #CALC! and be caught by IFERROR"
    );
}

#[test]
fn wraprows_wrapcols_wrap_sequence_and_pad() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=WRAPROWS(SEQUENCE(5),2)");
    sheet.set_formula("D1", "=WRAPROWS(SEQUENCE(5),2,0)");
    sheet.set_formula("A6", "=WRAPCOLS(SEQUENCE(5),2)");
    sheet.set_formula("A9", "=WRAPCOLS(SEQUENCE(5),2,0)");
    sheet.recalc();

    // WRAPROWS default pad (#N/A).
    assert_eq!(sheet.get("A1"), Value::Number(1.0));
    assert_eq!(sheet.get("B1"), Value::Number(2.0));
    assert_eq!(sheet.get("A2"), Value::Number(3.0));
    assert_eq!(sheet.get("B2"), Value::Number(4.0));
    assert_eq!(sheet.get("A3"), Value::Number(5.0));
    assert_eq!(sheet.get("B3"), Value::Error(ErrorKind::NA));

    // WRAPROWS custom pad.
    assert_eq!(sheet.get("D1"), Value::Number(1.0));
    assert_eq!(sheet.get("E1"), Value::Number(2.0));
    assert_eq!(sheet.get("D2"), Value::Number(3.0));
    assert_eq!(sheet.get("E2"), Value::Number(4.0));
    assert_eq!(sheet.get("D3"), Value::Number(5.0));
    assert_eq!(sheet.get("E3"), Value::Number(0.0));

    // WRAPCOLS default pad (#N/A).
    assert_eq!(sheet.get("A6"), Value::Number(1.0));
    assert_eq!(sheet.get("B6"), Value::Number(3.0));
    assert_eq!(sheet.get("C6"), Value::Number(5.0));
    assert_eq!(sheet.get("A7"), Value::Number(2.0));
    assert_eq!(sheet.get("B7"), Value::Number(4.0));
    assert_eq!(sheet.get("C7"), Value::Error(ErrorKind::NA));

    // WRAPCOLS custom pad.
    assert_eq!(sheet.get("A9"), Value::Number(1.0));
    assert_eq!(sheet.get("B9"), Value::Number(3.0));
    assert_eq!(sheet.get("C9"), Value::Number(5.0));
    assert_eq!(sheet.get("A10"), Value::Number(2.0));
    assert_eq!(sheet.get("B10"), Value::Number(4.0));
    assert_eq!(sheet.get("C10"), Value::Number(0.0));
}

#[test]
fn wraprows_wrapcols_blank_pad_argument_defaults_to_na() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=WRAPROWS(SEQUENCE(5),2,)");
    sheet.set_formula("A6", "=WRAPCOLS(SEQUENCE(5),2,)");
    sheet.recalc();

    // Trailing empty arg should behave like an omitted pad_with argument (default #N/A).
    assert_eq!(sheet.get("B3"), Value::Error(ErrorKind::NA));
    assert_eq!(sheet.get("C7"), Value::Error(ErrorKind::NA));
}

#[test]
fn wraprows_wrapcols_allow_error_pad_value() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=WRAPROWS(SEQUENCE(5),2,NA())");
    sheet.set_formula("A6", "=WRAPCOLS(SEQUENCE(5),2,NA())");
    sheet.recalc();

    // An explicit error pad value should be used as padding (not short-circuit the function).
    assert_eq!(sheet.get("A1"), Value::Number(1.0));
    assert_eq!(sheet.get("B3"), Value::Error(ErrorKind::NA));
    assert_eq!(sheet.get("C7"), Value::Error(ErrorKind::NA));
}
