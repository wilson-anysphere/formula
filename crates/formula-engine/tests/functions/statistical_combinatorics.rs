use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn permut_matches_known_value() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=PERMUT(10,3)"), 720.0);
}

#[test]
fn permutationa_matches_known_value() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=PERMUTATIONA(3,4)"), 81.0);
}

#[test]
fn combinatorics_domain_errors_match_excel() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=PERMUT(3,4)"), Value::Error(ErrorKind::Num));
    assert_eq!(
        sheet.eval("=PERMUTATIONA(0,1)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(sheet.eval("=PERMUT(-1,1)"), Value::Error(ErrorKind::Num));
    assert_eq!(
        sheet.eval("=PERMUTATIONA(-1,1)"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn permutationa_supports_array_lift() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=PERMUTATIONA({2,3},{2,2})");
    sheet.recalc();
    assert_number(&sheet.get("A1"), 4.0);
    assert_number(&sheet.get("B1"), 9.0);
}
