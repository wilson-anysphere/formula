use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn date_month_overflow_underflow() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=YEAR(DATE(2024,0,15))"), 2023.0);
    assert_number(&sheet.eval("=MONTH(DATE(2024,0,15))"), 12.0);
    assert_number(&sheet.eval("=DAY(DATE(2024,0,15))"), 15.0);

    assert_number(&sheet.eval("=YEAR(DATE(2024,13,1))"), 2025.0);
    assert_number(&sheet.eval("=MONTH(DATE(2024,13,1))"), 1.0);
    assert_number(&sheet.eval("=DAY(DATE(2024,13,1))"), 1.0);
}

#[test]
fn date_day_overflow_and_1900_leap_bug() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=MONTH(DATE(2024,1,32))"), 2.0);
    assert_number(&sheet.eval("=DAY(DATE(2024,1,32))"), 1.0);

    // Excel's Lotus 1-2-3 compatibility bug: DATE(1900,3,0) is Feb 29 1900.
    assert_number(&sheet.eval("=YEAR(DATE(1900,3,0))"), 1900.0);
    assert_number(&sheet.eval("=MONTH(DATE(1900,3,0))"), 2.0);
    assert_number(&sheet.eval("=DAY(DATE(1900,3,0))"), 29.0);
}

#[test]
fn today_and_now_are_volatile_and_consistent() {
    let mut sheet = TestSheet::new();
    let today = sheet.eval("=TODAY()");
    let now = sheet.eval("=NOW()");
    match (today, now) {
        (Value::Number(t), Value::Number(n)) => {
            assert!(n >= t);
            assert!(n < t + 1.0);
        }
        other => panic!("unexpected results: {other:?}"),
    }

    assert_eq!(sheet.eval("=INT(NOW())"), sheet.eval("=TODAY()"));
}

#[test]
fn year_month_day_errors_on_invalid_inputs() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=YEAR(#REF!)"), Value::Error(ErrorKind::Ref));
}

