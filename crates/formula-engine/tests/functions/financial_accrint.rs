use formula_engine::value::ErrorKind;
use formula_engine::Value;

use super::harness::{assert_number, TestSheet};

#[test]
fn accrintm_matches_yearfrac_based_interest() {
    let mut sheet = TestSheet::new();

    // 30/360 basis: exactly half a year.
    assert_number(
        &sheet.eval("=ACCRINTM(DATE(2020,1,1),DATE(2020,7,1),0.1,1000,0)"),
        50.0,
    );

    // Actual/Actual basis in a leap year: 182/366 of a year.
    let expected = 1000.0 * 0.1 * (182.0 / 366.0);
    assert_number(
        &sheet.eval("=ACCRINTM(DATE(2020,1,1),DATE(2020,7,1),0.1,1000,1)"),
        expected,
    );
}

#[test]
fn accrint_settlement_before_first_interest_short_first_coupon() {
    let mut sheet = TestSheet::new();

    // Issue date is inside a coupon period (short first coupon).
    // Coupon schedule anchored on first_interest=2020-05-15 with frequency=2 => quasi coupon dates
    // are 2019-11-15, 2020-05-15, 2020-11-15, ...
    //
    // Under 30/360, days(issue->settlement)=60 and days(period)=180, coupon=50.
    // Interest = 50 * 60/180 = 16.666666...
    let expected = 50.0 * (60.0 / 180.0);
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0.1,1000,2,0)"),
        expected,
    );
}

#[test]
fn accrint_calc_method_true_vs_false_differs_after_first_interest() {
    let mut sheet = TestSheet::new();

    // Settlement is after the first interest date.
    //
    // calc_method omitted (default TRUE): counts interest from issue -> settlement,
    // spanning two coupon periods (short first period + partial next period).
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,8,15),0.1,1000,2,0)"),
        50.0,
    );

    // calc_method FALSE: counts interest only from the last coupon date (2020-05-15) -> settlement.
    assert_number(
        &sheet.eval(
            "=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,8,15),0.1,1000,2,0,FALSE)",
        ),
        25.0,
    );
}

#[test]
fn accrint_supports_basis_variants() {
    let mut sheet = TestSheet::new();

    // Basis 1: actual/actual day counting within the coupon period.
    let expected_basis1 = 50.0 * (123.0 / 184.0);
    assert_number(
        &sheet.eval(
            "=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,9,15),0.1,1000,2,1,FALSE)",
        ),
        expected_basis1,
    );

    // Basis 4: European 30/360 differs from US 30/360 around month ends.
    let expected_basis4 = 30.0 * (28.0 / 90.0);
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,1,30),DATE(2020,4,30),DATE(2020,2,28),0.12,1000,4,4)"),
        expected_basis4,
    );
}

#[test]
fn accrint_errors_on_invalid_inputs() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=ACCRINTM(DATE(2020,1,1),DATE(2020,7,1),0.1,1000,5)"),
        Value::Error(ErrorKind::Num)
    );

    // Invalid frequency.
    assert_eq!(
        sheet.eval("=ACCRINT(DATE(2020,1,1),DATE(2020,7,1),DATE(2020,3,1),0.1,1000,3)"),
        Value::Error(ErrorKind::Num)
    );

    // issue >= settlement.
    assert_eq!(
        sheet.eval("=ACCRINTM(DATE(2020,1,1),DATE(2020,1,1),0.1,1000,0)"),
        Value::Error(ErrorKind::Num)
    );
}

