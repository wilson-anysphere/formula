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

    // Zero rate is allowed and should return 0 (matches Excel behavior across other bond funcs).
    assert_number(
        &sheet.eval("=ACCRINTM(DATE(2020,1,1),DATE(2020,7,1),0,1000,0)"),
        0.0,
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
    // `calc_method` omitted defaults to FALSE (accrue from issue).
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0.1,1000,2,0)"),
        expected,
    );
}

#[test]
fn accrint_calc_method_is_ignored_after_first_interest() {
    let mut sheet = TestSheet::new();

    // Settlement is after the first interest date.
    // Excel accrues from the previous coupon date (PCD) regardless of calc_method.
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,8,15),0.1,1000,2,0)"),
        25.0,
    );
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,8,15),0.1,1000,2,0,TRUE)"),
        25.0,
    );
}

#[test]
fn accrint_supports_basis_variants() {
    let mut sheet = TestSheet::new();

    // Basis 1: actual/actual day counting within the coupon period.
    let expected_basis1 = 50.0 * (123.0 / 184.0);
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,9,15),0.1,1000,2,1,FALSE)"),
        expected_basis1,
    );

    // Basis 2: Actual/360 uses actual days in numerator with a fixed 360/f denominator.
    let expected_basis2 = 50.0 * (123.0 / 180.0);
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,9,15),0.1,1000,2,2,FALSE)"),
        expected_basis2,
    );

    // Basis 3: Actual/365 uses actual days in numerator with a fixed 365/f denominator.
    let expected_basis3 = 50.0 * (123.0 / (365.0 / 2.0));
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,9,15),0.1,1000,2,3,FALSE)"),
        expected_basis3,
    );

    // Basis 4: European 30/360 differs from US 30/360 around month ends.
    let expected_basis4 = 30.0 * (28.0 / 90.0);
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,1,30),DATE(2020,4,30),DATE(2020,2,28),0.12,1000,4,4)"),
        expected_basis4,
    );
}

#[test]
fn accrint_allows_zero_rate() {
    let mut sheet = TestSheet::new();

    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0,1000,2,0)"),
        0.0,
    );
}

#[test]
fn accrint_coupon_schedule_is_anchored_to_first_interest_to_avoid_edate_drift() {
    let mut sheet = TestSheet::new();

    // When stepping coupon dates forward from an end-of-month date, naive iterative EDATE stepping
    // can "drift" the day-of-month (e.g. Jan 31 + 3 months = Apr 30, then +3 months = Jul 30).
    //
    // Excel's COUP* schedule functions avoid this by anchoring each coupon date as an offset from the
    // anchor date (maturity). ACCRINT should do the same when anchored at first_interest.
    //
    // first_interest=2020-01-31, frequency=4 => months=3.
    // Coupon dates anchored at first_interest are: 2020-01-31, 2020-04-30, 2020-07-31, 2020-10-31, ...
    //
    // Settlement 2020-08-15 is in the period 2020-07-31..2020-10-31.
    // Coupon payment C = par * rate / f = 1000 * 0.12 / 4 = 30.
    // Basis 1 => A = 15 days, E = 92 days.
    let expected = 30.0 * (15.0 / 92.0);
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2019,12,15),DATE(2020,1,31),DATE(2020,8,15),0.12,1000,4,1)"),
        expected,
    );
}

#[test]
fn accrint_coupon_schedule_pins_month_end_when_first_interest_is_month_end() {
    let mut sheet = TestSheet::new();

    // Similar to the drift regression above, but with an anchor date that is month-end but not the
    // 31st (April 30). Excel's coupon schedule rules treat this as an end-of-month schedule and pin
    // subsequent coupon dates to month-end (Jul 31, Oct 31, ...), which affects basis=1 computations.

    // Settlement after first interest date => accrue from PCD regardless of calc_method.
    // first_interest=2020-04-30, frequency=4 => coupon dates: 2020-04-30, 2020-07-31, 2020-10-31, ...
    // Settlement 2020-08-15 is in the period 2020-07-31..2020-10-31.
    // Coupon payment C = par * rate / f = 1000 * 0.12 / 4 = 30.
    // Basis 1 => A = 15 days, E = 92 days.
    let expected = 30.0 * (15.0 / 92.0);
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2019,12,15),DATE(2020,4,30),DATE(2020,8,15),0.12,1000,4,1)"),
        expected,
    );

    // Settlement before first interest date => calc_method controls whether we accrue from issue
    // or from the start of the regular coupon period (PCD). Under the EOM schedule:
    // PCD = 2020-01-31, E = 90.
    let from_issue = 30.0 * (31.0 / 90.0);
    let from_pcd = 30.0 * (15.0 / 90.0);
    assert_number(
        &sheet
            .eval("=ACCRINT(DATE(2020,1,15),DATE(2020,4,30),DATE(2020,2,15),0.12,1000,4,1,FALSE)"),
        from_issue,
    );
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,1,15),DATE(2020,4,30),DATE(2020,2,15),0.12,1000,4,1,TRUE)"),
        from_pcd,
    );
}

#[test]
fn accrint_coupon_schedule_restores_month_end_when_first_interest_is_eom_february() {
    let mut sheet = TestSheet::new();
    // If `first_interest` is the last day of February (28th in a non-leap year), Excel treats the
    // coupon schedule as end-of-month and restores month-end in later months (e.g. Aug 31 rather
    // than Aug 28).
    //
    // Settlement is after first interest, so accrual should be from PCD=first_interest to
    // settlement with E = days(first_interest, next_coupon) under basis 1.
    let expected = 60.0 * 15.0 / 184.0; // C=1000*0.12/2, A=15 days (Feb 28 -> Mar 15), E=184 days (Feb 28 -> Aug 31)
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,12,15),DATE(2021,2,28),DATE(2021,3,15),0.12,1000,2,1)"),
        expected,
    );
}

#[test]
fn accrint_calc_method_affects_only_stub_period_before_first_interest() {
    let mut sheet = TestSheet::new();

    // Settlement is before first interest date (still in the issue stub period).
    // Under calc_method=FALSE, accrue from issue; under calc_method=TRUE, accrue from the start of the
    // regular coupon period (PCD).
    //
    // For this schedule: issue=2020-02-15, first_interest=2020-05-15, frequency=2:
    // PCD = 2019-11-15, E = 180 (30/360), C = 50.
    //
    // A(issue->settlement) = 60, A(PCD->settlement) = 150.
    let from_issue = 50.0 * (60.0 / 180.0);
    let from_pcd = 50.0 * (150.0 / 180.0);

    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0.1,1000,2,0,FALSE)"),
        from_issue,
    );
    assert_number(
        &sheet.eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,4,15),0.1,1000,2,0,TRUE)"),
        from_pcd,
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
