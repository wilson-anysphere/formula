use super::harness::{assert_number, TestSheet};

#[test]
fn coup_schedule_semiannual_eom_matches_excel_example() {
    let mut sheet = TestSheet::new();

    // Example from Task: EOM maturity must keep coupon dates pinned to month-end.
    // Naïve `EDATE` stepping from 2021-02-28 would drift to 2021-08-28.
    assert_eq!(
        sheet.eval("=COUPPCD(DATE(2021,3,1),DATE(2021,8,31),2,0)"),
        sheet.eval("=DATE(2021,2,28)")
    );
    assert_eq!(
        sheet.eval("=COUPNCD(DATE(2021,3,1),DATE(2021,8,31),2,0)"),
        sheet.eval("=DATE(2021,8,31)")
    );

    assert_number(
        &sheet.eval("=DAY(COUPNCD(DATE(2021,3,1),DATE(2021,8,31),2,0))"),
        31.0,
    );
}

#[test]
fn coup_schedule_semiannual_eom_does_not_drift_across_year_boundary() {
    let mut sheet = TestSheet::new();

    // Settlement earlier in the schedule requires stepping through February and back to August.
    // Excel's EOM schedule should restore month-end (Aug 31), not drift to Aug 28.
    assert_eq!(
        sheet.eval("=COUPPCD(DATE(2020,12,1),DATE(2021,8,31),2,0)"),
        sheet.eval("=DATE(2020,8,31)")
    );
    assert_eq!(
        sheet.eval("=COUPNCD(DATE(2020,12,1),DATE(2021,8,31),2,0)"),
        sheet.eval("=DATE(2021,2,28)")
    );
}

#[test]
fn coup_schedule_eom_maturity_february_restores_later_month_end() {
    let mut sheet = TestSheet::new();

    // If maturity is the last day of February (28th in a non-leap year), Excel treats the entire
    // schedule as end-of-month and restores month-end in months with 30/31 days.
    //
    // Semiannual schedule: ... 2030-08-31, 2031-02-28 (maturity)
    assert_eq!(
        sheet.eval("=COUPPCD(DATE(2030,9,1),DATE(2031,2,28),2,0)"),
        sheet.eval("=DATE(2030,8,31)")
    );
    assert_eq!(
        sheet.eval("=COUPNCD(DATE(2030,9,1),DATE(2031,2,28),2,0)"),
        sheet.eval("=DATE(2031,2,28)")
    );

    // Regression assertion: should be Aug 31, not Aug 28.
    assert_number(
        &sheet.eval("=DAY(COUPPCD(DATE(2030,9,1),DATE(2031,2,28),2,0))"),
        31.0,
    );
}

#[test]
fn coup_schedule_quarterly_eom_restores_month_end_after_february() {
    let mut sheet = TestSheet::new();

    // Quarterly schedule:
    // ... 2020-11-30, 2021-02-28, 2021-05-31, 2021-08-31 (maturity)
    assert_eq!(
        sheet.eval("=COUPPCD(DATE(2021,3,1),DATE(2021,8,31),4,0)"),
        sheet.eval("=DATE(2021,2,28)")
    );
    assert_eq!(
        sheet.eval("=COUPNCD(DATE(2021,3,1),DATE(2021,8,31),4,0)"),
        sheet.eval("=DATE(2021,5,31)")
    );

    // Regression assertion: should be month-end, not May 28.
    assert_number(
        &sheet.eval("=DAY(COUPNCD(DATE(2021,3,1),DATE(2021,8,31),4,0))"),
        31.0,
    );
}

#[test]
fn coup_schedule_quarterly_eom_maturity_february_restores_november_month_end() {
    let mut sheet = TestSheet::new();

    // Quarterly schedule anchored at a February month-end maturity:
    // ... 2020-11-30, 2021-02-28 (maturity)
    assert_eq!(
        sheet.eval("=COUPPCD(DATE(2020,12,15),DATE(2021,2,28),4,0)"),
        sheet.eval("=DATE(2020,11,30)")
    );
    assert_eq!(
        sheet.eval("=COUPNCD(DATE(2020,12,15),DATE(2021,2,28),4,0)"),
        sheet.eval("=DATE(2021,2,28)")
    );
    assert_number(
        &sheet.eval("=DAY(COUPPCD(DATE(2020,12,15),DATE(2021,2,28),4,0))"),
        30.0,
    );
}

#[test]
fn coup_schedule_non_eom_maturity_does_not_pin_to_month_end() {
    let mut sheet = TestSheet::new();

    // If maturity is not the last day of the month, Excel does not apply EOM pinning. Coupon dates
    // preserve the maturity day-of-month when possible (here: 30th), instead of being forced to
    // month-end.
    assert_eq!(
        sheet.eval("=COUPPCD(DATE(2020,12,1),DATE(2021,8,30),2,0)"),
        sheet.eval("=DATE(2020,8,30)")
    );
    assert_eq!(
        sheet.eval("=COUPNCD(DATE(2020,12,1),DATE(2021,8,30),2,0)"),
        sheet.eval("=DATE(2021,2,28)")
    );
}
