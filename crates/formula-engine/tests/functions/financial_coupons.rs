use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::error::ExcelError;
use formula_engine::functions::date_time;
use formula_engine::functions::financial::{
    coupdaybs, coupdays, coupdaysnc, coupncd, coupnum, couppcd,
};

use super::harness::{assert_number, TestSheet};

#[test]
fn coup_functions_known_values_basis_0_semiannual() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2024, 4, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2025, 1, 1), system).unwrap();

    let expected_pcd = ymd_to_serial(ExcelDate::new(2024, 1, 1), system).unwrap();
    let expected_ncd = ymd_to_serial(ExcelDate::new(2024, 7, 1), system).unwrap();

    assert_eq!(
        couppcd(settlement, maturity, 2, 0, system).unwrap(),
        expected_pcd
    );
    assert_eq!(
        coupncd(settlement, maturity, 2, 0, system).unwrap(),
        expected_ncd
    );

    assert_eq!(coupdaybs(settlement, maturity, 2, 0, system).unwrap(), 90.0);
    assert_eq!(
        coupdaysnc(settlement, maturity, 2, 0, system).unwrap(),
        90.0
    );
    assert_eq!(coupdays(settlement, maturity, 2, 0, system).unwrap(), 180.0);
    assert_eq!(coupnum(settlement, maturity, 2, 0, system).unwrap(), 2.0);
}

#[test]
fn coup_functions_known_values_basis_1_quarterly() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2024, 2, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2024, 12, 15), system).unwrap();

    let expected_pcd = ymd_to_serial(ExcelDate::new(2023, 12, 15), system).unwrap();
    let expected_ncd = ymd_to_serial(ExcelDate::new(2024, 3, 15), system).unwrap();

    assert_eq!(
        couppcd(settlement, maturity, 4, 1, system).unwrap(),
        expected_pcd
    );
    assert_eq!(
        coupncd(settlement, maturity, 4, 1, system).unwrap(),
        expected_ncd
    );

    assert_eq!(coupdaybs(settlement, maturity, 4, 1, system).unwrap(), 48.0);
    assert_eq!(
        coupdaysnc(settlement, maturity, 4, 1, system).unwrap(),
        43.0
    );
    assert_eq!(coupdays(settlement, maturity, 4, 1, system).unwrap(), 91.0);
    assert_eq!(coupnum(settlement, maturity, 4, 1, system).unwrap(), 4.0);
}

#[test]
fn coup_days_consistency_properties() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2024, 2, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2024, 12, 15), system).unwrap();
    let basis = 1;

    let days = coupdays(settlement, maturity, 4, basis, system).unwrap();
    let daybs = coupdaybs(settlement, maturity, 4, basis, system).unwrap();
    let daysnc = coupdaysnc(settlement, maturity, 4, basis, system).unwrap();
    assert_eq!(days, daybs + daysnc);

    let settlement_on_coupon = couppcd(settlement, maturity, 4, basis, system).unwrap();
    assert_eq!(
        coupdaybs(settlement_on_coupon, maturity, 4, basis, system).unwrap(),
        0.0
    );
}

#[test]
fn coup_error_cases() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2024, 4, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2025, 1, 1), system).unwrap();

    assert_eq!(
        coupdaybs(settlement, maturity, 3, 0, system).unwrap_err(),
        ExcelError::Num
    );
    assert_eq!(
        coupdaybs(settlement, maturity, 2, 5, system).unwrap_err(),
        ExcelError::Num
    );
    assert_eq!(
        coupdaybs(maturity, maturity, 2, 0, system).unwrap_err(),
        ExcelError::Num
    );
    assert_eq!(
        coupdaybs(maturity, settlement, 2, 0, system).unwrap_err(),
        ExcelError::Num
    );
}

#[test]
fn builtins_support_date_strings_and_default_basis() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", "2024-04-01");
    sheet.set("A2", "2025-01-01");

    // Basis omitted -> defaults to 0.
    let v = sheet.eval("=COUPDAYBS(A1,A2,2)");
    assert_number(&v, 90.0);

    // Date-returning functions should also accept text and return serial numbers.
    let pcd = sheet.eval("=COUPPCD(A1,A2,2)");
    let expected_pcd =
        ymd_to_serial(ExcelDate::new(2024, 1, 1), ExcelDateSystem::EXCEL_1900).unwrap() as f64;
    assert_number(&pcd, expected_pcd);
}

#[test]
fn coup_functions_apply_end_of_month_schedule_when_maturity_is_month_end_basis_1() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Maturity at month-end but not the 31st: Excel pins coupon dates to month-end when maturity
    // is EOM. This affects basis=1 because COUPDAYS uses the actual day-count between coupon dates.
    //
    // Quarterly schedule, maturity=2020-04-30 => PCD=2020-01-31, NCD=2020-04-30.
    let settlement = ymd_to_serial(ExcelDate::new(2020, 2, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2020, 4, 30), system).unwrap();
    let expected_pcd = ymd_to_serial(ExcelDate::new(2020, 1, 31), system).unwrap();
    let expected_ncd = maturity;

    assert_eq!(
        couppcd(settlement, maturity, 4, 1, system).unwrap(),
        expected_pcd
    );
    assert_eq!(
        coupncd(settlement, maturity, 4, 1, system).unwrap(),
        expected_ncd
    );
    assert_eq!(coupnum(settlement, maturity, 4, 1, system).unwrap(), 1.0);

    assert_eq!(coupdaybs(settlement, maturity, 4, 1, system).unwrap(), 15.0);
    assert_eq!(
        coupdaysnc(settlement, maturity, 4, 1, system).unwrap(),
        75.0
    );
    assert_eq!(coupdays(settlement, maturity, 4, 1, system).unwrap(), 90.0);

    // Semiannual schedule, maturity=2021-02-28 => PCD=2020-08-31, NCD=2021-02-28.
    let settlement = ymd_to_serial(ExcelDate::new(2020, 11, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 2, 28), system).unwrap();
    let expected_pcd = ymd_to_serial(ExcelDate::new(2020, 8, 31), system).unwrap();
    let expected_ncd = maturity;

    assert_eq!(
        couppcd(settlement, maturity, 2, 1, system).unwrap(),
        expected_pcd
    );
    assert_eq!(
        coupncd(settlement, maturity, 2, 1, system).unwrap(),
        expected_ncd
    );
    assert_eq!(coupnum(settlement, maturity, 2, 1, system).unwrap(), 1.0);

    assert_eq!(coupdaybs(settlement, maturity, 2, 1, system).unwrap(), 76.0);
    assert_eq!(
        coupdaysnc(settlement, maturity, 2, 1, system).unwrap(),
        105.0
    );
    assert_eq!(coupdays(settlement, maturity, 2, 1, system).unwrap(), 181.0);
}

#[test]
fn coupdays_basis_4_uses_fixed_360_over_frequency_and_preserves_additivity() {
    let system = ExcelDateSystem::EXCEL_1900;

    // For basis=4 (European 30E/360):
    // - COUPDAYBS uses `DAYS360(PCD, settlement, TRUE)` for the day count A.
    // - COUPDAYS returns the modeled coupon period length E, which Excel treats as a fixed
    //   `360/frequency` (not `DAYS360(PCD, NCD, TRUE)`).
    // - COUPDAYSNC is computed as the remainder `E - A` to preserve the additivity invariant
    //   `COUPDAYBS + COUPDAYSNC == COUPDAYS` even though DAYS360 is not additive for some
    //   month-end schedules involving February.
    //
    // Semiannual schedule, maturity=2021-02-28 => PCD=2020-08-31, NCD=2021-02-28.
    // European `DAYS360(2020-08-31, 2021-02-28, TRUE) = 178` (not 180 = 360/frequency).
    let settlement = ymd_to_serial(ExcelDate::new(2020, 11, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 2, 28), system).unwrap();
    let expected_pcd = ymd_to_serial(ExcelDate::new(2020, 8, 31), system).unwrap();
    let expected_ncd = maturity;

    let days360_coupon =
        date_time::days360(expected_pcd, expected_ncd, true, system).unwrap() as f64;
    assert_eq!(days360_coupon, 178.0);

    assert_eq!(
        couppcd(settlement, maturity, 2, 4, system).unwrap(),
        expected_pcd
    );
    assert_eq!(
        coupncd(settlement, maturity, 2, 4, system).unwrap(),
        expected_ncd
    );
    assert_eq!(coupnum(settlement, maturity, 2, 4, system).unwrap(), 1.0);
    let expected_daybs = date_time::days360(expected_pcd, settlement, true, system).unwrap() as f64;
    assert_eq!(expected_daybs, 75.0);

    let expected_days = 360.0 / 2.0;
    assert_eq!(expected_days, 180.0);
    assert_ne!(expected_days, days360_coupon);
    let expected_daysnc = expected_days - expected_daybs;
    assert_eq!(expected_daysnc, 105.0);

    let days = coupdays(settlement, maturity, 2, 4, system).unwrap();
    let daybs = coupdaybs(settlement, maturity, 2, 4, system).unwrap();
    let daysnc = coupdaysnc(settlement, maturity, 2, 4, system).unwrap();

    assert_eq!(days, expected_days);
    assert_eq!(daybs, expected_daybs);
    assert_eq!(daysnc, expected_daysnc);
    assert_eq!(daysnc, days - daybs);

    // For basis=4, DSC is computed as `E - A` and is not always equal to DAYS360(settlement, NCD).
    let european_days_settlement_to_ncd =
        date_time::days360(settlement, expected_ncd, true, system).unwrap() as f64;
    assert_eq!(european_days_settlement_to_ncd, 103.0);
    assert_ne!(daysnc, european_days_settlement_to_ncd);
}

#[test]
fn coupdays_basis_4_quarterly_eom_february_diverges_from_european_days360() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Same idea as `coupdays_basis_4_uses_fixed_360_over_frequency_and_preserves_additivity`,
    // but for a quarterly schedule:
    //
    // Quarterly schedule, maturity=2021-02-28 => PCD=2020-11-30, NCD=2021-02-28.
    //
    // European `DAYS360(2020-11-30, 2021-02-28, TRUE) = 88` (not 90 = 360/frequency).
    let settlement = ymd_to_serial(ExcelDate::new(2020, 12, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 2, 28), system).unwrap();
    let expected_pcd = ymd_to_serial(ExcelDate::new(2020, 11, 30), system).unwrap();
    let expected_ncd = maturity;

    assert_eq!(
        couppcd(settlement, maturity, 4, 4, system).unwrap(),
        expected_pcd
    );
    assert_eq!(
        coupncd(settlement, maturity, 4, 4, system).unwrap(),
        expected_ncd
    );
    assert_eq!(coupnum(settlement, maturity, 4, 4, system).unwrap(), 1.0);

    assert_eq!(
        date_time::days360(expected_pcd, expected_ncd, true, system).unwrap() as f64,
        88.0
    );

    let expected_days = 360.0 / 4.0;
    let expected_daybs = date_time::days360(expected_pcd, settlement, true, system).unwrap() as f64;
    assert_eq!(expected_daybs, 15.0);

    let expected_daysnc = expected_days - expected_daybs;
    assert_eq!(expected_daysnc, 75.0);

    let days = coupdays(settlement, maturity, 4, 4, system).unwrap();
    let daybs = coupdaybs(settlement, maturity, 4, 4, system).unwrap();
    let daysnc = coupdaysnc(settlement, maturity, 4, 4, system).unwrap();

    assert_eq!(days, expected_days);
    assert_eq!(daybs, expected_daybs);
    assert_eq!(daysnc, expected_daysnc);
    assert_eq!(daysnc, days - daybs);

    // For basis=4, DSC is computed as `E - A` and is not always equal to DAYS360(settlement, NCD).
    let european_days_settlement_to_ncd =
        date_time::days360(settlement, expected_ncd, true, system).unwrap() as f64;
    assert_eq!(european_days_settlement_to_ncd, 73.0);
    assert_ne!(daysnc, european_days_settlement_to_ncd);
}

#[test]
fn coup_schedule_eom_maturity_clamps_previous_coupon_date() {
    let system = ExcelDateSystem::EXCEL_1900;
    // Maturity is EOM; stepping back 6 months must clamp (Mar 31 -> Sep 30).
    let settlement = ymd_to_serial(ExcelDate::new(2023, 10, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2024, 3, 31), system).unwrap();

    let expected_pcd = ymd_to_serial(ExcelDate::new(2023, 9, 30), system).unwrap();
    let expected_ncd = maturity;

    for basis in [0, 1, 2, 3, 4] {
        assert_eq!(
            couppcd(settlement, maturity, 2, basis, system).unwrap(),
            expected_pcd
        );
        assert_eq!(
            coupncd(settlement, maturity, 2, basis, system).unwrap(),
            expected_ncd
        );
        assert_eq!(
            coupnum(settlement, maturity, 2, basis, system).unwrap(),
            1.0
        );
    }

    // Day-count sanity across bases.
    let pcd_b1 = couppcd(settlement, maturity, 2, 1, system).unwrap();
    let ncd_b1 = coupncd(settlement, maturity, 2, 1, system).unwrap();
    let days_b1 = coupdays(settlement, maturity, 2, 1, system).unwrap();
    assert_eq!(days_b1, (ncd_b1 - pcd_b1) as f64);
    let daybs_b1 = coupdaybs(settlement, maturity, 2, 1, system).unwrap();
    let daysnc_b1 = coupdaysnc(settlement, maturity, 2, 1, system).unwrap();
    assert_eq!(days_b1, daybs_b1 + daysnc_b1);

    for basis in [0, 2] {
        assert_eq!(
            coupdays(settlement, maturity, 2, basis, system).unwrap(),
            360.0 / 2.0
        );
    }
    assert_eq!(coupdays(settlement, maturity, 2, 4, system).unwrap(), 180.0);
    assert_eq!(
        coupdays(settlement, maturity, 2, 3, system).unwrap(),
        365.0 / 2.0
    );
}

#[test]
fn coup_schedule_leap_day_clamps_previous_coupon_date() {
    let system = ExcelDateSystem::EXCEL_1900;
    // Maturity is EOM; stepping back 3 months from May 31 should clamp to leap-day (Feb 29).
    let settlement = ymd_to_serial(ExcelDate::new(2024, 3, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2024, 5, 31), system).unwrap();

    let expected_pcd = ymd_to_serial(ExcelDate::new(2024, 2, 29), system).unwrap();
    let expected_ncd = maturity;

    for basis in [0, 1, 2, 3, 4] {
        assert_eq!(
            couppcd(settlement, maturity, 4, basis, system).unwrap(),
            expected_pcd
        );
        assert_eq!(
            coupncd(settlement, maturity, 4, basis, system).unwrap(),
            expected_ncd
        );
        assert_eq!(
            coupnum(settlement, maturity, 4, basis, system).unwrap(),
            1.0
        );
    }

    // Day-count sanity across bases.
    let pcd_b1 = couppcd(settlement, maturity, 4, 1, system).unwrap();
    let ncd_b1 = coupncd(settlement, maturity, 4, 1, system).unwrap();
    let days_b1 = coupdays(settlement, maturity, 4, 1, system).unwrap();
    assert_eq!(days_b1, (ncd_b1 - pcd_b1) as f64);
    let daybs_b1 = coupdaybs(settlement, maturity, 4, 1, system).unwrap();
    let daysnc_b1 = coupdaysnc(settlement, maturity, 4, 1, system).unwrap();
    assert_eq!(days_b1, daybs_b1 + daysnc_b1);

    for basis in [0, 2] {
        assert_eq!(
            coupdays(settlement, maturity, 4, basis, system).unwrap(),
            360.0 / 4.0
        );
    }
    assert_eq!(coupdays(settlement, maturity, 4, 4, system).unwrap(), 90.0);
    assert_eq!(
        coupdays(settlement, maturity, 4, 3, system).unwrap(),
        365.0 / 4.0
    );
}

#[test]
fn coup_schedule_leap_day_maturity_clamps_back_to_february_month_end() {
    let system = ExcelDateSystem::EXCEL_1900;
    // Maturity itself is the leap-day month-end. The schedule should be treated as EOM and remain
    // pinned to month-end, clamping back to Feb 28 in the non-leap year when stepping back 12
    // months.
    //
    // Semiannual schedule anchored at maturity:
    //   ... 2023-02-28, 2023-08-31, 2024-02-29 (maturity)
    let settlement = ymd_to_serial(ExcelDate::new(2023, 3, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2024, 2, 29), system).unwrap();

    let expected_pcd = ymd_to_serial(ExcelDate::new(2023, 2, 28), system).unwrap();
    let expected_ncd = ymd_to_serial(ExcelDate::new(2023, 8, 31), system).unwrap();

    for basis in [0, 1, 2, 3, 4] {
        assert_eq!(
            couppcd(settlement, maturity, 2, basis, system).unwrap(),
            expected_pcd
        );
        assert_eq!(
            coupncd(settlement, maturity, 2, basis, system).unwrap(),
            expected_ncd
        );
        assert_eq!(
            coupnum(settlement, maturity, 2, basis, system).unwrap(),
            2.0
        );
    }

    // Day-count sanity across bases.
    let pcd_b1 = couppcd(settlement, maturity, 2, 1, system).unwrap();
    let ncd_b1 = coupncd(settlement, maturity, 2, 1, system).unwrap();
    let days_b1 = coupdays(settlement, maturity, 2, 1, system).unwrap();
    assert_eq!(days_b1, (ncd_b1 - pcd_b1) as f64);
    let daybs_b1 = coupdaybs(settlement, maturity, 2, 1, system).unwrap();
    let daysnc_b1 = coupdaysnc(settlement, maturity, 2, 1, system).unwrap();
    assert_eq!(days_b1, daybs_b1 + daysnc_b1);

    for basis in [0, 2, 4] {
        assert_eq!(
            coupdays(settlement, maturity, 2, basis, system).unwrap(),
            360.0 / 2.0
        );
    }
    assert_eq!(
        coupdays(settlement, maturity, 2, 3, system).unwrap(),
        365.0 / 2.0
    );
}

#[test]
fn builtins_coup_dates_handle_eom_and_leap_day_schedules() {
    let system = ExcelDateSystem::EXCEL_1900;
    let mut sheet = TestSheet::new();

    sheet.set("A1", "2023-10-01");
    sheet.set("A2", "2024-03-31");
    sheet.set("B1", "2024-03-01");
    sheet.set("B2", "2024-05-31");

    // EOM maturity clamping (Mar 31 -> Sep 30).
    let expected_pcd_a = ymd_to_serial(ExcelDate::new(2023, 9, 30), system).unwrap() as f64;
    let expected_ncd_a = ymd_to_serial(ExcelDate::new(2024, 3, 31), system).unwrap() as f64;
    for basis in [0, 1, 2, 3, 4] {
        assert_number(
            &sheet.eval(&format!("=COUPPCD(A1,A2,2,{basis})")),
            expected_pcd_a,
        );
        assert_number(
            &sheet.eval(&format!("=COUPNCD(A1,A2,2,{basis})")),
            expected_ncd_a,
        );
    }

    // Leap-day clamping (May 31 -> Feb 29).
    let expected_pcd_b = ymd_to_serial(ExcelDate::new(2024, 2, 29), system).unwrap() as f64;
    let expected_ncd_b = ymd_to_serial(ExcelDate::new(2024, 5, 31), system).unwrap() as f64;
    for basis in [0, 1, 2, 3, 4] {
        assert_number(
            &sheet.eval(&format!("=COUPPCD(B1,B2,4,{basis})")),
            expected_pcd_b,
        );
        assert_number(
            &sheet.eval(&format!("=COUPNCD(B1,B2,4,{basis})")),
            expected_ncd_b,
        );
    }
}

#[test]
fn builtins_coup_handle_leap_day_maturity_schedule() {
    let system = ExcelDateSystem::EXCEL_1900;
    let mut sheet = TestSheet::new();

    sheet.set("A1", "2023-03-01");
    sheet.set("A2", "2024-02-29");

    let expected_pcd = ymd_to_serial(ExcelDate::new(2023, 2, 28), system).unwrap() as f64;
    let expected_ncd = ymd_to_serial(ExcelDate::new(2023, 8, 31), system).unwrap() as f64;

    for basis in [0, 1, 2, 3, 4] {
        assert_number(
            &sheet.eval(&format!("=COUPPCD(A1,A2,2,{basis})")),
            expected_pcd,
        );
        assert_number(
            &sheet.eval(&format!("=COUPNCD(A1,A2,2,{basis})")),
            expected_ncd,
        );
        assert_number(&sheet.eval(&format!("=COUPNUM(A1,A2,2,{basis})")), 2.0);
    }

    // Day-count sanity for the bases that are most sensitive to month-end / leap-year behavior.
    assert_number(&sheet.eval("=COUPDAYBS(A1,A2,2,0)"), 1.0);
    assert_number(&sheet.eval("=COUPDAYSNC(A1,A2,2,0)"), 179.0);
    assert_number(&sheet.eval("=COUPDAYS(A1,A2,2,0)"), 180.0);

    assert_number(&sheet.eval("=COUPDAYBS(A1,A2,2,1)"), 1.0);
    assert_number(&sheet.eval("=COUPDAYSNC(A1,A2,2,1)"), 183.0);
    assert_number(&sheet.eval("=COUPDAYS(A1,A2,2,1)"), 184.0);

    assert_number(&sheet.eval("=COUPDAYBS(A1,A2,2,4)"), 3.0);
    assert_number(&sheet.eval("=COUPDAYSNC(A1,A2,2,4)"), 177.0);
    assert_number(&sheet.eval("=COUPDAYS(A1,A2,2,4)"), 180.0);
}

#[test]
fn builtins_coup_day_counts_handle_eom_and_leap_day_schedules() {
    let _system = ExcelDateSystem::EXCEL_1900;
    let mut sheet = TestSheet::new();

    sheet.set("A1", "2023-10-01");
    sheet.set("A2", "2024-03-31");
    sheet.set("B1", "2024-03-01");
    sheet.set("B2", "2024-05-31");

    // EOM maturity clamping (Mar 31 -> Sep 30).
    assert_number(&sheet.eval("=COUPDAYBS(A1,A2,2,0)"), 1.0);
    assert_number(&sheet.eval("=COUPDAYSNC(A1,A2,2,0)"), 179.0);
    assert_number(&sheet.eval("=COUPDAYS(A1,A2,2,0)"), 180.0);
    assert_number(&sheet.eval("=COUPNUM(A1,A2,2,0)"), 1.0);

    assert_number(&sheet.eval("=COUPDAYBS(A1,A2,2,1)"), 1.0);
    assert_number(&sheet.eval("=COUPDAYSNC(A1,A2,2,1)"), 182.0);
    assert_number(&sheet.eval("=COUPDAYS(A1,A2,2,1)"), 183.0);
    assert_number(&sheet.eval("=COUPNUM(A1,A2,2,1)"), 1.0);

    // COUPDAYS is basis-dependent for 2/3/4, but the coupon *dates* are not.
    assert_number(&sheet.eval("=COUPDAYS(A1,A2,2,2)"), 180.0);
    assert_number(&sheet.eval("=COUPDAYS(A1,A2,2,3)"), 182.5);
    assert_number(&sheet.eval("=COUPDAYS(A1,A2,2,4)"), 180.0);
    assert_number(&sheet.eval("=COUPDAYBS(A1,A2,2,4)"), 1.0);
    assert_number(&sheet.eval("=COUPDAYSNC(A1,A2,2,4)"), 179.0);
    assert_number(&sheet.eval("=COUPNUM(A1,A2,2,2)"), 1.0);
    assert_number(&sheet.eval("=COUPNUM(A1,A2,2,3)"), 1.0);
    assert_number(&sheet.eval("=COUPNUM(A1,A2,2,4)"), 1.0);

    // Leap-day clamping (May 31 -> Feb 29).
    assert_number(&sheet.eval("=COUPDAYBS(B1,B2,4,0)"), 1.0);
    assert_number(&sheet.eval("=COUPDAYSNC(B1,B2,4,0)"), 89.0);
    assert_number(&sheet.eval("=COUPDAYS(B1,B2,4,0)"), 90.0);
    assert_number(&sheet.eval("=COUPNUM(B1,B2,4,0)"), 1.0);

    assert_number(&sheet.eval("=COUPDAYBS(B1,B2,4,1)"), 1.0);
    assert_number(&sheet.eval("=COUPDAYSNC(B1,B2,4,1)"), 91.0);
    assert_number(&sheet.eval("=COUPDAYS(B1,B2,4,1)"), 92.0);
    assert_number(&sheet.eval("=COUPNUM(B1,B2,4,1)"), 1.0);

    assert_number(&sheet.eval("=COUPDAYS(B1,B2,4,2)"), 90.0);
    assert_number(&sheet.eval("=COUPDAYS(B1,B2,4,3)"), 91.25);
    assert_number(&sheet.eval("=COUPDAYS(B1,B2,4,4)"), 90.0);
    assert_number(&sheet.eval("=COUPDAYBS(B1,B2,4,4)"), 2.0);
    assert_number(&sheet.eval("=COUPDAYSNC(B1,B2,4,4)"), 88.0);
    assert_number(&sheet.eval("=COUPNUM(B1,B2,4,2)"), 1.0);
    assert_number(&sheet.eval("=COUPNUM(B1,B2,4,3)"), 1.0);
    assert_number(&sheet.eval("=COUPNUM(B1,B2,4,4)"), 1.0);
}
