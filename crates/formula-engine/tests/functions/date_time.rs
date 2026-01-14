use chrono::{TimeZone, Utc};

use formula_engine::coercion::ValueLocaleConfig;
use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::functions::date_time;
use formula_engine::ExcelError;

#[test]
fn time_builds_fractional_days() {
    assert!((date_time::time(1, 30, 0).unwrap() - 0.0625).abs() < 1.0e-12);
    assert_eq!(date_time::time(24, 0, 0).unwrap(), 1.0);
    assert_eq!(date_time::time(-1, 0, 0).unwrap_err(), ExcelError::Num);
}

#[test]
fn timevalue_parses_common_formats() {
    let cfg = ValueLocaleConfig::en_us();
    assert!((date_time::timevalue("1:30", cfg).unwrap() - 0.0625).abs() < 1.0e-12);
    assert!((date_time::timevalue("1:30 PM", cfg).unwrap() - 0.5625).abs() < 1.0e-12);
}

#[test]
fn datevalue_parses_iso_and_us_formats() {
    let system = ExcelDateSystem::EXCEL_1900;
    let now = Utc.with_ymd_and_hms(2024, 6, 1, 0, 0, 0).unwrap();
    let cfg = ValueLocaleConfig::en_us();
    let expected = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    assert_eq!(
        date_time::datevalue("2020-01-01", cfg, now, system).unwrap(),
        expected
    );
    assert_eq!(
        date_time::datevalue("1/1/2020", cfg, now, system).unwrap(),
        expected
    );
}

#[test]
fn datevalue_returns_value_error_for_invalid_dates() {
    let system = ExcelDateSystem::EXCEL_1900;
    let now = Utc.with_ymd_and_hms(2024, 6, 1, 0, 0, 0).unwrap();
    let cfg = ValueLocaleConfig::en_us();
    assert_eq!(
        date_time::datevalue("2019-02-29", cfg, now, system).unwrap_err(),
        ExcelError::Value
    );
}

#[test]
fn eomonth_returns_last_day_of_offset_month() {
    let system = ExcelDateSystem::EXCEL_1900;
    let start = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let jan_last = ymd_to_serial(ExcelDate::new(2020, 1, 31), system).unwrap();
    let feb_last = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();

    assert_eq!(date_time::eomonth(start, 0, system).unwrap(), jan_last);
    assert_eq!(date_time::eomonth(start, 1, system).unwrap(), feb_last);
}

#[test]
fn days360_matches_excel_examples() {
    let system = ExcelDateSystem::EXCEL_1900;
    let start = ymd_to_serial(ExcelDate::new(2011, 1, 1), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2011, 12, 31), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 360);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 359);

    let start = ymd_to_serial(ExcelDate::new(2020, 2, 1), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 30);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 28);

    // Leap-day and end-of-month edge cases: US/NASD and European methods diverge
    // based on how they treat February month-end and 31st-of-month dates.
    let start = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 3, 31), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 30);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 31);

    // In leap years, Feb 28 is *not* month-end, so US/NASD month-end rollover does not apply to
    // the start date.
    let start = ymd_to_serial(ExcelDate::new(2020, 2, 28), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 3, 31), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 33);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 32);

    // US/NASD-only behavior: end-of-month can roll forward to the 1st of next month when the
    // start day is < 30.
    let start = ymd_to_serial(ExcelDate::new(2020, 2, 28), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 3);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 1);

    let start = ymd_to_serial(ExcelDate::new(2020, 1, 31), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 30);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 29);

    // In leap years, Feb 28 is not month-end, so both methods treat it as day 28.
    let start = ymd_to_serial(ExcelDate::new(2020, 1, 31), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 2, 28), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 28);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 28);

    let start = ymd_to_serial(ExcelDate::new(2019, 1, 31), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2019, 2, 28), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 30);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 28);

    // Non-leap February month-end to 31st-of-month.
    let start = ymd_to_serial(ExcelDate::new(2019, 2, 28), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2019, 3, 31), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 30);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 32);

    // US/NASD-only behavior: if end_date is month-end and start_day < 30, end_date rolls to the
    // 1st of the next month (this includes end-of-February handling).
    let start = ymd_to_serial(ExcelDate::new(2019, 2, 15), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2019, 2, 28), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 16);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 13);

    // Cross-year February month-end behavior.
    let start = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2021, 2, 28), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 360);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 359);

    // Leap-year Feb 28 is not month-end, so US/NASD can roll the end-of-month forward.
    let start = ymd_to_serial(ExcelDate::new(2020, 2, 28), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2021, 2, 28), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 363);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 360);

    // Cross-year non-leap February month-end to leap-day month-end.
    let start = ymd_to_serial(ExcelDate::new(2019, 2, 28), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 360);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 361);

    // Cross-year non-leap February month-end to leap-year Feb 28 (not month-end).
    let start = ymd_to_serial(ExcelDate::new(2019, 2, 28), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 2, 28), system).unwrap();
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 358);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 360);
}

#[test]
fn days360_month_end_rollover_applies_outside_february() {
    let system = ExcelDateSystem::EXCEL_1900;

    // US/NASD-only behavior: if end_date is month-end and the (adjusted) start day is < 30,
    // end_date rolls forward to the 1st of the next month. This can differ from the European
    // method even in months that are not February.
    let start = ymd_to_serial(ExcelDate::new(2019, 4, 29), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2019, 4, 30), system).unwrap(); // month-end (30-day month)
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 2);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 1);

    let start = ymd_to_serial(ExcelDate::new(2019, 4, 29), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2019, 5, 31), system).unwrap(); // month-end (31-day month)
    assert_eq!(date_time::days360(start, end, false, system).unwrap(), 32);
    assert_eq!(date_time::days360(start, end, true, system).unwrap(), 31);
}

#[test]
fn days360_accounts_for_lotus_bug_feb_1900() {
    let system = ExcelDateSystem::EXCEL_1900;
    let jan31 = ymd_to_serial(ExcelDate::new(1900, 1, 31), system).unwrap();
    let feb28 = ymd_to_serial(ExcelDate::new(1900, 2, 28), system).unwrap();
    let feb29 = ymd_to_serial(ExcelDate::new(1900, 2, 29), system).unwrap();
    let mar1 = ymd_to_serial(ExcelDate::new(1900, 3, 1), system).unwrap();
    let mar31 = ymd_to_serial(ExcelDate::new(1900, 3, 31), system).unwrap();

    // Under the Lotus 1900 leap-year bug, Feb 29 1900 exists as serial 60.
    // That means Feb 28 1900 is not month-end, while Feb 29 1900 is.
    assert_eq!(date_time::days360(feb28, mar1, false, system).unwrap(), 3);
    assert_eq!(date_time::days360(feb28, mar1, true, system).unwrap(), 3);

    assert_eq!(date_time::days360(feb28, feb29, false, system).unwrap(), 3);
    assert_eq!(date_time::days360(feb28, feb29, true, system).unwrap(), 1);

    // Feb 29 is treated as month-end, so US/NASD adjusts it to day 30.
    assert_eq!(date_time::days360(feb29, mar1, false, system).unwrap(), 1);
    assert_eq!(date_time::days360(feb29, mar1, true, system).unwrap(), 2);

    assert_eq!(date_time::days360(jan31, feb29, false, system).unwrap(), 30);
    assert_eq!(date_time::days360(jan31, feb29, true, system).unwrap(), 29);

    assert_eq!(date_time::days360(feb29, mar31, false, system).unwrap(), 30);
    assert_eq!(date_time::days360(feb29, mar31, true, system).unwrap(), 31);

    // Feb 28 is not month-end in the Lotus-bug date system (because Feb 29 exists), so US/NASD
    // does not adjust the end date here.
    assert_eq!(date_time::days360(jan31, feb28, false, system).unwrap(), 28);
    assert_eq!(date_time::days360(jan31, feb28, true, system).unwrap(), 28);
}

#[test]
fn days360_respects_lotus_compat_flag_for_feb_1900() {
    let system = ExcelDateSystem::Excel1900 {
        lotus_compat: false,
    };
    let feb28 = ymd_to_serial(ExcelDate::new(1900, 2, 28), system).unwrap();
    let mar1 = ymd_to_serial(ExcelDate::new(1900, 3, 1), system).unwrap();

    // Without the Lotus bug, Feb 28 1900 is month-end (so the US/NASD method adjusts it to day 30).
    assert_eq!(date_time::days360(feb28, mar1, false, system).unwrap(), 1);
    assert_eq!(date_time::days360(feb28, mar1, true, system).unwrap(), 3);

    let jan31 = ymd_to_serial(ExcelDate::new(1900, 1, 31), system).unwrap();
    // Without the Lotus bug, Feb 28 is month-end so US/NASD adjusts it to day 30.
    assert_eq!(date_time::days360(jan31, feb28, false, system).unwrap(), 30);
    assert_eq!(date_time::days360(jan31, feb28, true, system).unwrap(), 28);
}

#[test]
fn yearfrac_respects_basis_conventions() {
    let system = ExcelDateSystem::EXCEL_1900;
    let start = ymd_to_serial(ExcelDate::new(2011, 1, 1), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2011, 12, 31), system).unwrap();
    assert!((date_time::yearfrac(start, end, 0, system).unwrap() - 1.0).abs() < 1e-12);
    assert!((date_time::yearfrac(start, end, 4, system).unwrap() - (359.0 / 360.0)).abs() < 1e-12);

    let start = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let actual_days = (i64::from(end) - i64::from(start)) as f64;
    assert!(
        (date_time::yearfrac(start, end, 2, system).unwrap() - (actual_days / 360.0)).abs() < 1e-12
    );
    assert!(
        (date_time::yearfrac(start, end, 3, system).unwrap() - (actual_days / 365.0)).abs() < 1e-12
    );

    // Basis 1 counts whole-year anniversaries with leap-day clamping.
    let start = ymd_to_serial(ExcelDate::new(2020, 3, 1), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2021, 3, 1), system).unwrap();
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - 1.0).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + 1.0).abs() < 1e-12);

    // Regression tests around Feb 29 / end-of-month clamping.
    let a = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();
    let b = ymd_to_serial(ExcelDate::new(2021, 2, 28), system).unwrap();
    let c = ymd_to_serial(ExcelDate::new(2021, 3, 1), system).unwrap();
    let d = ymd_to_serial(ExcelDate::new(2019, 2, 28), system).unwrap();
    let e = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();

    assert_eq!(date_time::yearfrac(a, a, 1, system).unwrap(), 0.0);

    let ab = date_time::yearfrac(a, b, 1, system).unwrap();
    assert!((ab - 1.0).abs() < 1e-12);
    assert!((ab + date_time::yearfrac(b, a, 1, system).unwrap()).abs() < 1e-12);
    assert!((0.0..=2.0).contains(&ab));

    let ac = date_time::yearfrac(a, c, 1, system).unwrap();
    let expected_ac = 1.0 + 1.0 / 365.0;
    assert!((ac - expected_ac).abs() < 1e-12);
    assert!((ac + date_time::yearfrac(c, a, 1, system).unwrap()).abs() < 1e-12);
    assert!((0.0..=2.0).contains(&ac));

    let de = date_time::yearfrac(d, e, 1, system).unwrap();
    let expected_de = 1.0 + 1.0 / 366.0;
    assert!((de - expected_de).abs() < 1e-12);
    assert!((de + date_time::yearfrac(e, d, 1, system).unwrap()).abs() < 1e-12);
    assert!((0.0..=2.0).contains(&de));

    // Multi-year leap-day clamping should still count whole-year anniversaries.
    let a2 = ymd_to_serial(ExcelDate::new(2022, 2, 28), system).unwrap();
    let aa2 = date_time::yearfrac(a, a2, 1, system).unwrap();
    assert!((aa2 - 2.0).abs() < 1e-12);
    assert!((aa2 + date_time::yearfrac(a2, a, 1, system).unwrap()).abs() < 1e-12);

    let a3 = ymd_to_serial(ExcelDate::new(2023, 2, 28), system).unwrap();
    let aa3 = date_time::yearfrac(a, a3, 1, system).unwrap();
    assert!((aa3 - 3.0).abs() < 1e-12);
    assert!((aa3 + date_time::yearfrac(a3, a, 1, system).unwrap()).abs() < 1e-12);

    let a4 = ymd_to_serial(ExcelDate::new(2024, 2, 29), system).unwrap();
    let aa4 = date_time::yearfrac(a, a4, 1, system).unwrap();
    assert!((aa4 - 4.0).abs() < 1e-12);
    assert!((aa4 + date_time::yearfrac(a4, a, 1, system).unwrap()).abs() < 1e-12);

    // Partial-year spans should use the actual year length between anniversaries (365 vs 366).
    let f = ymd_to_serial(ExcelDate::new(2019, 3, 1), system).unwrap();
    let g = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();
    let fg = date_time::yearfrac(f, g, 1, system).unwrap();
    let expected_fg = 365.0 / 366.0;
    assert!((fg - expected_fg).abs() < 1e-12);
    assert!((fg + date_time::yearfrac(g, f, 1, system).unwrap()).abs() < 1e-12);

    let h = ymd_to_serial(ExcelDate::new(2020, 3, 1), system).unwrap();
    let i = ymd_to_serial(ExcelDate::new(2021, 2, 28), system).unwrap();
    let hi = date_time::yearfrac(h, i, 1, system).unwrap();
    let expected_hi = 364.0 / 365.0;
    assert!((hi - expected_hi).abs() < 1e-12);
    assert!((hi + date_time::yearfrac(i, h, 1, system).unwrap()).abs() < 1e-12);

    // Short spans around leap day should use the correct anniversary denominator.
    let j = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();
    let k = ymd_to_serial(ExcelDate::new(2020, 3, 1), system).unwrap();
    let jk = date_time::yearfrac(j, k, 1, system).unwrap();
    let expected_jk = 1.0 / 365.0;
    assert!((jk - expected_jk).abs() < 1e-12);
    assert!((jk + date_time::yearfrac(k, j, 1, system).unwrap()).abs() < 1e-12);

    let l = ymd_to_serial(ExcelDate::new(2020, 2, 28), system).unwrap();
    let m = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap();
    let lm = date_time::yearfrac(l, m, 1, system).unwrap();
    let expected_lm = 1.0 / 366.0;
    assert!((lm - expected_lm).abs() < 1e-12);
    assert!((lm + date_time::yearfrac(m, l, 1, system).unwrap()).abs() < 1e-12);

    // Feb 28 is an anniversary boundary even when the interval crosses a leap day.
    let n = ymd_to_serial(ExcelDate::new(2020, 2, 28), system).unwrap();
    let o = ymd_to_serial(ExcelDate::new(2021, 2, 28), system).unwrap();
    let no = date_time::yearfrac(n, o, 1, system).unwrap();
    assert!((no - 1.0).abs() < 1e-12);
    assert!((no + date_time::yearfrac(o, n, 1, system).unwrap()).abs() < 1e-12);

    // One day short of the anniversary should use the 366-day denominator for this span.
    let p = ymd_to_serial(ExcelDate::new(2021, 2, 27), system).unwrap();
    let np = date_time::yearfrac(n, p, 1, system).unwrap();
    let expected_np = 365.0 / 366.0;
    assert!((np - expected_np).abs() < 1e-12);
    assert!((np + date_time::yearfrac(p, n, 1, system).unwrap()).abs() < 1e-12);

    assert_eq!(
        date_time::yearfrac(start, end, 9, system).unwrap_err(),
        ExcelError::Num
    );
}

#[test]
fn yearfrac_basis1_accounts_for_lotus_bug_feb_1900() {
    let system = ExcelDateSystem::EXCEL_1900;
    let jan1 = ymd_to_serial(ExcelDate::new(1900, 1, 1), system).unwrap();
    let dec31 = ymd_to_serial(ExcelDate::new(1900, 12, 31), system).unwrap();
    let feb28 = ymd_to_serial(ExcelDate::new(1900, 2, 28), system).unwrap();
    let feb29 = ymd_to_serial(ExcelDate::new(1900, 2, 29), system).unwrap();
    let mar1 = ymd_to_serial(ExcelDate::new(1900, 3, 1), system).unwrap();
    let feb28_1901 = ymd_to_serial(ExcelDate::new(1901, 2, 28), system).unwrap();

    // 1900 is treated as a leap year, so there are 366 days between 1900-01-01 and 1901-01-01.
    let jan_to_dec = date_time::yearfrac(jan1, dec31, 1, system).unwrap();
    assert!((jan_to_dec - (365.0 / 366.0)).abs() < 1e-12);

    let feb28_to_feb29 = date_time::yearfrac(feb28, feb29, 1, system).unwrap();
    assert!((feb28_to_feb29 - (1.0 / 366.0)).abs() < 1e-12);
    assert!((feb28_to_feb29 + date_time::yearfrac(feb29, feb28, 1, system).unwrap()).abs() < 1e-12);

    let feb28_to_mar1 = date_time::yearfrac(feb28, mar1, 1, system).unwrap();
    assert!((feb28_to_mar1 - (2.0 / 366.0)).abs() < 1e-12);
    assert!((feb28_to_mar1 + date_time::yearfrac(mar1, feb28, 1, system).unwrap()).abs() < 1e-12);

    // Year anniversaries from Feb 29 clamp to Feb 28 in non-leap years, yielding a 365-day span.
    let feb29_to_mar1 = date_time::yearfrac(feb29, mar1, 1, system).unwrap();
    assert!((feb29_to_mar1 - (1.0 / 365.0)).abs() < 1e-12);
    assert!((feb29_to_mar1 + date_time::yearfrac(mar1, feb29, 1, system).unwrap()).abs() < 1e-12);

    let feb29_to_feb28_1901 = date_time::yearfrac(feb29, feb28_1901, 1, system).unwrap();
    assert!((feb29_to_feb28_1901 - 1.0).abs() < 1e-12);
    assert!(
        (feb29_to_feb28_1901 + date_time::yearfrac(feb28_1901, feb29, 1, system).unwrap()).abs()
            < 1e-12
    );
}

#[test]
fn yearfrac_basis1_respects_lotus_compat_flag_for_1900() {
    let system = ExcelDateSystem::Excel1900 {
        lotus_compat: false,
    };
    let jan1 = ymd_to_serial(ExcelDate::new(1900, 1, 1), system).unwrap();
    let dec31 = ymd_to_serial(ExcelDate::new(1900, 12, 31), system).unwrap();

    // Without the Lotus bug, 1900 is not a leap year, so the denominator is 365 days.
    let jan_to_dec = date_time::yearfrac(jan1, dec31, 1, system).unwrap();
    assert!((jan_to_dec - (364.0 / 365.0)).abs() < 1e-12);
    assert!((jan_to_dec + date_time::yearfrac(dec31, jan1, 1, system).unwrap()).abs() < 1e-12);

    let feb28 = ymd_to_serial(ExcelDate::new(1900, 2, 28), system).unwrap();
    let mar1 = ymd_to_serial(ExcelDate::new(1900, 3, 1), system).unwrap();
    let frac = date_time::yearfrac(feb28, mar1, 1, system).unwrap();
    assert!((frac - (1.0 / 365.0)).abs() < 1e-12);
    assert!((frac + date_time::yearfrac(mar1, feb28, 1, system).unwrap()).abs() < 1e-12);
}

#[test]
fn yearfrac_basis1_uses_correct_year_length_within_year() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Non-leap year: denom = 365.
    let start = ymd_to_serial(ExcelDate::new(2019, 1, 1), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2019, 12, 31), system).unwrap();
    let expected = 364.0 / 365.0;
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - expected).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + expected).abs() < 1e-12);

    let start = ymd_to_serial(ExcelDate::new(2019, 2, 28), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2019, 12, 31), system).unwrap();
    let expected = 306.0 / 365.0;
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - expected).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + expected).abs() < 1e-12);

    // Leap year: denom = 366 (unless the anniversary is Feb 29 clamped, which is covered elsewhere).
    let start = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 12, 31), system).unwrap();
    let expected = 365.0 / 366.0;
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - expected).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + expected).abs() < 1e-12);

    let start = ymd_to_serial(ExcelDate::new(2020, 2, 28), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 12, 31), system).unwrap();
    let expected = 307.0 / 366.0;
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - expected).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + expected).abs() < 1e-12);

    // The anniversary denominator can span into the next year, so it can include (or exclude) a leap
    // day even if both dates are within the same calendar year.
    let start = ymd_to_serial(ExcelDate::new(2019, 3, 1), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2019, 12, 31), system).unwrap();
    let expected = 305.0 / 366.0;
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - expected).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + expected).abs() < 1e-12);

    let start = ymd_to_serial(ExcelDate::new(2020, 3, 1), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 12, 31), system).unwrap();
    let expected = 305.0 / 365.0;
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - expected).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + expected).abs() < 1e-12);

    // One-day spans across year-end should also pick up the correct denominator based on the
    // anniversary year length.
    let start = ymd_to_serial(ExcelDate::new(2019, 12, 31), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let expected = 1.0 / 366.0;
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - expected).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + expected).abs() < 1e-12);

    let start = ymd_to_serial(ExcelDate::new(2020, 12, 31), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2021, 1, 1), system).unwrap();
    let expected = 1.0 / 365.0;
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - expected).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + expected).abs() < 1e-12);

    // Near-anniversary behavior at year-end should use the correct denominator (365 vs 366) based
    // on the anniversary year span.
    let start = ymd_to_serial(ExcelDate::new(2019, 12, 31), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 12, 30), system).unwrap();
    let expected = 365.0 / 366.0;
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - expected).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + expected).abs() < 1e-12);

    let end = ymd_to_serial(ExcelDate::new(2020, 12, 31), system).unwrap();
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - 1.0).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + 1.0).abs() < 1e-12);

    let start = ymd_to_serial(ExcelDate::new(2020, 12, 31), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2021, 12, 30), system).unwrap();
    let expected = 364.0 / 365.0;
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - expected).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + expected).abs() < 1e-12);

    let end = ymd_to_serial(ExcelDate::new(2021, 12, 31), system).unwrap();
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - 1.0).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + 1.0).abs() < 1e-12);

    // Feb 28 is month-end in non-leap years. When crossing into a leap year, the anniversary
    // denominator can still be 365 if the leap day occurs *after* the anniversary date.
    let start = ymd_to_serial(ExcelDate::new(2019, 2, 28), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 2, 27), system).unwrap();
    let expected = 364.0 / 365.0;
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - expected).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + expected).abs() < 1e-12);

    // Conversely, for a start date after Feb 29, the leap day can be inside the anniversary
    // denominator even if both dates are before year-end.
    let start = ymd_to_serial(ExcelDate::new(2019, 3, 1), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 2, 28), system).unwrap();
    let expected = 364.0 / 366.0;
    assert!((date_time::yearfrac(start, end, 1, system).unwrap() - expected).abs() < 1e-12);
    assert!((date_time::yearfrac(end, start, 1, system).unwrap() + expected).abs() < 1e-12);
}

#[test]
fn datedif_matches_excel_units() {
    let system = ExcelDateSystem::EXCEL_1900;
    let start = ymd_to_serial(ExcelDate::new(2011, 1, 15), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2012, 1, 14), system).unwrap();

    assert_eq!(date_time::datedif(start, end, "Y", system).unwrap(), 0);
    assert_eq!(date_time::datedif(start, end, "M", system).unwrap(), 11);
    assert_eq!(date_time::datedif(start, end, "D", system).unwrap(), 364);
    assert_eq!(date_time::datedif(start, end, "YM", system).unwrap(), 11);
    assert_eq!(date_time::datedif(start, end, "MD", system).unwrap(), 30);
    assert_eq!(date_time::datedif(start, end, "YD", system).unwrap(), 364);

    assert_eq!(
        date_time::datedif(start, end, "  ym  ", system).unwrap(),
        11
    );

    assert_eq!(
        date_time::datedif(end, start, "D", system).unwrap_err(),
        ExcelError::Num
    );
    assert_eq!(
        date_time::datedif(start, end, "NOPE", system).unwrap_err(),
        ExcelError::Num
    );
}

#[test]
fn weekday_matches_excel_return_types() {
    let system = ExcelDateSystem::EXCEL_1900;
    // 1900-01-01 is serial 1 and a Monday.
    assert_eq!(date_time::weekday(1, None, system).unwrap(), 2);
    assert_eq!(date_time::weekday(1, Some(2), system).unwrap(), 1);
    assert_eq!(date_time::weekday(1, Some(3), system).unwrap(), 0);
    assert_eq!(date_time::weekday(1, Some(11), system).unwrap(), 1);
    assert_eq!(date_time::weekday(1, Some(17), system).unwrap(), 2);
}

#[test]
fn workday_and_networkdays_skip_weekends_and_holidays() {
    let system = ExcelDateSystem::EXCEL_1900;
    let start = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap(); // Wednesday
    let jan2 = ymd_to_serial(ExcelDate::new(2020, 1, 2), system).unwrap();
    let jan3 = ymd_to_serial(ExcelDate::new(2020, 1, 3), system).unwrap();
    let jan6 = ymd_to_serial(ExcelDate::new(2020, 1, 6), system).unwrap();

    assert_eq!(date_time::workday(start, 1, None, system).unwrap(), jan2);
    assert_eq!(date_time::workday(start, 3, None, system).unwrap(), jan6);

    assert_eq!(
        date_time::workday(start, 1, Some(&[jan2]), system).unwrap(),
        jan3
    );

    let end = ymd_to_serial(ExcelDate::new(2020, 1, 10), system).unwrap();
    // Working days from Jan 1 to Jan 10 2020 inclusive: 8 (Wed-Fri + Mon-Fri).
    assert_eq!(date_time::networkdays(start, end, None, system).unwrap(), 8);
}
