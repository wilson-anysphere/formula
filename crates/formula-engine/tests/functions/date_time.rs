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
}

#[test]
fn yearfrac_respects_basis_conventions() {
    let system = ExcelDateSystem::EXCEL_1900;
    let start = ymd_to_serial(ExcelDate::new(2011, 1, 1), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2011, 12, 31), system).unwrap();
    assert!((date_time::yearfrac(start, end, 0, system).unwrap() - 1.0).abs() < 1e-12);
    assert!(
        (date_time::yearfrac(start, end, 4, system).unwrap() - (359.0 / 360.0)).abs() < 1e-12
    );

    let start = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let end = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let actual_days = (i64::from(end) - i64::from(start)) as f64;
    assert!(
        (date_time::yearfrac(start, end, 2, system).unwrap() - (actual_days / 360.0)).abs()
            < 1e-12
    );
    assert!(
        (date_time::yearfrac(start, end, 3, system).unwrap() - (actual_days / 365.0)).abs()
            < 1e-12
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

    assert_eq!(date_time::yearfrac(start, end, 9, system).unwrap_err(), ExcelError::Num);
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
