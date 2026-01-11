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
    assert!((date_time::timevalue("1:30").unwrap() - 0.0625).abs() < 1.0e-12);
    assert!((date_time::timevalue("1:30 PM").unwrap() - 0.5625).abs() < 1.0e-12);
}

#[test]
fn datevalue_parses_iso_and_us_formats() {
    let system = ExcelDateSystem::EXCEL_1900;
    let expected = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    assert_eq!(date_time::datevalue("2020-01-01", system).unwrap(), expected);
    assert_eq!(date_time::datevalue("1/1/2020", system).unwrap(), expected);
}

#[test]
fn datevalue_returns_value_error_for_invalid_dates() {
    let system = ExcelDateSystem::EXCEL_1900;
    assert_eq!(
        date_time::datevalue("2019-02-29", system).unwrap_err(),
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
