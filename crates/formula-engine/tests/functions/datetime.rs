use formula_engine::date::ExcelDateSystem;
use formula_engine::locale::ValueLocaleConfig;
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

#[test]
fn respects_excel_1904_date_system_for_date_serials() {
    let mut sheet = TestSheet::new();
    sheet.set_date_system(ExcelDateSystem::Excel1904);
    assert_number(&sheet.eval("=DATE(1904,1,1)"), 0.0);
}

#[test]
fn lotus_bug_serial_60_maps_to_feb_29_1900() {
    let mut sheet = TestSheet::new();
    sheet.set_date_system(ExcelDateSystem::Excel1900 { lotus_compat: true });
    assert_number(&sheet.eval("=YEAR(60)"), 1900.0);
    assert_number(&sheet.eval("=MONTH(60)"), 2.0);
    assert_number(&sheet.eval("=DAY(60)"), 29.0);
}

#[test]
fn lotus_bug_can_be_disabled() {
    let mut sheet = TestSheet::new();
    sheet.set_date_system(ExcelDateSystem::Excel1900 {
        lotus_compat: false,
    });
    assert_number(&sheet.eval("=YEAR(60)"), 1900.0);
    assert_number(&sheet.eval("=MONTH(60)"), 3.0);
    assert_number(&sheet.eval("=DAY(60)"), 1.0);
}

#[test]
fn time_and_timevalue_conversions() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=TIME(1,30,0)"), 0.0625);
    assert_number(&sheet.eval("=TIME(24,0,0)"), 1.0);
    assert_eq!(sheet.eval("=TIME(-1,0,0)"), Value::Error(ErrorKind::Num));
    sheet.set("A1", Value::Number(f64::INFINITY));
    assert_eq!(sheet.eval("=TIME(A1,0,0)"), Value::Error(ErrorKind::Num));

    assert_number(&sheet.eval("=TIMEVALUE(\"1:30\")"), 0.0625);
    assert_number(&sheet.eval("=TIMEVALUE(\"1:30 PM\")"), 0.5625);
    assert_number(&sheet.eval("=TIMEVALUE(\"1 PM\")"), 13.0 / 24.0);
    assert_number(
        &sheet.eval("=TIMEVALUE(\"2020-01-01 1:30 PM\")"),
        13.5 / 24.0,
    );
    assert_eq!(
        sheet.eval("=TIMEVALUE(\"nope\")"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn datevalue_edate_and_eomonth() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=DATEVALUE(\"2020-01-01\")"),
        sheet.eval("=DATE(2020,1,1)")
    );
    assert_eq!(
        sheet.eval("=DATEVALUE(\"1/2/2020\")"),
        sheet.eval("=DATE(2020,1,2)")
    );
    assert_eq!(
        sheet.eval("=DATEVALUE(\"January 2, 2020\")"),
        sheet.eval("=DATE(2020,1,2)")
    );
    assert_eq!(
        sheet.eval("=DATEVALUE(\"nope\")"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=DATEVALUE(\"2019-02-29\")"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=DATEVALUE(\"2020-02-30\")"),
        Value::Error(ErrorKind::Value)
    );

    assert_number(&sheet.eval("=DAY(EOMONTH(DATE(2020,1,15),0))"), 31.0);
    assert_number(&sheet.eval("=DAY(EOMONTH(DATE(2020,1,15),1))"), 29.0);

    assert_number(&sheet.eval("=MONTH(EDATE(DATE(2020,1,31),1))"), 2.0);
    assert_number(&sheet.eval("=DAY(EDATE(DATE(2020,1,31),1))"), 29.0);
}

#[test]
fn value_parses_datetime_text() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=VALUE(\"2020-01-01 1:30 PM\")"),
        sheet.eval("=DATEVALUE(\"2020-01-01\")+TIMEVALUE(\"1:30 PM\")")
    );
}

#[test]
fn value_locale_controls_numeric_and_date_order_parsing() {
    let mut sheet = TestSheet::new();
    sheet.set_value_locale(ValueLocaleConfig::de_de());

    assert_eq!(
        sheet.eval("=DATEVALUE(\"1/2/2020\")"),
        sheet.eval("=DATE(2020,2,1)")
    );
    assert_eq!(
        sheet.eval("=VALUE(\"2020.01.01\")"),
        sheet.eval("=DATE(2020,1,1)")
    );
    assert_eq!(
        sheet.eval("=VALUE(\"1.2.2020\")"),
        sheet.eval("=DATE(2020,2,1)")
    );
    assert_eq!(
        sheet.eval("=VALUE(\"1.2\")=DATE(YEAR(TODAY()),2,1)"),
        Value::Bool(true)
    );
    assert_number(&sheet.eval("=VALUE(\"1,5\")"), 1.5);
}

#[test]
fn days_returns_day_difference() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=DAYS(DATE(2020,1,2),DATE(2020,1,1))"), 1.0);
    assert_number(&sheet.eval("=DAYS(\"2020-01-02\",\"2020-01-01\")"), 1.0);
    assert_eq!(
        sheet.eval("=DAYS(\"nope\",DATE(2020,1,1))"),
        Value::Error(ErrorKind::Value)
    );

    sheet.set("A1", Value::Number(f64::INFINITY));
    assert_eq!(sheet.eval("=DAYS(A1,0)"), Value::Error(ErrorKind::Num));
}

#[test]
fn days_spills_over_array_inputs() {
    let mut sheet = TestSheet::new();
    sheet.set_formula(
        "A1",
        "=DAYS({\"2020-01-02\";\"2020-01-03\"},\"2020-01-01\")",
    );
    sheet.recalc();
    assert_number(&sheet.get("A1"), 1.0);
    assert_number(&sheet.get("A2"), 2.0);
}

#[test]
fn days360_matches_excel_examples() {
    let mut sheet = TestSheet::new();
    assert_number(
        &sheet.eval("=DAYS360(DATE(2011,1,1),DATE(2011,12,31))"),
        360.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2011,1,1),DATE(2011,12,31),TRUE)"),
        359.0,
    );
    assert_number(&sheet.eval("=DAYS360(\"2020-02-01\",\"2020-02-29\")"), 30.0);

    // Leap-day / end-of-month regression cases.
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,2,29),DATE(2020,3,31))"),
        30.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,2,29),DATE(2020,3,31),TRUE)"),
        31.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,2,28),DATE(2020,3,31))"),
        33.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,2,28),DATE(2020,3,31),TRUE)"),
        32.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,2,28),DATE(2020,2,29))"),
        3.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,2,28),DATE(2020,2,29),TRUE)"),
        1.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,1,31),DATE(2020,2,29))"),
        30.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,1,31),DATE(2020,2,29),TRUE)"),
        29.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,1,31),DATE(2020,2,28))"),
        28.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,1,31),DATE(2020,2,28),TRUE)"),
        28.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,1,31),DATE(2019,2,28))"),
        30.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,1,31),DATE(2019,2,28),TRUE)"),
        28.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,2,28),DATE(2019,3,31))"),
        30.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,2,28),DATE(2019,3,31),TRUE)"),
        32.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,2,15),DATE(2019,2,28))"),
        16.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,2,15),DATE(2019,2,28),TRUE)"),
        13.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,2,29),DATE(2021,2,28))"),
        360.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,2,29),DATE(2021,2,28),TRUE)"),
        359.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,2,28),DATE(2021,2,28))"),
        363.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2020,2,28),DATE(2021,2,28),TRUE)"),
        360.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,2,28),DATE(2020,2,29))"),
        360.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,2,28),DATE(2020,2,29),TRUE)"),
        361.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,2,28),DATE(2020,2,28))"),
        358.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,2,28),DATE(2020,2,28),TRUE)"),
        360.0,
    );
    assert_eq!(
        sheet.eval("=DAYS360(\"nope\",DATE(2020,1,1))"),
        Value::Error(ErrorKind::Value)
    );

    sheet.set("A1", Value::Number(f64::INFINITY));
    assert_eq!(
        sheet.eval("=DAYS360(DATE(2011,1,1),DATE(2011,12,31),A1)"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn days360_month_end_rollover_applies_outside_february() {
    let mut sheet = TestSheet::new();

    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,4,29),DATE(2019,4,30))"),
        2.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,4,29),DATE(2019,4,30),TRUE)"),
        1.0,
    );

    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,4,29),DATE(2019,5,31))"),
        32.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(2019,4,29),DATE(2019,5,31),TRUE)"),
        31.0,
    );
}

#[test]
fn days360_accounts_for_lotus_bug_feb_1900() {
    let mut sheet = TestSheet::new();
    sheet.set_date_system(ExcelDateSystem::Excel1900 { lotus_compat: true });

    assert_number(&sheet.eval("=DAYS360(DATE(1900,2,28),DATE(1900,3,1))"), 3.0);
    assert_number(
        &sheet.eval("=DAYS360(DATE(1900,2,28),DATE(1900,3,1),TRUE)"),
        3.0,
    );

    assert_number(
        &sheet.eval("=DAYS360(DATE(1900,2,28),DATE(1900,2,29))"),
        3.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(1900,2,28),DATE(1900,2,29),TRUE)"),
        1.0,
    );

    assert_number(&sheet.eval("=DAYS360(DATE(1900,2,29),DATE(1900,3,1))"), 1.0);
    assert_number(
        &sheet.eval("=DAYS360(DATE(1900,2,29),DATE(1900,3,1),TRUE)"),
        2.0,
    );

    assert_number(
        &sheet.eval("=DAYS360(DATE(1900,1,31),DATE(1900,2,29))"),
        30.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(1900,1,31),DATE(1900,2,29),TRUE)"),
        29.0,
    );

    assert_number(
        &sheet.eval("=DAYS360(DATE(1900,2,29),DATE(1900,3,31))"),
        30.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(1900,2,29),DATE(1900,3,31),TRUE)"),
        31.0,
    );

    // Feb 28 is not month-end in the Lotus-bug date system (because Feb 29 exists).
    assert_number(
        &sheet.eval("=DAYS360(DATE(1900,1,31),DATE(1900,2,28))"),
        28.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(1900,1,31),DATE(1900,2,28),TRUE)"),
        28.0,
    );
}

#[test]
fn days360_respects_lotus_compat_flag_for_feb_1900() {
    let mut sheet = TestSheet::new();
    sheet.set_date_system(ExcelDateSystem::Excel1900 {
        lotus_compat: false,
    });

    // Without the Lotus bug, Feb 28 1900 is month-end.
    assert_number(&sheet.eval("=DAYS360(DATE(1900,2,28),DATE(1900,3,1))"), 1.0);
    assert_number(
        &sheet.eval("=DAYS360(DATE(1900,2,28),DATE(1900,3,1),TRUE)"),
        3.0,
    );

    // Without the Lotus bug, Feb 28 is month-end so US/NASD adjusts it to day 30.
    assert_number(
        &sheet.eval("=DAYS360(DATE(1900,1,31),DATE(1900,2,28))"),
        30.0,
    );
    assert_number(
        &sheet.eval("=DAYS360(DATE(1900,1,31),DATE(1900,2,28),TRUE)"),
        28.0,
    );
}

#[test]
fn days360_spills_over_array_inputs() {
    let mut sheet = TestSheet::new();
    sheet.set_formula(
        "A1",
        "=DAYS360({DATE(2011,1,1);DATE(2011,1,31)},DATE(2011,12,31))",
    );
    sheet.recalc();
    assert_number(&sheet.get("A1"), 360.0);
    assert_number(&sheet.get("A2"), 330.0);
}

#[test]
fn yearfrac_matches_basis_conventions() {
    let mut sheet = TestSheet::new();
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2011,1,1),DATE(2011,12,31))"),
        1.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2011,1,1),DATE(2011,12,31),4)"),
        359.0 / 360.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,3,1),DATE(2021,3,1),1)"),
        1.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,29),DATE(2021,2,28),1)"),
        1.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,29),DATE(2022,2,28),1)"),
        2.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,29),DATE(2023,2,28),1)"),
        3.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,29),DATE(2024,2,29),1)"),
        4.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,29),DATE(2021,3,1),1)"),
        1.0 + 1.0 / 365.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2019,2,28),DATE(2020,2,29),1)"),
        1.0 + 1.0 / 366.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,29),DATE(2019,2,28),1)"),
        -(1.0 + 1.0 / 366.0),
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,29),DATE(2020,2,29),1)"),
        0.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2021,2,28),DATE(2020,2,29),1)"),
        -1.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2021,3,1),DATE(2020,2,29),1)"),
        -(1.0 + 1.0 / 365.0),
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2019,3,1),DATE(2020,2,29),1)"),
        365.0 / 366.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,29),DATE(2019,3,1),1)"),
        -(365.0 / 366.0),
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,3,1),DATE(2021,2,28),1)"),
        364.0 / 365.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2021,2,28),DATE(2020,3,1),1)"),
        -(364.0 / 365.0),
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,29),DATE(2020,3,1),1)"),
        1.0 / 365.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,3,1),DATE(2020,2,29),1)"),
        -(1.0 / 365.0),
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,28),DATE(2020,2,29),1)"),
        1.0 / 366.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,29),DATE(2020,2,28),1)"),
        -(1.0 / 366.0),
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,28),DATE(2021,2,28),1)"),
        1.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2021,2,28),DATE(2020,2,28),1)"),
        -1.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,28),DATE(2021,2,27),1)"),
        365.0 / 366.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2021,2,27),DATE(2020,2,28),1)"),
        -(365.0 / 366.0),
    );

    // Excel-style integer coercion: basis is truncated toward zero before validation.
    assert_eq!(
        sheet.eval("=YEARFRAC(DATE(2011,1,1),DATE(2011,12,31),0.9)"),
        sheet.eval("=YEARFRAC(DATE(2011,1,1),DATE(2011,12,31),0)")
    );
    assert_eq!(
        sheet.eval("=YEARFRAC(DATE(2011,1,1),DATE(2011,12,31),1.9)"),
        sheet.eval("=YEARFRAC(DATE(2011,1,1),DATE(2011,12,31),1)")
    );
    assert_eq!(
        sheet.eval("=YEARFRAC(DATE(2011,1,1),DATE(2011,12,31),-0.1)"),
        sheet.eval("=YEARFRAC(DATE(2011,1,1),DATE(2011,12,31),0)")
    );

    assert_eq!(
        sheet.eval("=YEARFRAC(DATE(2020,1,1),DATE(2020,12,31),9)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=YEARFRAC(\"nope\",DATE(2020,1,1))"),
        Value::Error(ErrorKind::Value)
    );

    sheet.set("A1", Value::Number(f64::INFINITY));
    assert_eq!(
        sheet.eval("=YEARFRAC(DATE(2020,1,1),DATE(2020,12,31),A1)"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn yearfrac_basis1_accounts_for_lotus_bug_feb_1900() {
    let mut sheet = TestSheet::new();
    sheet.set_date_system(ExcelDateSystem::Excel1900 { lotus_compat: true });

    assert_number(
        &sheet.eval("=YEARFRAC(DATE(1900,1,1),DATE(1900,12,31),1)"),
        365.0 / 366.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(1900,2,28),DATE(1900,2,29),1)"),
        1.0 / 366.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(1900,2,28),DATE(1900,3,1),1)"),
        2.0 / 366.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(1900,2,29),DATE(1900,3,1),1)"),
        1.0 / 365.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(1900,3,1),DATE(1900,2,29),1)"),
        -(1.0 / 365.0),
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(1900,2,29),DATE(1901,2,28),1)"),
        1.0,
    );
}

#[test]
fn yearfrac_basis1_respects_lotus_compat_flag_for_1900() {
    let mut sheet = TestSheet::new();
    sheet.set_date_system(ExcelDateSystem::Excel1900 {
        lotus_compat: false,
    });

    assert_number(
        &sheet.eval("=YEARFRAC(DATE(1900,1,1),DATE(1900,12,31),1)"),
        364.0 / 365.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(1900,2,28),DATE(1900,3,1),1)"),
        1.0 / 365.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(1900,3,1),DATE(1900,2,28),1)"),
        -(1.0 / 365.0),
    );
}

#[test]
fn yearfrac_basis1_uses_correct_year_length_within_year() {
    let mut sheet = TestSheet::new();

    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2019,1,1),DATE(2019,12,31),1)"),
        364.0 / 365.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2019,2,28),DATE(2019,12,31),1)"),
        306.0 / 365.0,
    );

    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,1,1),DATE(2020,12,31),1)"),
        365.0 / 366.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,28),DATE(2020,12,31),1)"),
        307.0 / 366.0,
    );

    // Denominator can span into the next year and include/exclude a leap day based on the anniversary.
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2019,3,1),DATE(2019,12,31),1)"),
        305.0 / 366.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,3,1),DATE(2020,12,31),1)"),
        305.0 / 365.0,
    );

    // One-day spans across year-end should use the correct anniversary-year denominator.
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2019,12,31),DATE(2020,1,1),1)"),
        1.0 / 366.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,12,31),DATE(2021,1,1),1)"),
        1.0 / 365.0,
    );

    // Sign handling should remain consistent for within-year spans.
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2019,12,31),DATE(2019,1,1),1)"),
        -(364.0 / 365.0),
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,12,31),DATE(2020,1,1),1)"),
        -(365.0 / 366.0),
    );

    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2019,12,31),DATE(2019,3,1),1)"),
        -(305.0 / 366.0),
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,12,31),DATE(2020,3,1),1)"),
        -(305.0 / 365.0),
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,1,1),DATE(2019,12,31),1)"),
        -(1.0 / 366.0),
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2021,1,1),DATE(2020,12,31),1)"),
        -(1.0 / 365.0),
    );

    // Near-anniversary behavior at year-end should use the correct denominator.
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2019,12,31),DATE(2020,12,30),1)"),
        365.0 / 366.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,12,30),DATE(2019,12,31),1)"),
        -(365.0 / 366.0),
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2019,12,31),DATE(2020,12,31),1)"),
        1.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,12,31),DATE(2019,12,31),1)"),
        -1.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,12,31),DATE(2021,12,30),1)"),
        364.0 / 365.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2021,12,30),DATE(2020,12,31),1)"),
        -(364.0 / 365.0),
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,12,31),DATE(2021,12,31),1)"),
        1.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2021,12,31),DATE(2020,12,31),1)"),
        -1.0,
    );

    // Feb 28 is month-end in non-leap years; the leap day can fall outside the anniversary
    // denominator.
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2019,2,28),DATE(2020,2,27),1)"),
        364.0 / 365.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,27),DATE(2019,2,28),1)"),
        -(364.0 / 365.0),
    );

    // Start dates after Feb 29 can have a 366-day denominator even when the end date is still in
    // February of the following year.
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2019,3,1),DATE(2020,2,28),1)"),
        364.0 / 366.0,
    );
    assert_number(
        &sheet.eval("=YEARFRAC(DATE(2020,2,28),DATE(2019,3,1),1)"),
        -(364.0 / 366.0),
    );
}

#[test]
fn yearfrac_spills_over_array_inputs() {
    let mut sheet = TestSheet::new();
    let expected1 = sheet.eval("=YEARFRAC(DATE(2020,1,1),DATE(2021,1,1))");
    let expected2 = sheet.eval("=YEARFRAC(DATE(2020,1,2),DATE(2021,1,1))");
    sheet.set_formula(
        "A1",
        "=YEARFRAC({DATE(2020,1,1);DATE(2020,1,2)},DATE(2021,1,1))",
    );
    sheet.recalc();
    assert_eq!(sheet.get("A1"), expected1);
    assert_eq!(sheet.get("A2"), expected2);
}

#[test]
fn datedif_matches_excel_units() {
    let mut sheet = TestSheet::new();
    assert_number(
        &sheet.eval("=DATEDIF(DATE(2011,1,15),DATE(2012,1,14),\"Y\")"),
        0.0,
    );
    assert_number(
        &sheet.eval("=DATEDIF(\"2011-01-15\",\"2012-01-14\",\"M\")"),
        11.0,
    );
    assert_number(
        &sheet.eval("=DATEDIF(DATE(2011,1,15),DATE(2012,1,14),\"D\")"),
        364.0,
    );
    assert_number(
        &sheet.eval("=DATEDIF(DATE(2011,1,15),DATE(2012,1,14),\"YM\")"),
        11.0,
    );
    assert_number(
        &sheet.eval("=DATEDIF(DATE(2011,1,15),DATE(2012,1,14),\"MD\")"),
        30.0,
    );
    assert_number(
        &sheet.eval("=DATEDIF(DATE(2011,1,15),DATE(2012,1,14),\"YD\")"),
        364.0,
    );

    assert_eq!(
        sheet.eval("=DATEDIF(DATE(2012,1,14),DATE(2011,1,15),\"D\")"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=DATEDIF(DATE(2011,1,15),DATE(2012,1,14),\"NOPE\")"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn datedif_spills_over_array_inputs() {
    let mut sheet = TestSheet::new();
    sheet.set_formula(
        "A1",
        "=DATEDIF({DATE(2011,1,15);DATE(2011,2,15)},DATE(2012,1,14),\"M\")",
    );
    sheet.recalc();
    assert_number(&sheet.get("A1"), 11.0);
    assert_number(&sheet.get("A2"), 10.0);
}

#[test]
fn hour_minute_second_extract_time_components() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=HOUR(TIME(1,2,3))"), 1.0);
    assert_number(&sheet.eval("=MINUTE(TIME(1,2,3))"), 2.0);
    assert_number(&sheet.eval("=SECOND(TIME(1,2,3))"), 3.0);
}

#[test]
fn date_time_1x1_arrays_coerce_back_to_scalars() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=ABS(TIME({1},0,0))"), 1.0 / 24.0);
    assert_number(&sheet.eval("=ABS(YEAR({DATE(2020,1,1)}))"), 2020.0);
}

#[test]
fn date_spills_over_array_inputs() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=DATE({2020;2021},1,1)");
    sheet.recalc();
    let expected_2020 = sheet.eval("=DATE(2020,1,1)");
    let expected_2021 = sheet.eval("=DATE(2021,1,1)");
    assert_eq!(sheet.get("A1"), expected_2020);
    assert_eq!(sheet.get("A2"), expected_2021);
}

#[test]
fn date_errors_on_non_finite_numbers() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Number(f64::INFINITY));
    assert_eq!(sheet.eval("=DATE(A1,1,1)"), Value::Error(ErrorKind::Num));
}

#[test]
fn date_time_mismatched_array_shapes_return_value_error() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=TIME({1,2},{3;4},0)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn weekday_and_weeknum_return_types() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=WEEKDAY(1)"), 2.0);
    assert_number(&sheet.eval("=WEEKDAY(1,2)"), 1.0);
    assert_number(&sheet.eval("=WEEKDAY(1,3)"), 0.0);
    assert_eq!(sheet.eval("=WEEKDAY(1,0)"), Value::Error(ErrorKind::Num));

    assert_number(&sheet.eval("=WEEKNUM(DATE(2020,1,1),1)"), 1.0);
    assert_number(&sheet.eval("=WEEKNUM(DATE(2020,1,5),1)"), 2.0);
    assert_number(&sheet.eval("=WEEKNUM(DATE(2020,1,5),2)"), 1.0);
    assert_number(&sheet.eval("=WEEKNUM(DATE(2020,1,6),2)"), 2.0);
    assert_number(&sheet.eval("=WEEKNUM(DATE(2021,1,1),21)"), 53.0);
    assert_number(&sheet.eval("=ISOWEEKNUM(DATE(2021,1,1))"), 53.0);
    assert_number(&sheet.eval("=ISO.WEEKNUM(DATE(2021,1,1))"), 53.0);
    assert_eq!(sheet.eval("=WEEKNUM(1,9)"), Value::Error(ErrorKind::Num));
}

#[test]
fn workday_and_networkdays_skip_weekends_and_holidays() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=WORKDAY(DATE(2020,1,1),1)"),
        sheet.eval("=DATE(2020,1,2)")
    );
    assert_eq!(
        sheet.eval("=WORKDAY(DATE(2020,1,1),1,DATE(2020,1,2))"),
        sheet.eval("=DATE(2020,1,3)")
    );

    assert_number(
        &sheet.eval("=NETWORKDAYS(DATE(2020,1,1),DATE(2020,1,10))"),
        8.0,
    );
    assert_number(
        &sheet.eval("=NETWORKDAYS(DATE(2020,1,1),DATE(2020,1,10),{DATE(2020,1,2),DATE(2020,1,3)})"),
        6.0,
    );

    assert_eq!(
        sheet.eval("=WORKDAY.INTL(DATE(2020,1,3),1,11)"),
        sheet.eval("=DATE(2020,1,4)")
    );
    assert_eq!(
        sheet.eval("=WORKDAY.INTL(DATE(2020,1,3),1,99)"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        sheet.eval("=WORKDAY.INTL(DATE(2020,1,3),1,\"abc\")"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn year_spills_over_array_inputs() {
    let mut sheet = TestSheet::new();
    sheet.set_formula("A1", "=YEAR({DATE(2019,1,1);DATE(2020,1,1)})");
    sheet.recalc();
    assert_number(&sheet.get("A1"), 2019.0);
    assert_number(&sheet.get("A2"), 2020.0);
}
