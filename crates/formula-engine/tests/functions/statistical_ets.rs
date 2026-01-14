use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::{ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn forecast_ets_constant_series_forecasts_constant() {
    let mut sheet = TestSheet::new();
    assert_number(
        &sheet.eval("=FORECAST.ETS(6,{10,10,10,10,10},{1,2,3,4,5},1)"),
        10.0,
    );
}

#[test]
fn forecast_ets_confint_is_zero_for_perfect_fit() {
    let mut sheet = TestSheet::new();
    assert_number(
        &sheet.eval("=FORECAST.ETS.CONFINT(6,{10,10,10,10,10},{1,2,3,4,5},0.95,1)"),
        0.0,
    );
}

#[test]
fn forecast_ets_seasonality_detects_simple_pattern() {
    let mut sheet = TestSheet::new();
    assert_number(
        &sheet.eval("=FORECAST.ETS.SEASONALITY({10,20,10,20,10,20,10,20},{1,2,3,4,5,6,7,8})"),
        2.0,
    );
}

#[test]
fn forecast_ets_auto_seasonality_forecasts_repeating_pattern() {
    let mut sheet = TestSheet::new();
    // With a clear period-2 alternating pattern, the next value should repeat the first.
    assert_number(
        &sheet.eval("=FORECAST.ETS(9,{10,20,10,20,10,20,10,20},{1,2,3,4,5,6,7,8})"),
        10.0,
    );
}

#[test]
fn forecast_ets_stat_rmse_is_zero_for_perfect_fit() {
    let mut sheet = TestSheet::new();
    assert_number(
        &sheet.eval("=FORECAST.ETS.STAT({10,10,10,10,10},{1,2,3,4,5},1,1,1,8)"),
        0.0,
    );
}

#[test]
fn forecast_ets_accepts_monthly_date_timeline() {
    let system = ExcelDateSystem::EXCEL_1900;
    let mut sheet = TestSheet::new();
    sheet.set_date_system(system);

    let jan = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap() as f64;
    let feb = ymd_to_serial(ExcelDate::new(2020, 2, 1), system).unwrap() as f64;
    let mar = ymd_to_serial(ExcelDate::new(2020, 3, 1), system).unwrap() as f64;
    let apr = ymd_to_serial(ExcelDate::new(2020, 4, 1), system).unwrap() as f64;
    let may = ymd_to_serial(ExcelDate::new(2020, 5, 1), system).unwrap() as f64;

    // Constant monthly series: Excel should treat these as evenly spaced in calendar months, not
    // serial-day deltas (31/29/31...), and therefore should not raise #NUM!.
    for row in 1..=4 {
        sheet.set(&format!("A{row}"), 10.0);
    }
    sheet.set("B1", jan);
    sheet.set("B2", feb);
    sheet.set("B3", mar);
    sheet.set("B4", apr);
    sheet.set("B5", may);

    assert_number(&sheet.eval("=FORECAST.ETS(B5,A1:A4,B1:B4,1)"), 10.0);
    assert_number(
        &sheet.eval("=FORECAST.ETS.CONFINT(B5,A1:A4,B1:B4,0.95,1)"),
        0.0,
    );
    assert_number(&sheet.eval("=FORECAST.ETS.SEASONALITY(A1:A4,B1:B4)"), 1.0);
    assert_number(&sheet.eval("=FORECAST.ETS.STAT(A1:A4,B1:B4,1,1,1,8)"), 0.0);
}

#[test]
fn forecast_ets_accepts_month_end_date_timeline_across_leap_day() {
    let system = ExcelDateSystem::EXCEL_1900;
    let mut sheet = TestSheet::new();
    sheet.set_date_system(system);

    // Month-end timeline starting on Feb 29 (leap day). This sequence follows EOMONTH semantics,
    // but does *not* follow EDATE semantics (EDATE(2020-02-29,1) = 2020-03-29).
    let feb29 = ymd_to_serial(ExcelDate::new(2020, 2, 29), system).unwrap() as f64;
    let mar31 = ymd_to_serial(ExcelDate::new(2020, 3, 31), system).unwrap() as f64;
    let apr30 = ymd_to_serial(ExcelDate::new(2020, 4, 30), system).unwrap() as f64;
    let may31 = ymd_to_serial(ExcelDate::new(2020, 5, 31), system).unwrap() as f64;
    let jun30 = ymd_to_serial(ExcelDate::new(2020, 6, 30), system).unwrap() as f64;

    for row in 1..=4 {
        sheet.set(&format!("A{row}"), 10.0);
    }
    sheet.set("B1", feb29);
    sheet.set("B2", mar31);
    sheet.set("B3", apr30);
    sheet.set("B4", may31);
    sheet.set("B5", jun30);

    assert_number(&sheet.eval("=FORECAST.ETS(B5,A1:A4,B1:B4,1)"), 10.0);
    assert_number(
        &sheet.eval("=FORECAST.ETS.CONFINT(B5,A1:A4,B1:B4,0.95,1)"),
        0.0,
    );
    assert_number(&sheet.eval("=FORECAST.ETS.SEASONALITY(A1:A4,B1:B4)"), 1.0);
    assert_number(&sheet.eval("=FORECAST.ETS.STAT(A1:A4,B1:B4,1,1,1,8)"), 0.0);
}

#[test]
fn forecast_ets_accepts_yearly_date_timeline_across_leap_years() {
    let system = ExcelDateSystem::EXCEL_1900;
    let mut sheet = TestSheet::new();
    sheet.set_date_system(system);

    let d2019 = ymd_to_serial(ExcelDate::new(2019, 1, 1), system).unwrap() as f64;
    let d2020 = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap() as f64;
    let d2021 = ymd_to_serial(ExcelDate::new(2021, 1, 1), system).unwrap() as f64;
    let d2022 = ymd_to_serial(ExcelDate::new(2022, 1, 1), system).unwrap() as f64;

    sheet.set("A1", 10.0);
    sheet.set("A2", 10.0);
    sheet.set("A3", 10.0);

    sheet.set("B1", d2019);
    sheet.set("B2", d2020);
    sheet.set("B3", d2021);
    sheet.set("B4", d2022);

    assert_number(&sheet.eval("=FORECAST.ETS(B4,A1:A3,B1:B3,1)"), 10.0);
}

#[test]
fn forecast_ets_accepts_monthly_date_timeline_excel1904() {
    let system = ExcelDateSystem::Excel1904;
    let mut sheet = TestSheet::new();
    sheet.set_date_system(system);

    let jan = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap() as f64;
    let feb = ymd_to_serial(ExcelDate::new(2020, 2, 1), system).unwrap() as f64;
    let mar = ymd_to_serial(ExcelDate::new(2020, 3, 1), system).unwrap() as f64;
    let apr = ymd_to_serial(ExcelDate::new(2020, 4, 1), system).unwrap() as f64;
    let may = ymd_to_serial(ExcelDate::new(2020, 5, 1), system).unwrap() as f64;

    for row in 1..=4 {
        sheet.set(&format!("A{row}"), 10.0);
    }
    sheet.set("B1", jan);
    sheet.set("B2", feb);
    sheet.set("B3", mar);
    sheet.set("B4", apr);
    sheet.set("B5", may);

    assert_number(&sheet.eval("=FORECAST.ETS(B5,A1:A4,B1:B4,1)"), 10.0);
}

#[test]
fn forecast_ets_dedupes_overlapping_reference_unions() {
    let mut sheet = TestSheet::new();

    sheet.set("A1", 10.0);
    sheet.set("A2", 20.0);
    sheet.set("A3", 30.0);

    sheet.set("B1", 1.0);
    sheet.set("B2", 2.0);
    sheet.set("B3", 3.0);

    // Use SUM aggregation (7) so duplicate timeline entries materially change the input series if
    // we accidentally double-count overlapping union cells.
    let base = sheet.eval("=FORECAST.ETS(4,A1:A3,B1:B3,1,1,7)");
    let union = sheet.eval("=FORECAST.ETS(4,(A1:A3,A2:A3),(B1:B3,B2:B3),1,1,7)");

    match (base, union) {
        (Value::Number(a), Value::Number(b)) => {
            assert!(
                (a - b).abs() < 1e-9,
                "expected union to match base; got {a} vs {b}"
            );
        }
        (a, b) => panic!("expected numbers, got {a:?} and {b:?}"),
    }
}

#[test]
fn forecast_ets_rejects_mismatched_series_lengths() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=FORECAST.ETS(4,{1,2,3},{1,2})"),
        Value::Error(ErrorKind::NA)
    );
}

#[test]
fn forecast_ets_confint_rejects_invalid_confidence_level() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=FORECAST.ETS.CONFINT(6,{10,10,10,10,10},{1,2,3,4,5},1.5,1)"),
        Value::Error(ErrorKind::Num)
    );
}
