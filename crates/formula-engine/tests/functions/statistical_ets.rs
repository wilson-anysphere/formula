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

