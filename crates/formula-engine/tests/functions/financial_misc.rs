use formula_engine::functions::financial::{dollarde, dollarfr, ispmt};
use formula_engine::ExcelError;

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn ispmt_matches_equal_principal_schedule() {
    // Simple 3-period, equal-principal repayment:
    // pv = 300, rate = 10%, interest is computed on remaining principal:
    // period 1: 300 * 0.1 = 30
    // period 2: 200 * 0.1 = 20
    // period 3: 100 * 0.1 = 10
    assert_close(ispmt(0.1, 1.0, 3.0, 300.0).unwrap(), -30.0, 1e-12);
    assert_close(ispmt(0.1, 2.0, 3.0, 300.0).unwrap(), -20.0, 1e-12);
    assert_close(ispmt(0.1, 3.0, 3.0, 300.0).unwrap(), -10.0, 1e-12);
}

#[test]
fn ispmt_errors_match_excel() {
    assert_eq!(ispmt(0.1, 0.0, 3.0, 300.0), Err(ExcelError::Num));
    assert_eq!(ispmt(0.1, 4.0, 3.0, 300.0), Err(ExcelError::Num));
    assert_eq!(ispmt(0.1, 1.0, 0.0, 300.0), Err(ExcelError::Num));
    assert_eq!(ispmt(f64::NAN, 1.0, 3.0, 300.0), Err(ExcelError::Num));
}

#[test]
fn dollarde_matches_excel_example() {
    // Excel docs: DOLLARDE(1.02, 16) -> 1.125
    assert_close(dollarde(1.02, 16.0).unwrap(), 1.125, 1e-12);
    assert_close(dollarde(-1.02, 16.0).unwrap(), -1.125, 1e-12);
}

#[test]
fn dollarfr_matches_excel_example_and_roundtrips() {
    // Excel docs: DOLLARFR(1.125, 16) -> 1.02
    assert_close(dollarfr(1.125, 16.0).unwrap(), 1.02, 1e-12);
    assert_close(dollarfr(-1.125, 16.0).unwrap(), -1.02, 1e-12);

    let frac = dollarfr(1.125, 16.0).unwrap();
    assert_close(dollarde(frac, 16.0).unwrap(), 1.125, 1e-12);
}

#[test]
fn dollarde_dollarfr_error_cases() {
    // Invalid fraction.
    assert_eq!(dollarde(1.02, 0.0), Err(ExcelError::Div0));
    assert_eq!(dollarfr(1.125, 0.0), Err(ExcelError::Div0));
    assert_eq!(dollarde(1.02, -16.0), Err(ExcelError::Num));
    assert_eq!(dollarfr(1.125, -16.0), Err(ExcelError::Num));

    // Non-finite inputs.
    assert_eq!(dollarde(f64::NAN, 16.0), Err(ExcelError::Num));
    assert_eq!(dollarfr(1.125, f64::INFINITY), Err(ExcelError::Num));
}
