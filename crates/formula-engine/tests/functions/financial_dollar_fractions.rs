use formula_engine::functions::financial::{dollarde, dollarfr};
use formula_engine::ExcelError;

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn dollarde_matches_excel_doc_examples() {
    // Excel docs: DOLLARDE(1.02, 16) == 1.125
    let n = dollarde(1.02, 16.0).unwrap();
    assert_close(n, 1.125, 1e-12);

    // Excel docs: DOLLARDE(1.1, 32) == 1.3125
    let n = dollarde(1.1, 32.0).unwrap();
    assert_close(n, 1.3125, 1e-12);
}

#[test]
fn dollarfr_matches_excel_doc_examples() {
    // Excel docs: DOLLARFR(1.125, 16) == 1.02
    let n = dollarfr(1.125, 16.0).unwrap();
    assert_close(n, 1.02, 1e-12);

    // Excel docs: DOLLARFR(1.3125, 32) == 1.1
    let n = dollarfr(1.3125, 32.0).unwrap();
    assert_close(n, 1.1, 1e-12);
}

#[test]
fn negative_inputs_preserve_sign() {
    let n = dollarde(-1.02, 16.0).unwrap();
    assert_close(n, -1.125, 1e-12);

    let n = dollarfr(-1.125, 16.0).unwrap();
    assert_close(n, -1.02, 1e-12);
}

#[test]
fn fraction_is_truncated() {
    let n = dollarde(1.02, 16.9).unwrap();
    assert_close(n, 1.125, 1e-12);

    let n = dollarfr(1.125, 16.9).unwrap();
    assert_close(n, 1.02, 1e-12);
}

#[test]
fn error_cases() {
    assert_eq!(dollarde(1.02, 0.0), Err(ExcelError::Div0));
    assert_eq!(dollarde(1.02, 0.9), Err(ExcelError::Div0));
    assert_eq!(dollarde(1.02, -16.0), Err(ExcelError::Num));

    assert_eq!(dollarfr(1.125, 0.0), Err(ExcelError::Div0));
    assert_eq!(dollarfr(1.125, 0.9), Err(ExcelError::Div0));
    assert_eq!(dollarfr(1.125, -16.0), Err(ExcelError::Num));

    assert_eq!(dollarde(f64::INFINITY, 16.0), Err(ExcelError::Num));
    assert_eq!(dollarde(1.0, f64::NAN), Err(ExcelError::Num));
    assert_eq!(dollarfr(f64::NEG_INFINITY, 16.0), Err(ExcelError::Num));
    assert_eq!(dollarfr(1.0, f64::INFINITY), Err(ExcelError::Num));
}
