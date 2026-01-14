use formula_engine::functions::financial::{dollarde, dollarfr};
use formula_engine::value::{ErrorKind, Value};
use formula_engine::ExcelError;

use super::harness::{assert_number, TestSheet};

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

#[test]
fn builtins_evaluate_doc_examples() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=DOLLARDE(1.02, 16)"), 1.125);
    assert_number(&sheet.eval("=DOLLARFR(1.125, 16)"), 1.02);
    assert_number(&sheet.eval("=DOLLARDE(1.1, 32)"), 1.3125);
    assert_number(&sheet.eval("=DOLLARFR(1.3125, 32)"), 1.1);

    assert_number(&sheet.eval("=DOLLARDE(-1.02, 16)"), -1.125);
    assert_number(&sheet.eval("=DOLLARFR(-1.125, 16)"), -1.02);
    assert_number(&sheet.eval("=DOLLARDE(-1.1, 32)"), -1.3125);
    assert_number(&sheet.eval("=DOLLARFR(-1.3125, 32)"), -1.1);

    // Single-digit fraction denominators use a scale of 10^1.
    assert_number(&sheet.eval("=DOLLARDE(1.2, 4)"), 1.5);
    assert_number(&sheet.eval("=DOLLARFR(1.5, 4)"), 1.2);

    // Fraction is truncated to an integer.
    assert_number(&sheet.eval("=DOLLARDE(1.02, 16.9)"), 1.125);
    assert_number(&sheet.eval("=DOLLARFR(1.125, 16.9)"), 1.02);

    // fraction==0 -> #DIV/0!
    assert_eq!(
        sheet.eval("=DOLLARDE(1.02, 0)"),
        Value::Error(ErrorKind::Div0)
    );
    assert_eq!(
        sheet.eval("=DOLLARFR(1.125, 0)"),
        Value::Error(ErrorKind::Div0)
    );
}
