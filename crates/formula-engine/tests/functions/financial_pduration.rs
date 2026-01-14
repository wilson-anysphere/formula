use formula_engine::functions::financial::pduration;
use formula_engine::ExcelError;

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn pduration_matches_excel_example() {
    // Excel docs: PDURATION(0.025, 2000, 2200) ≈ 3.859866163
    let periods = pduration(0.025, 2_000.0, 2_200.0).unwrap();
    assert_close(periods, 3.859866162622662, 1e-12);
}

#[test]
fn pduration_errors_on_nonpositive_rate() {
    assert_eq!(pduration(0.0, 100.0, 110.0), Err(ExcelError::Num));
    assert_eq!(pduration(-0.1, 100.0, 110.0), Err(ExcelError::Num));
}

#[test]
fn pduration_errors_on_nonpositive_values() {
    assert_eq!(pduration(0.1, 0.0, 110.0), Err(ExcelError::Num));
    assert_eq!(pduration(0.1, -100.0, 110.0), Err(ExcelError::Num));
    assert_eq!(pduration(0.1, 100.0, 0.0), Err(ExcelError::Num));
    assert_eq!(pduration(0.1, 100.0, -110.0), Err(ExcelError::Num));
}

#[test]
fn pduration_returns_div0_when_denominator_ln_is_zero() {
    // A positive rate so small that `1 + rate` rounds to 1.0 in f64.
    assert_eq!(pduration(1e-20, 100.0, 110.0), Err(ExcelError::Div0));
}
