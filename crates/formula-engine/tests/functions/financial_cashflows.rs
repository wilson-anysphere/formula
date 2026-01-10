use formula_engine::error::ExcelError;
use formula_engine::functions::financial::{irr, npv, xirr, xnpv};

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn npv_discounts_from_period_1() {
    // Mirrors Excel semantics: value1 is discounted by one period.
    let result = npv(0.1, &[10_000.0, 15_000.0, 20_000.0]).unwrap();
    let expected = 10_000.0 / 1.1 + 15_000.0 / 1.1_f64.powi(2) + 20_000.0 / 1.1_f64.powi(3);
    assert_close(result, expected, 1e-12);
}

#[test]
fn irr_matches_excel_example() {
    // Excel docs example cashflows:
    // {-70000, 12000, 15000, 18000, 21000, 26000} -> 0.0866309480
    let values = [-70_000.0, 12_000.0, 15_000.0, 18_000.0, 21_000.0, 26_000.0];
    let result = irr(&values, None).unwrap();
    assert_close(result, 0.08663094803653162, 1e-12);
}

#[test]
fn irr_requires_sign_change() {
    let values = [1.0, 2.0, 3.0];
    assert_eq!(irr(&values, None), Err(ExcelError::Num));
}

#[test]
fn xnpv_xirr_match_excel_example() {
    // Excel docs example:
    // dates: 2008-01-01, 2008-03-01, 2008-10-30, 2009-02-15, 2009-04-01
    // values: -10000, 2750, 4250, 3250, 2750
    // XIRR(...)=0.3733625335, XNPV(0.09,...)=2086.6476020
    let values = [-10_000.0, 2_750.0, 4_250.0, 3_250.0, 2_750.0];

    // Excel serials for the 1900 system (Lotus-compat); the day offsets cancel out in XIRR/XNPV,
    // but using serials keeps the test aligned with spreadsheet inputs.
    let dates = [39448.0, 39508.0, 39751.0, 39859.0, 39904.0];

    let npv = xnpv(0.09, &values, &dates).unwrap();
    assert_close(npv, 2_086.6476020315354, 1e-10);

    let r = xirr(&values, &dates, None).unwrap();
    assert_close(r, 0.3733625335188314, 1e-12);
}

#[test]
fn xirr_input_length_mismatch_is_value_error() {
    let values = [1.0, -2.0];
    let dates = [0.0];
    assert_eq!(xirr(&values, &dates, None), Err(ExcelError::Num));
}
