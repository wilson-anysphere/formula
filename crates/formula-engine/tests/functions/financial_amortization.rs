use formula_engine::functions::financial::{cumipmt, cumprinc, pmt};
use formula_engine::ExcelError;

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn cumipmt_matches_excel_example() {
    // Excel docs: CUMIPMT(0.09/12, 30*12, 125000, 13, 24, 0) ≈ -11135.23213
    let rate = 0.09 / 12.0;
    let nper = 30.0 * 12.0;
    let pv = 125_000.0;

    let interest = cumipmt(rate, nper, pv, 13.0, 24.0, 0.0).unwrap();
    assert_close(interest, -11_135.232130750843, 1e-10);
}

#[test]
fn cumprinc_matches_excel_example() {
    // Excel docs: CUMPRINC(0.09/12, 30*12, 125000, 13, 24, 0) ≈ -934.1071234
    let rate = 0.09 / 12.0;
    let nper = 30.0 * 12.0;
    let pv = 125_000.0;

    let principal = cumprinc(rate, nper, pv, 13.0, 24.0, 0.0).unwrap();
    assert_close(principal, -934.107123420897, 1e-10);
}

#[test]
fn cum_interest_plus_principal_equals_total_payment() {
    let rate = 0.08 / 12.0;
    let nper = 10.0 * 12.0;
    let pv = 10_000.0;
    let start = 1.0;
    let end = 12.0;
    let typ = 0.0;

    let payment = pmt(rate, nper, pv, None, Some(typ)).unwrap();
    let total_payment = payment * (end - start + 1.0);
    let interest = cumipmt(rate, nper, pv, start, end, typ).unwrap();
    let principal = cumprinc(rate, nper, pv, start, end, typ).unwrap();

    assert_close(total_payment, interest + principal, 1e-10);
}

#[test]
fn amortization_rejects_invalid_inputs() {
    // Invalid type
    assert_eq!(
        cumipmt(0.1, 10.0, 1000.0, 1.0, 2.0, 2.0),
        Err(ExcelError::Num)
    );
    assert_eq!(
        cumprinc(0.1, 10.0, 1000.0, 1.0, 2.0, -1.0),
        Err(ExcelError::Num)
    );

    // Non-positive rate/nper/pv
    assert_eq!(
        cumipmt(0.0, 10.0, 1000.0, 1.0, 2.0, 0.0),
        Err(ExcelError::Num)
    );
    assert_eq!(
        cumipmt(0.1, 0.0, 1000.0, 1.0, 2.0, 0.0),
        Err(ExcelError::Num)
    );
    assert_eq!(cumipmt(0.1, 10.0, 0.0, 1.0, 2.0, 0.0), Err(ExcelError::Num));

    // Invalid period bounds.
    assert_eq!(
        cumipmt(0.1, 10.0, 1000.0, 0.0, 2.0, 0.0),
        Err(ExcelError::Num)
    );
    assert_eq!(
        cumipmt(0.1, 10.0, 1000.0, 3.0, 2.0, 0.0),
        Err(ExcelError::Num)
    );
    assert_eq!(
        cumipmt(0.1, 10.0, 1000.0, 1.0, 11.0, 0.0),
        Err(ExcelError::Num)
    );

    // Non-finite values.
    assert_eq!(
        cumipmt(f64::INFINITY, 10.0, 1000.0, 1.0, 2.0, 0.0),
        Err(ExcelError::Num)
    );
    assert_eq!(
        cumprinc(0.1, f64::NAN, 1000.0, 1.0, 2.0, 0.0),
        Err(ExcelError::Num)
    );
}
