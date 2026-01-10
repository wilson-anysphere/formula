use formula_engine::functions::financial::{fv, ipmt, nper, pmt, ppmt, pv, rate};

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn pmt_matches_excel_example() {
    // Excel docs: PMT(0.08/12, 10*12, 10000) ≈ -121.3275944
    let payment = pmt(0.08 / 12.0, 10.0 * 12.0, 10_000.0, None, None).unwrap();
    assert_close(payment, -121.32759435535776, 1e-10);
}

#[test]
fn pv_matches_excel_example() {
    // Excel docs: PV(0.08/12, 20*12, -500) ≈ 59777.14585
    let present = pv(0.08 / 12.0, 20.0 * 12.0, -500.0, None, None).unwrap();
    assert_close(present, 59_777.14585118777, 1e-9);
}

#[test]
fn fv_matches_excel_example() {
    // Example: FV(0.06/12, 10*12, -200, -1000) ≈ 34595.26610
    let future = fv(0.06 / 12.0, 10.0 * 12.0, -200.0, Some(-1_000.0), None).unwrap();
    assert_close(future, 34_595.266095323896, 1e-9);
}

#[test]
fn nper_matches_formula_rearrangement() {
    let r = 0.05 / 12.0;
    let periods = nper(r, -100.0, 5_000.0, Some(0.0), None).unwrap();

    // Validate by reconstructing PV using the computed NPER.
    let present = pv(r, periods, -100.0, None, None).unwrap();
    assert_close(present, 5_000.0, 1e-9);
}

#[test]
fn rate_matches_excel_example() {
    // Excel docs: RATE(4*12, -200, 8000) ≈ 0.0077014725
    let r = rate(4.0 * 12.0, -200.0, 8_000.0, None, None, None).unwrap();
    assert_close(r, 0.00770147248820165, 1e-12);
}

#[test]
fn ipmt_ppmt_decompose_payment() {
    let rate_per_period = 0.1;
    let nper = 2.0;
    let pv_amount = 1_000.0;

    let payment = pmt(rate_per_period, nper, pv_amount, None, None).unwrap();
    let interest = ipmt(rate_per_period, 1.0, nper, pv_amount, None, None).unwrap();
    let principal = ppmt(rate_per_period, 1.0, nper, pv_amount, None, None).unwrap();

    assert_close(payment, interest + principal, 1e-12);
}
