use formula_engine::functions::financial::{
    effect, fv, ipmt, nominal, nper, pmt, ppmt, pv, rate, rri,
};
use formula_engine::ExcelError;

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
fn rate_converges_with_zero_guess() {
    let r = rate(4.0 * 12.0, -200.0, 8_000.0, None, None, Some(0.0)).unwrap();
    assert_close(r, 0.00770147248820165, 1e-12);
}

#[test]
fn rate_returns_num_when_no_solution_exists() {
    // With PV and FV both positive and no payments, there's no real rate > -1
    // that makes PV*(1+r)^n + FV == 0.
    let result = rate(10.0, 0.0, 1_000.0, Some(1_000.0), None, None);
    assert_eq!(result, Err(ExcelError::Num));
}

#[test]
fn effect_and_nominal_roundtrip() {
    let nominal_rate = 0.0525;
    let npery = 4.0;

    let eff = effect(nominal_rate, npery).unwrap();
    let expected = (1.0 + nominal_rate / 4.0).powi(4) - 1.0;
    assert_close(eff, expected, 1e-12);

    let back = nominal(eff, npery).unwrap();
    assert_close(back, nominal_rate, 1e-12);

    // Excel-style integer coercion: `npery` is truncated toward zero before validation.
    let eff_trunc_2 = effect(nominal_rate, 2.0).unwrap();
    assert_close(effect(nominal_rate, 2.9).unwrap(), eff_trunc_2, 1e-12);

    let eff_trunc_1 = effect(nominal_rate, 1.0).unwrap();
    assert_close(effect(nominal_rate, 1.1).unwrap(), eff_trunc_1, 1e-12);
    assert_close(
        effect(nominal_rate, 1.999999999).unwrap(),
        eff_trunc_1,
        1e-12,
    );

    let nom_trunc_2 = nominal(eff, 2.0).unwrap();
    assert_close(nominal(eff, 2.9).unwrap(), nom_trunc_2, 1e-12);
    let nom_trunc_1 = nominal(eff, 1.0).unwrap();
    assert_close(nominal(eff, 1.1).unwrap(), nom_trunc_1, 1e-12);

    assert_eq!(effect(nominal_rate, 0.0), Err(ExcelError::Num));
    assert_eq!(nominal(eff, 0.0), Err(ExcelError::Num));
}

#[test]
fn rri_matches_simple_growth() {
    let rate = rri(2.0, 100.0, 121.0).unwrap();
    assert_close(rate, 0.1, 1e-12);

    assert_eq!(rri(0.0, 100.0, 121.0), Err(ExcelError::Num));
    assert_eq!(rri(2.0, 0.0, 121.0), Err(ExcelError::Num));
    assert_eq!(rri(2.0, -100.0, 121.0), Err(ExcelError::Num));
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

#[test]
fn ipmt_type_beginning_first_period_is_zero() {
    let rate_per_period = 0.1;
    let nper = 2.0;
    let pv_amount = 1_000.0;

    let payment = pmt(rate_per_period, nper, pv_amount, None, Some(1.0)).unwrap();

    let interest_first = ipmt(rate_per_period, 1.0, nper, pv_amount, None, Some(1.0)).unwrap();
    let principal_first = ppmt(rate_per_period, 1.0, nper, pv_amount, None, Some(1.0)).unwrap();

    assert_close(interest_first, 0.0, 1e-12);
    assert_close(payment, principal_first, 1e-12);

    // Subsequent periods should still satisfy PMT = IPMT + PPMT.
    let interest_second = ipmt(rate_per_period, 2.0, nper, pv_amount, None, Some(1.0)).unwrap();
    let principal_second = ppmt(rate_per_period, 2.0, nper, pv_amount, None, Some(1.0)).unwrap();
    assert_close(payment, interest_second + principal_second, 1e-12);
}
