use formula_engine::functions::math;
use formula_engine::ExcelError;

#[test]
fn product_multiplies_values() {
    assert_eq!(math::product(&[2.0, 3.0, 4.0]).unwrap(), 24.0);
    assert_eq!(math::product(&[]).unwrap(), 1.0);
}

#[test]
fn power_matches_excel_domain_errors() {
    assert_eq!(math::power(2.0, 3.0).unwrap(), 8.0);
    assert_eq!(math::power(0.0, -1.0).unwrap_err(), ExcelError::Div0);
    assert_eq!(math::power(-1.0, 0.5).unwrap_err(), ExcelError::Num);
}

#[test]
fn ln_log_exp_match_known_values() {
    let ln_e = math::ln(std::f64::consts::E).unwrap();
    assert!((ln_e - 1.0).abs() < 1.0e-12);

    assert_eq!(math::log(10.0, None).unwrap(), 1.0);
    assert_eq!(math::log(8.0, Some(2.0)).unwrap(), 3.0);
    assert_eq!(math::log(10.0, Some(1.0)).unwrap_err(), ExcelError::Num);
    assert_eq!(math::ln(-1.0).unwrap_err(), ExcelError::Num);

    let exp_1 = math::exp(1.0).unwrap();
    assert!((exp_1 - std::f64::consts::E).abs() < 1.0e-12);
}

