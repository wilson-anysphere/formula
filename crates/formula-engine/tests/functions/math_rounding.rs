use formula_engine::functions::math;
use formula_engine::ExcelError;

#[test]
fn ceiling_floor_legacy_require_sign_match() {
    assert_eq!(math::ceiling(4.3, 2.0).unwrap(), 6.0);
    assert_eq!(math::floor(4.3, 2.0).unwrap(), 4.0);

    assert_eq!(math::ceiling(-4.3, -2.0).unwrap(), -4.0);
    assert_eq!(math::floor(-4.3, -2.0).unwrap(), -6.0);

    assert_eq!(math::ceiling(-4.3, 2.0).unwrap_err(), ExcelError::Num);
    assert_eq!(math::floor(-4.3, 2.0).unwrap_err(), ExcelError::Num);
}

#[test]
fn ceiling_math_and_floor_math_handle_negative_modes() {
    assert_eq!(math::ceiling_math(-5.5, Some(2.0), None).unwrap(), -4.0);
    assert_eq!(math::ceiling_math(-5.5, Some(2.0), Some(1.0)).unwrap(), -6.0);

    assert_eq!(math::floor_math(-5.5, Some(2.0), None).unwrap(), -6.0);
    assert_eq!(math::floor_math(-5.5, Some(2.0), Some(1.0)).unwrap(), -4.0);
}

#[test]
fn precise_and_iso_ceiling_ignore_significance_sign() {
    assert_eq!(math::ceiling_precise(-4.3, None).unwrap(), -4.0);
    assert_eq!(math::floor_precise(-4.3, None).unwrap(), -5.0);
    assert_eq!(math::iso_ceiling(-4.3, Some(-2.0)).unwrap(), -4.0);
}

