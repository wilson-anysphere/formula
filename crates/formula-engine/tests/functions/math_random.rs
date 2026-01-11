use formula_engine::functions::math;
use formula_engine::ExcelError;
use formula_engine::Value;

use super::harness::TestSheet;

#[test]
fn rand_is_in_unit_interval() {
    for _ in 0..100 {
        let r = math::rand();
        assert!(r >= 0.0);
        assert!(r < 1.0);
    }
}

#[test]
fn randbetween_respects_bounds() {
    assert_eq!(math::randbetween(5.0, 5.0).unwrap(), 5);

    for _ in 0..100 {
        let r = math::randbetween(1.0, 3.0).unwrap();
        assert!((1..=3).contains(&r));
    }

    assert_eq!(math::randbetween(5.0, 1.0).unwrap_err(), ExcelError::Num);
}

#[test]
fn rand_worksheet_function_is_in_unit_interval() {
    let mut sheet = TestSheet::new();
    for _ in 0..100 {
        match sheet.eval("=RAND()") {
            Value::Number(n) => {
                assert!(n >= 0.0);
                assert!(n < 1.0);
            }
            other => panic!("expected numeric result, got {other:?}"),
        }
    }
}

#[test]
fn randbetween_worksheet_function_respects_bounds() {
    let mut sheet = TestSheet::new();
    for _ in 0..100 {
        match sheet.eval("=RANDBETWEEN(1,3)") {
            Value::Number(n) => {
                assert_eq!(n.fract(), 0.0);
                assert!((1.0..=3.0).contains(&n));
            }
            other => panic!("expected numeric result, got {other:?}"),
        }
    }
}

#[test]
fn random_functions_accept_xlfn_prefix() {
    let mut sheet = TestSheet::new();
    match sheet.eval("=_xlfn.RAND()") {
        Value::Number(n) => {
            assert!(n >= 0.0);
            assert!(n < 1.0);
        }
        other => panic!("expected numeric result, got {other:?}"),
    }

    match sheet.eval("=_xlfn.RANDBETWEEN(1,3)") {
        Value::Number(n) => {
            assert_eq!(n.fract(), 0.0);
            assert!((1.0..=3.0).contains(&n));
        }
        other => panic!("expected numeric result, got {other:?}"),
    }
}
