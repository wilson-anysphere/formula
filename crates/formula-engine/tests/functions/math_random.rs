use formula_engine::functions::math;
use formula_engine::ExcelError;

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

