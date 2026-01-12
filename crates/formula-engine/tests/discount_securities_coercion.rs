use formula_engine::{Engine, Value};

fn assert_number(value: Value) -> f64 {
    match value {
        Value::Number(n) => n,
        other => panic!("expected number, got {other:?}"),
    }
}

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn discount_security_optional_basis_treats_blank_as_zero() {
    let mut engine = Engine::new();

    // B1 is unset/blank by default.
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=DISC(DATE(2020,1,1),DATE(2021,1,1),97,100,B1)",
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=DISC(DATE(2020,1,1),DATE(2021,1,1),97,100,0)")
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            "=PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),0.05,100,B1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            "=PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),0.05,100,0)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "A5",
            "=RECEIVED(DATE(2020,1,1),DATE(2021,1,1),95,0.05,B1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A6",
            "=RECEIVED(DATE(2020,1,1),DATE(2021,1,1),95,0.05,0)",
        )
        .unwrap();

    engine.recalculate();

    assert_close(assert_number(engine.get_cell_value("Sheet1", "A1")), 0.03, 1e-12);
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        assert_number(engine.get_cell_value("Sheet1", "A2")),
        1e-12,
    );

    assert_close(assert_number(engine.get_cell_value("Sheet1", "A3")), 95.0, 1e-12);
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A3")),
        assert_number(engine.get_cell_value("Sheet1", "A4")),
        1e-12,
    );

    assert_close(assert_number(engine.get_cell_value("Sheet1", "A5")), 100.0, 1e-12);
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A5")),
        assert_number(engine.get_cell_value("Sheet1", "A6")),
        1e-12,
    );
}

#[test]
fn discount_security_date_coercion_floors_numeric_serials() {
    let mut engine = Engine::new();

    // 43831 = 2020-01-01, 44197 = 2021-01-01 in the Excel 1900 date system.
    engine
        .set_cell_formula("Sheet1", "A1", "=DISC(43831.9,44197.2,97,100)")
        .unwrap();

    engine.recalculate();
    assert_close(assert_number(engine.get_cell_value("Sheet1", "A1")), 0.03, 1e-12);
}

