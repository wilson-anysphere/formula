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

    // Z1 is unset/blank by default.
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=DISC(DATE(2020,1,1),DATE(2021,1,1),97,100,Z1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=DISC(DATE(2020,1,1),DATE(2021,1,1),97,100,0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            "=DISC(DATE(2020,1,1),DATE(2021,1,1),97,100,)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "B1",
            "=PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),0.05,100,Z1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B2",
            "=PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),0.05,100,0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B3",
            "=PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),0.05,100,)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "C1",
            "=RECEIVED(DATE(2020,1,1),DATE(2021,1,1),95,0.05,Z1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C2",
            "=RECEIVED(DATE(2020,1,1),DATE(2021,1,1),95,0.05,0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C3",
            "=RECEIVED(DATE(2020,1,1),DATE(2021,1,1),95,0.05,)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "D1",
            "=YIELDDISC(DATE(2020,1,1),DATE(2021,1,1),97,100,Z1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D2",
            "=YIELDDISC(DATE(2020,1,1),DATE(2021,1,1),97,100,0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "D3",
            "=YIELDDISC(DATE(2020,1,1),DATE(2021,1,1),97,100,)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "E1",
            "=INTRATE(DATE(2020,1,1),DATE(2021,1,1),97,100,Z1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "E2",
            "=INTRATE(DATE(2020,1,1),DATE(2021,1,1),97,100,0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "E3",
            "=INTRATE(DATE(2020,1,1),DATE(2021,1,1),97,100,)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "F1",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,Z1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "F2",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "F3",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "G1",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,100.76923076923077,Z1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "G2",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,100.76923076923077,0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "G3",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,100.76923076923077,)",
        )
        .unwrap();

    engine.recalculate();

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        0.03,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        assert_number(engine.get_cell_value("Sheet1", "A2")),
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        assert_number(engine.get_cell_value("Sheet1", "A3")),
        1e-12,
    );

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "B1")),
        95.0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "B1")),
        assert_number(engine.get_cell_value("Sheet1", "B2")),
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "B1")),
        assert_number(engine.get_cell_value("Sheet1", "B3")),
        1e-12,
    );

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "C1")),
        100.0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "C1")),
        assert_number(engine.get_cell_value("Sheet1", "C2")),
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "C1")),
        assert_number(engine.get_cell_value("Sheet1", "C3")),
        1e-12,
    );

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "D1")),
        assert_number(engine.get_cell_value("Sheet1", "D2")),
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "D1")),
        assert_number(engine.get_cell_value("Sheet1", "D3")),
        1e-12,
    );

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "E1")),
        assert_number(engine.get_cell_value("Sheet1", "E2")),
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "E1")),
        assert_number(engine.get_cell_value("Sheet1", "E3")),
        1e-12,
    );

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "F1")),
        assert_number(engine.get_cell_value("Sheet1", "F2")),
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "F1")),
        assert_number(engine.get_cell_value("Sheet1", "F3")),
        1e-12,
    );

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "G1")),
        assert_number(engine.get_cell_value("Sheet1", "G2")),
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "G1")),
        assert_number(engine.get_cell_value("Sheet1", "G3")),
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
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        0.03,
        1e-12,
    );
}
