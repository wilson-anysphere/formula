use formula_engine::{Engine, ErrorKind, Value};

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

fn assert_number(cell: Value) -> f64 {
    match cell {
        Value::Number(n) => n,
        other => panic!("expected number, got {other:?}"),
    }
}

#[test]
fn evaluates_discount_security_and_tbill_financial_functions() {
    let mut engine = Engine::new();

    // Use dates whose YEARFRAC(.,.,0) is an integer to keep expected values simple.
    engine
        .set_cell_formula("Sheet1", "A1", "=DISC(DATE(2020,1,1),DATE(2021,1,1),97,100)")
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),0.05,100)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            "=YIELDDISC(DATE(2020,1,1),DATE(2021,1,1),97,100)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            "=INTRATE(DATE(2020,1,1),DATE(2021,1,1),97,100)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A5",
            "=RECEIVED(DATE(2020,1,1),DATE(2021,1,1),95,0.05)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "A6",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A7",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,100.76923076923077)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "A8",
            "=TBILLPRICE(DATE(2020,1,1),DATE(2020,7,1),0.05)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A9",
            "=TBILLYIELD(DATE(2020,1,1),DATE(2020,7,1),97.47222222222223)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A10",
            "=TBILLEQ(DATE(2020,1,1),DATE(2020,12,31),0.05)",
        )
        .unwrap();

    engine.recalculate();

    assert_close(assert_number(engine.get_cell_value("Sheet1", "A1")), 0.03, 1e-12);
    assert_close(assert_number(engine.get_cell_value("Sheet1", "A2")), 95.0, 1e-12);
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A3")),
        3.0 / 97.0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A4")),
        3.0 / 97.0,
        1e-12,
    );
    assert_close(assert_number(engine.get_cell_value("Sheet1", "A5")), 100.0, 1e-12);

    let expected_pricemat = 110.0 / 1.04 - 5.0;
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A6")),
        expected_pricemat,
        1e-12,
    );
    assert_close(assert_number(engine.get_cell_value("Sheet1", "A7")), 0.04, 1e-12);

    let expected_tbillprice = 100.0 * (1.0 - 0.05 * 182.0 / 360.0);
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A8")),
        expected_tbillprice,
        1e-12,
    );
    let expected_tbillyield = (100.0 - 97.47222222222223) / 97.47222222222223 * (360.0 / 182.0);
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A9")),
        expected_tbillyield,
        1e-12,
    );

    let dsm: f64 = 365.0;
    let price_factor: f64 = 1.0 - 0.05 * dsm / 360.0;
    let expected_tbilleq = 2.0 * ((1.0 / price_factor).sqrt() - 1.0);
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A10")),
        expected_tbilleq,
        1e-12,
    );
}

#[test]
fn discount_security_functions_validate_dates_and_basis() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=DISC(DATE(2020,1,1),DATE(2020,1,1),97,100)")
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=DISC(DATE(2020,1,1),DATE(2021,1,1),97,100,5)",
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=DISC(\"not a date\",DATE(2021,1,1),97,100)")
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            "=RECEIVED(DATE(2020,1,1),DATE(2021,1,1),100,1)",
        )
        .unwrap();

    engine.recalculate();

    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Error(ErrorKind::Num));
    assert_eq!(engine.get_cell_value("Sheet1", "A2"), Value::Error(ErrorKind::Num));
    assert_eq!(
        engine.get_cell_value("Sheet1", "A3"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(engine.get_cell_value("Sheet1", "A4"), Value::Error(ErrorKind::Div0));
}
