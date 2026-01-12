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
fn discount_security_basis_coerces_boolean_and_numeric_text() {
    let mut engine = Engine::new();

    // Use dates where different bases produce different YEARFRAC results (2020 is a leap year).
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=DISC(DATE(2020,1,1),DATE(2020,7,1),97,100,TRUE)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=DISC(DATE(2020,1,1),DATE(2020,7,1),97,100,1)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "B1",
            r#"=PRICEDISC(DATE(2020,1,1),DATE(2020,7,1),0.05,100,"2")"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B2",
            "=PRICEDISC(DATE(2020,1,1),DATE(2020,7,1),0.05,100,2)",
        )
        .unwrap();

    engine.recalculate();

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        assert_number(engine.get_cell_value("Sheet1", "A2")),
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "B1")),
        assert_number(engine.get_cell_value("Sheet1", "B2")),
        1e-12,
    );
}

#[test]
fn tbill_functions_parse_iso_date_text() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            r#"=TBILLPRICE("2020-01-01","2020-07-01",0.05)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=TBILLPRICE(DATE(2020,1,1),DATE(2020,7,1),0.05)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "B1",
            r#"=TBILLEQ("2020-01-01","2020-12-31",0.05)"#,
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B2",
            "=TBILLEQ(DATE(2020,1,1),DATE(2020,12,31),0.05)",
        )
        .unwrap();

    engine.recalculate();

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        assert_number(engine.get_cell_value("Sheet1", "A2")),
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "B1")),
        assert_number(engine.get_cell_value("Sheet1", "B2")),
        1e-12,
    );
}

