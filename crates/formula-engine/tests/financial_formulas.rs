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
fn evaluates_time_value_financial_functions() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula("Sheet1", "A1", "=PMT(0.08/12, 10*12, 10000)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=PV(0.08/12, 20*12, -500)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=RATE(4*12, -200, 8000)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=EFFECT(0.0525, 4)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", "=NOMINAL(A4, 4)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A6", "=RRI(2, 100, 121)")
        .unwrap();

    engine.recalculate();

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        -121.32759435535776,
        1e-10,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A2")),
        59_777.14585118777,
        1e-9,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A3")),
        0.00770147248820165,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A4")),
        (1.0_f64 + 0.0525 / 4.0).powi(4) - 1.0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A5")),
        0.0525,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A6")),
        0.1,
        1e-12,
    );
}

#[test]
fn evaluates_cashflow_financial_functions() {
    let mut engine = Engine::new();

    // IRR example from Excel docs.
    let values = [-70_000.0, 12_000.0, 15_000.0, 18_000.0, 21_000.0, 26_000.0];
    for (i, v) in values.iter().enumerate() {
        let addr = format!("A{}", i + 1);
        engine.set_cell_value("Sheet1", &addr, *v).unwrap();
    }
    engine
        .set_cell_formula("Sheet1", "B1", "=IRR(A1:A6)")
        .unwrap();

    // NPV example (discount from period 1).
    engine.set_cell_value("Sheet1", "C1", 10_000.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 15_000.0).unwrap();
    engine.set_cell_value("Sheet1", "C3", 20_000.0).unwrap();
    engine
        .set_cell_formula("Sheet1", "B2", "=NPV(0.1, C1:C3)")
        .unwrap();

    // XNPV / XIRR example from Excel docs.
    let x_values = [-10_000.0, 2_750.0, 4_250.0, 3_250.0, 2_750.0];
    let x_dates = [39_448.0, 39_508.0, 39_751.0, 39_859.0, 39_904.0];
    for (i, (v, d)) in x_values.iter().zip(x_dates.iter()).enumerate() {
        let v_addr = format!("D{}", i + 1);
        let d_addr = format!("E{}", i + 1);
        engine.set_cell_value("Sheet1", &v_addr, *v).unwrap();
        engine.set_cell_value("Sheet1", &d_addr, *d).unwrap();
    }
    engine
        .set_cell_formula("Sheet1", "B3", "=XNPV(0.09, D1:D5, E1:E5)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "B4", "=XIRR(D1:D5, E1:E5)")
        .unwrap();

    // MIRR example from Excel docs.
    let mirr_values = [-120_000.0, 39_000.0, 30_000.0, 21_000.0, 37_000.0, 46_000.0];
    for (i, v) in mirr_values.iter().enumerate() {
        let addr = format!("F{}", i + 1);
        engine.set_cell_value("Sheet1", &addr, *v).unwrap();
    }
    engine
        .set_cell_formula("Sheet1", "B5", "=MIRR(F1:F6, 0.1, 0.12)")
        .unwrap();

    engine.recalculate();

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "B1")),
        0.08663094803653162,
        1e-12,
    );
    let expected_npv = 10_000.0 / 1.1 + 15_000.0 / 1.1_f64.powi(2) + 20_000.0 / 1.1_f64.powi(3);
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "B2")),
        expected_npv,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "B3")),
        2_086.6476020315354,
        1e-10,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "B4")),
        0.3733625335188314,
        1e-12,
    );

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "B5")),
        0.1260941303659051,
        1e-12,
    );
}

#[test]
fn evaluates_amortization_financial_functions() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=CUMIPMT(0.09/12, 30*12, 125000, 13, 24, 0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=CUMPRINC(0.09/12, 30*12, 125000, 13, 24, 0)",
        )
        .unwrap();
    // Coercion: text numbers should be accepted.
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            r#"=CUMIPMT("0.0075","360","125000","13","24","0")"#,
        )
        .unwrap();
    // Non-finite numbers should produce #NUM.
    engine
        .set_cell_value("Sheet1", "B1", f64::INFINITY)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=CUMIPMT(B1, 360, 125000, 13, 24, 0)")
        .unwrap();

    engine.recalculate();

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        -11_135.232130750843,
        1e-10,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A2")),
        -934.107123420897,
        1e-10,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A3")),
        -11_135.232130750843,
        1e-10,
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn xirr_xnpv_length_mismatch_returns_num_error() {
    let mut engine = Engine::new();

    engine.set_cell_value("Sheet1", "A1", -1000.0).unwrap();
    engine.set_cell_value("Sheet1", "A2", 1100.0).unwrap();
    engine.set_cell_value("Sheet1", "B1", 40_000.0).unwrap();

    engine
        .set_cell_formula("Sheet1", "C1", "=XIRR(A1:A2, B1)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "C2", "=XNPV(0.1, A1:A2, B1)")
        .unwrap();

    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "C1"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "C2"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn cashflow_functions_reject_lambda_values() {
    let mut engine = Engine::new();

    // Lambdas are not coercible to numbers for cashflow functions.
    engine
        .set_cell_formula("Sheet1", "A1", "=NPV(0.1, LAMBDA(x,x))")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=IRR(LAMBDA(x,x))")
        .unwrap();

    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::Value)
    );
}
