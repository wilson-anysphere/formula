use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::{Engine, ErrorKind, Value};

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

fn serial_1900(year: i32, month: u8, day: u8) -> i32 {
    ymd_to_serial(
        ExcelDate::new(year, month, day),
        ExcelDateSystem::EXCEL_1900,
    )
    .unwrap()
}

#[test]
fn pricemat_yieldmat_accept_iso_date_text_and_blank_basis() {
    let mut engine = Engine::new();

    // B1 is left blank to exercise optional basis defaulting.
    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=PRICEMAT(\"2020-01-01\",\"2021-01-01\",\"2019-01-01\",0.05,0.04,B1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,0)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            "=YIELDMAT(\"2020-01-01\",\"2021-01-01\",\"2019-01-01\",0.05,A1,B1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,A2,0)",
        )
        .unwrap();

    engine.recalculate();

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        assert_number(engine.get_cell_value("Sheet1", "A2")),
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A3")),
        assert_number(engine.get_cell_value("Sheet1", "A4")),
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A3")),
        0.04,
        1e-12,
    );
}

#[test]
fn pricemat_floors_fractional_date_serials() {
    let settlement = serial_1900(2020, 1, 1);
    let maturity = serial_1900(2021, 1, 1);
    let issue = serial_1900(2019, 1, 1);

    let mut engine = Engine::new();
    engine
        .set_cell_value("Sheet1", "B1", f64::from(settlement) + 0.9)
        .unwrap();
    engine
        .set_cell_value("Sheet1", "B2", f64::from(maturity) + 0.1)
        .unwrap();
    engine
        .set_cell_value("Sheet1", "B3", f64::from(issue) + 0.9)
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "A1", "=PRICEMAT(B1,B2,B3,0.05,0.04)")
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04)",
        )
        .unwrap();

    engine.recalculate();
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        assert_number(engine.get_cell_value("Sheet1", "A2")),
        1e-12,
    );
}

#[test]
fn pricemat_yieldmat_coerce_basis_like_excel() {
    use formula_engine::functions::financial;

    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = serial_1900(2020, 1, 1);
    let maturity = serial_1900(2021, 1, 1);
    let issue = serial_1900(2019, 1, 1);
    let rate = 0.05;
    let yld = 0.04;

    let expected_pricemat_basis0 =
        financial::pricemat(settlement, maturity, issue, rate, yld, 0, system).unwrap();
    let expected_pricemat_basis1 =
        financial::pricemat(settlement, maturity, issue, rate, yld, 1, system).unwrap();
    let expected_pricemat_basis2 =
        financial::pricemat(settlement, maturity, issue, rate, yld, 2, system).unwrap();
    let expected_pricemat_basis3 =
        financial::pricemat(settlement, maturity, issue, rate, yld, 3, system).unwrap();
    let expected_pricemat_basis4 =
        financial::pricemat(settlement, maturity, issue, rate, yld, 4, system).unwrap();

    let mut engine = Engine::new();

    // Reference basis values through cells to exercise coercion rules.
    engine.set_cell_value("Sheet1", "B1", "2").unwrap(); // text -> number 2
    engine.set_cell_value("Sheet1", "B2", true).unwrap(); // TRUE -> 1
    engine.set_cell_value("Sheet1", "B3", false).unwrap(); // FALSE -> 0
                                                           // B4 intentionally left blank -> blank -> 0
    engine.set_cell_value("Sheet1", "B5", 4.9).unwrap(); // trunc -> 4

    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,B4)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,B3)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A5",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,B2)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A6",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,B1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A7",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,\"3\")",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A8",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,B5)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "A9",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,\"5\")",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A10",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,\"nope\")",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "C1",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,A1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C2",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,A2,)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C3",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,A3,B4)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C4",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,A4,B3)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C5",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,A5,B2)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C6",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,A6,B1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C7",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,A7,\"3\")",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C8",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,A8,B5)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C9",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,A1,\"5\")",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C10",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,A1,\"nope\")",
        )
        .unwrap();

    engine.recalculate();

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        expected_pricemat_basis0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A2")),
        expected_pricemat_basis0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A3")),
        expected_pricemat_basis0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A4")),
        expected_pricemat_basis0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A5")),
        expected_pricemat_basis1,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A6")),
        expected_pricemat_basis2,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A7")),
        expected_pricemat_basis3,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A8")),
        expected_pricemat_basis4,
        1e-12,
    );

    assert_eq!(
        engine.get_cell_value("Sheet1", "A9"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A10"),
        Value::Error(ErrorKind::Value)
    );

    // YIELDMAT should invert PRICEMAT (up to floating-point error) for any basis.
    for addr in ["C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8"] {
        assert_close(
            assert_number(engine.get_cell_value("Sheet1", addr)),
            yld,
            1e-12,
        );
    }
    assert_eq!(
        engine.get_cell_value("Sheet1", "C9"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "C10"),
        Value::Error(ErrorKind::Value)
    );
}
