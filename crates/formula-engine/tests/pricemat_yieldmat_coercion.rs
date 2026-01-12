use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
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

fn serial_1900(year: i32, month: u8, day: u8) -> i32 {
    ymd_to_serial(ExcelDate::new(year, month, day), ExcelDateSystem::EXCEL_1900).unwrap()
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
    assert_close(assert_number(engine.get_cell_value("Sheet1", "A3")), 0.04, 1e-12);
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

