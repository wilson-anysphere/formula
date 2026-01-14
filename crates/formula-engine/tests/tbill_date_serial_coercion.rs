use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::functions::financial;
use formula_engine::{Engine, ErrorKind, Value};

fn assert_number(cell: Value) -> f64 {
    match cell {
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
fn tbill_functions_coerce_date_serials_like_excel() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();

    let discount = 0.05;
    let expected_price = financial::tbillprice(settlement, maturity, discount).unwrap();
    let expected_eq = financial::tbilleq(settlement, maturity, discount).unwrap();
    let expected_yield = financial::tbillyield(settlement, maturity, expected_price).unwrap();

    let mut engine = Engine::new();
    engine
        .set_cell_value("Sheet1", "B1", f64::from(settlement) + 0.9)
        .unwrap();
    engine
        .set_cell_value("Sheet1", "B2", f64::from(maturity) + 0.1)
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "A1", "=TBILLPRICE(B1,B2,0.05)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=TBILLEQ(B1,B2,0.05)")
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=TBILLYIELD(B1,B2,A1)")
        .unwrap();

    // Serial must be within i32 range.
    engine
        .set_cell_value("Sheet1", "B3", (i32::MAX as f64) + 1.0)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A4", "=TBILLPRICE(B3,B2,0.05)")
        .unwrap();

    // Non-finite serials should return #NUM!
    engine
        .set_cell_value("Sheet1", "B4", f64::INFINITY)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A5", "=TBILLEQ(B4,B2,0.05)")
        .unwrap();

    engine.recalculate();

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        expected_price,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A2")),
        expected_eq,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A3")),
        expected_yield,
        1e-12,
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A5"),
        Value::Error(ErrorKind::Num)
    );
}
