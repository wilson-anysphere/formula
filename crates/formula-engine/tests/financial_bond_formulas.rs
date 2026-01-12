use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::{Engine, ErrorKind, Value};

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

fn assert_number(v: Value) -> f64 {
    match v {
        Value::Number(n) => n,
        other => panic!("expected number, got {other:?}"),
    }
}

#[test]
fn evaluates_bond_functions() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=COUPPCD(DATE(2024,6,15), DATE(2025,1,1), 2, 0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=COUPNCD(DATE(2024,6,15), DATE(2025,1,1), 2, 0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            "=COUPNUM(DATE(2024,6,15), DATE(2025,1,1), 2, 0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            "=COUPDAYBS(DATE(2024,6,15), DATE(2025,1,1), 2, 0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A5",
            "=COUPDAYSNC(DATE(2024,6,15), DATE(2025,1,1), 2, 0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A6",
            "=COUPDAYS(DATE(2024,6,15), DATE(2025,1,1), 2, 0)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "B1",
            "=ACCRINTM(DATE(2024,1,1), DATE(2024,7,1), 0.12, 1000, 0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B2",
            "=ACCRINT(DATE(2024,1,1), DATE(2024,7,1), DATE(2024,4,1), 0.12, 1000, 2, 0)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "C1",
            "=PRICE(DATE(2024,1,1), DATE(2026,1,1), 0.10, 0.05, 100, 1, 0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C2",
            "=YIELD(DATE(2025,1,1), DATE(2026,1,1), 0.10, 110/1.05, 100, 1, 0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C3",
            "=DURATION(DATE(2025,1,1), DATE(2026,1,1), 0.10, 0.05, 1, 0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "C4",
            "=MDURATION(DATE(2025,1,1), DATE(2026,1,1), 0.10, 0.05, 1, 0)",
        )
        .unwrap();

    // Error propagation / validation.
    engine
        .set_cell_formula(
            "Sheet1",
            "D1",
            "=COUPNUM(DATE(2024,6,15), DATE(2025,1,1), 3, 0)",
        )
        .unwrap();

    engine.recalculate();

    let system = ExcelDateSystem::EXCEL_1900;
    let pcd_expected = ymd_to_serial(ExcelDate::new(2024, 1, 1), system).unwrap() as f64;
    let ncd_expected = ymd_to_serial(ExcelDate::new(2024, 7, 1), system).unwrap() as f64;

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        pcd_expected,
        0.0,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A2")),
        ncd_expected,
        0.0,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A3")),
        2.0,
        0.0,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A4")),
        164.0,
        0.0,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A5")),
        16.0,
        0.0,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A6")),
        180.0,
        0.0,
    );

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "B1")),
        60.0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "B2")),
        30.0,
        1e-12,
    );

    let expected_price = 10.0 / 1.05 + 110.0 / 1.05_f64.powi(2);
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "C1")),
        expected_price,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "C2")),
        0.05,
        1e-10,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "C3")),
        1.0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "C4")),
        1.0 / 1.05,
        1e-12,
    );

    assert_eq!(
        engine.get_cell_value("Sheet1", "D1"),
        Value::Error(ErrorKind::Num)
    );
}
