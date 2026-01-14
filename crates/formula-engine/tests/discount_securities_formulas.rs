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
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=DISC(DATE(2020,1,1),DATE(2021,1,1),97,100)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "B1",
            "=DISC(\"2020-01-01\",\"2021-01-01\",97,100)",
        )
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
    engine
        .set_cell_formula(
            "Sheet1",
            "A11",
            "=TBILLEQ(DATE(2020,1,1),DATE(2020,7,1),0.05)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A12",
            "=DISC(DATE(2011,1,1),DATE(2011,12,31),97,100,4)",
        )
        .unwrap();

    engine.recalculate();

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        0.03,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "B1")),
        0.03,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A2")),
        95.0,
        1e-12,
    );
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
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A5")),
        100.0,
        1e-12,
    );

    let expected_pricemat = 110.0 / 1.04 - 5.0;
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A6")),
        expected_pricemat,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A7")),
        0.04,
        1e-12,
    );

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

    let expected_tbilleq_short = 365.0 * 0.05 / (360.0 - 0.05 * 182.0);
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A11")),
        expected_tbilleq_short,
        1e-12,
    );

    // Basis 4 (European 30/360) yields YEARFRAC == 359/360 between 2011-01-01 and 2011-12-31.
    let expected_disc_basis4 = (100.0 - 97.0) / 100.0 / (359.0 / 360.0);
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A12")),
        expected_disc_basis4,
        1e-12,
    );
}

#[test]
fn discount_security_functions_validate_dates_and_basis() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=DISC(DATE(2020,1,1),DATE(2020,1,1),97,100)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=DISC(DATE(2020,1,1),DATE(2021,1,1),97,100,5)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            "=DISC(\"not a date\",DATE(2021,1,1),97,100)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            "=RECEIVED(DATE(2020,1,1),DATE(2021,1,1),100,1)",
        )
        .unwrap();

    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A3"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn discount_security_functions_validate_additional_constraints() {
    let mut engine = Engine::new();

    // pr must be > 0
    engine
        .set_cell_formula("Sheet1", "A1", "=DISC(DATE(2020,1,1),DATE(2021,1,1),0,100)")
        .unwrap();

    // Discount too large -> non-positive price => #NUM!
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),2,100)",
        )
        .unwrap();

    // RECEIVED denominator < 0 => #NUM!
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            "=RECEIVED(DATE(2020,1,1),DATE(2021,1,1),100,2)",
        )
        .unwrap();

    // TBILL* dsm must be <= 365
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            "=TBILLPRICE(DATE(2020,1,1),DATE(2021,1,1),0.05)",
        )
        .unwrap();

    // PRICEMAT/YIELDMAT issue must be <= settlement
    engine
        .set_cell_formula(
            "Sheet1",
            "A5",
            "=PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2020,6,1),0.05,0.04)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A6",
            "=YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2020,6,1),0.05,100)",
        )
        .unwrap();

    // TBILL* price must be positive (#NUM! if discount implies non-positive price).
    engine
        .set_cell_formula(
            "Sheet1",
            "A7",
            "=TBILLPRICE(DATE(2020,1,1),DATE(2020,12,31),1)",
        )
        .unwrap();

    // TBILLEQ requires a positive price factor as well.
    engine
        .set_cell_formula(
            "Sheet1",
            "A8",
            "=TBILLEQ(DATE(2020,1,1),DATE(2020,12,31),1)",
        )
        .unwrap();

    // TBILLYIELD requires a strictly positive price.
    engine
        .set_cell_formula(
            "Sheet1",
            "A9",
            "=TBILLYIELD(DATE(2020,1,1),DATE(2020,7,1),0)",
        )
        .unwrap();

    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A3"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A4"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A5"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A6"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A7"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A8"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A9"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn discount_security_functions_reject_non_finite_numbers() {
    let mut engine = Engine::new();

    engine
        .set_cell_value("Sheet1", "B1", f64::INFINITY)
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=DISC(DATE(2020,1,1),DATE(2021,1,1),B1,100)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=TBILLYIELD(DATE(2020,1,1),DATE(2020,7,1),B1)",
        )
        .unwrap();

    engine.recalculate();

    assert_eq!(
        engine.get_cell_value("Sheet1", "A1"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn discount_security_functions_coerce_basis_like_excel() {
    use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
    use formula_engine::functions::financial;

    let system = ExcelDateSystem::EXCEL_1900;

    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let pr = 97.0;
    let redemption = 100.0;

    let expected_basis0 = financial::disc(settlement, maturity, pr, redemption, 0, system).unwrap();
    let expected_basis1 = financial::disc(settlement, maturity, pr, redemption, 1, system).unwrap();
    let expected_basis2 = financial::disc(settlement, maturity, pr, redemption, 2, system).unwrap();
    let expected_basis3 = financial::disc(settlement, maturity, pr, redemption, 3, system).unwrap();
    let expected_basis4 = financial::disc(settlement, maturity, pr, redemption, 4, system).unwrap();

    let mut engine = Engine::new();

    // Reference basis values through cells to exercise coercion rules.
    engine.set_cell_value("Sheet1", "B1", "2").unwrap(); // text -> number 2
    engine.set_cell_value("Sheet1", "B2", true).unwrap(); // TRUE -> 1
    engine.set_cell_value("Sheet1", "B3", false).unwrap(); // FALSE -> 0
                                                           // B4 intentionally left blank (missing cell) -> blank -> 0
    engine.set_cell_value("Sheet1", "B5", 4.9).unwrap(); // trunc -> 4

    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=DISC(DATE(2020,1,1),DATE(2020,7,1),97,100)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=DISC(DATE(2020,1,1),DATE(2020,7,1),97,100,)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A3",
            "=DISC(DATE(2020,1,1),DATE(2020,7,1),97,100,B4)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A4",
            "=DISC(DATE(2020,1,1),DATE(2020,7,1),97,100,B3)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A5",
            "=DISC(DATE(2020,1,1),DATE(2020,7,1),97,100,B2)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A6",
            "=DISC(DATE(2020,1,1),DATE(2020,7,1),97,100,B1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A7",
            "=DISC(DATE(2020,1,1),DATE(2020,7,1),97,100,\"3\")",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A8",
            "=DISC(DATE(2020,1,1),DATE(2020,7,1),97,100,B5)",
        )
        .unwrap();

    engine
        .set_cell_formula(
            "Sheet1",
            "A9",
            "=DISC(DATE(2020,1,1),DATE(2020,7,1),97,100,\"5\")",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A10",
            "=DISC(DATE(2020,1,1),DATE(2020,7,1),97,100,\"nope\")",
        )
        .unwrap();

    engine.recalculate();

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        expected_basis0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A2")),
        expected_basis0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A3")),
        expected_basis0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A4")),
        expected_basis0,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A5")),
        expected_basis1,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A6")),
        expected_basis2,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A7")),
        expected_basis3,
        1e-12,
    );
    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A8")),
        expected_basis4,
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
}

#[test]
fn discount_security_functions_coerce_date_serials_like_excel() {
    use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
    use formula_engine::functions::financial;

    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();

    let pr = 97.0;
    let redemption = 100.0;
    let basis = 0;
    let expected = financial::disc(settlement, maturity, pr, redemption, basis, system).unwrap();

    let mut engine = Engine::new();

    // Excel ignores any time components in date serials for these security functions. We model
    // that by flooring serials (e.g. 43831.9 -> 43831).
    engine
        .set_cell_value("Sheet1", "B1", f64::from(settlement) + 0.9)
        .unwrap();
    engine
        .set_cell_value("Sheet1", "B2", f64::from(maturity) + 0.1)
        .unwrap();

    engine
        .set_cell_formula("Sheet1", "A1", "=DISC(B1,B2,97,100)")
        .unwrap();

    // Serial must be within i32 range.
    engine
        .set_cell_value("Sheet1", "B3", (i32::MAX as f64) + 1.0)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A2", "=DISC(B3,B2,97,100)")
        .unwrap();

    // Non-finite serials should return #NUM!
    engine
        .set_cell_value("Sheet1", "B4", f64::INFINITY)
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=DISC(B4,B2,97,100)")
        .unwrap();

    engine.recalculate();

    assert_close(
        assert_number(engine.get_cell_value("Sheet1", "A1")),
        expected,
        1e-12,
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A2"),
        Value::Error(ErrorKind::Num)
    );
    assert_eq!(
        engine.get_cell_value("Sheet1", "A3"),
        Value::Error(ErrorKind::Num)
    );
}

#[test]
fn discount_security_basis_variants_produce_different_results_on_leap_year_interval() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=DISC(DATE(2024,1,1),DATE(2024,7,2),97,100,0)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=DISC(DATE(2024,1,1),DATE(2024,7,2),97,100,1)",
        )
        .unwrap();

    engine.recalculate();

    let basis_0 = assert_number(engine.get_cell_value("Sheet1", "A1"));
    let basis_1 = assert_number(engine.get_cell_value("Sheet1", "A2"));
    assert!(
        (basis_0 - basis_1).abs() > 1e-12,
        "expected basis variants to differ; got basis0={basis_0}, basis1={basis_1}"
    );
}

#[test]
fn discount_security_date_text_inputs_work_for_iso_and_slash_formats() {
    let mut engine = Engine::new();

    engine
        .set_cell_formula(
            "Sheet1",
            "A1",
            "=DISC(DATE(2024,1,1),DATE(2024,7,2),97,100,1)",
        )
        .unwrap();
    engine
        .set_cell_formula(
            "Sheet1",
            "A2",
            "=DISC(\"2024-01-01\",\"2024-07-02\",97,100,1)",
        )
        .unwrap();
    engine
        .set_cell_formula("Sheet1", "A3", "=DISC(\"1/1/2024\",\"7/2/2024\",97,100,1)")
        .unwrap();

    engine.recalculate();

    let expected = assert_number(engine.get_cell_value("Sheet1", "A1"));
    let iso = assert_number(engine.get_cell_value("Sheet1", "A2"));
    let slash = assert_number(engine.get_cell_value("Sheet1", "A3"));

    assert_close(iso, expected, 1e-12);
    assert_close(slash, expected, 1e-12);
}
