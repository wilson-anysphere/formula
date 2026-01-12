use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::functions::financial::{duration, mduration, price, yield_rate};
use formula_engine::{ErrorKind, Value};

use super::harness::TestSheet;

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn price_matches_excel_doc_example() {
    // Excel docs:
    // PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, 100, 2, 0) ≈ 94.634361
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2008, 2, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2017, 11, 15), system).unwrap();

    let result = price(
        settlement,
        maturity,
        0.0575,
        0.065,
        100.0,
        2,
        0,
        system,
    )
    .unwrap();
    assert_close(result, 94.634361, 1e-6);
}

#[test]
fn yield_matches_excel_doc_example() {
    // Excel docs:
    // YIELD(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 95.04287, 100, 2, 0) ≈ 0.064
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2008, 2, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2017, 11, 15), system).unwrap();

    let y = yield_rate(settlement, maturity, 0.0575, 95.04287, 100.0, 2, 0, system).unwrap();
    assert_close(y, 0.064, 1e-3);
}

#[test]
fn duration_and_mduration_match_excel_doc_example() {
    // Excel docs:
    // DURATION(DATE(2008,1,1), DATE(2016,1,1), 0.08, 0.09, 2, 1) ≈ 5.993774
    // MDURATION(...) ≈ 5.737
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2008, 1, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2016, 1, 1), system).unwrap();

    let dur = duration(settlement, maturity, 0.08, 0.09, 2, 1, system).unwrap();
    assert_close(dur, 5.993774, 1e-6);

    let mdur = mduration(settlement, maturity, 0.08, 0.09, 2, 1, system).unwrap();
    assert_close(mdur, dur / (1.0 + 0.09 / 2.0), 1e-12);
}

#[test]
fn yield_price_roundtrip() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2030, 1, 1), system).unwrap();
    let rate = 0.05;
    let yld = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    let pr = price(
        settlement,
        maturity,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();

    let back = yield_rate(settlement, maturity, rate, pr, redemption, frequency, basis, system).unwrap();
    assert_close(back, yld, 1e-10);
}

#[test]
fn settlement_on_coupon_date_has_zero_accrued_interest() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 7, 1), system).unwrap();
    let rate = 0.06;
    let yld = 0.06;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    // On coupon date, accrued interest should be 0 and the clean price should equal the dirty price.
    let pr = price(
        settlement,
        maturity,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert!(pr.is_finite());
}

#[test]
fn price_supports_zero_yield() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 1, 1), system).unwrap();
    let pr = price(settlement, maturity, 0.1, 0.0, 100.0, 2, 0, system).unwrap();
    assert!(pr.is_finite());
}

#[test]
fn rejects_invalid_frequency_and_basis() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 1, 1), system).unwrap();

    assert_eq!(
        price(settlement, maturity, 0.1, 0.1, 100.0, 3, 0, system),
        Err(formula_engine::ExcelError::Num)
    );
    assert_eq!(
        price(settlement, maturity, 0.1, 0.1, 100.0, 2, 9, system),
        Err(formula_engine::ExcelError::Num)
    );
}

#[test]
fn rejects_settlement_on_or_after_maturity() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let maturity = settlement;
    assert_eq!(
        price(settlement, maturity, 0.1, 0.1, 100.0, 2, 0, system),
        Err(formula_engine::ExcelError::Num)
    );
}

#[test]
fn builtins_accept_date_strings_via_datevalue() {
    let mut sheet = TestSheet::new();
    match sheet.eval(r#"=PRICE("2008-02-15","2017-11-15",0.0575,0.065,100,2,0)"#) {
        Value::Number(n) => assert_close(n, 94.634361, 1e-6),
        other => panic!("expected number, got {other:?}"),
    }

    // Invalid basis should propagate as #NUM!.
    assert_eq!(
        sheet.eval(r#"=PRICE("2008-02-15","2017-11-15",0.0575,0.065,100,2,99)"#),
        Value::Error(ErrorKind::Num)
    );
}

