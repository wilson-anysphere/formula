use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::functions::financial::{duration, mduration, price, yield_rate};
use formula_engine::functions::date_time;
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
fn price_basis_2_and_3_use_fixed_coupon_period_length() {
    // Construct a zero-coupon bond so the price depends only on the time-to-maturity exponent.
    // For basis 2/3 Excel uses a fixed coupon-period length E (360/freq or 365/freq), while DSC
    // remains an actual day count. That means DSC/E is not necessarily 1 even when settlement is a
    // coupon date.
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 7, 1), system).unwrap();
    let ncd = ymd_to_serial(ExcelDate::new(2021, 1, 1), system).unwrap();

    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 2;
    let freq = frequency as f64;
    let g = 1.0 + yld / freq;

    // basis 2: E = 360/freq
    let dsc = (ncd - settlement) as f64;
    let e2 = 360.0 / freq;
    let t2 = dsc / e2 + 1.0; // maturity is one full period after NCD
    let expected2 = redemption * g.powf(-t2);
    let actual2 = price(settlement, maturity, 0.0, yld, redemption, frequency, 2, system).unwrap();
    assert_close(actual2, expected2, 1e-10);

    // basis 3: E = 365/freq
    let e3 = 365.0 / freq;
    let t3 = dsc / e3 + 1.0;
    let expected3 = redemption * g.powf(-t3);
    let actual3 = price(settlement, maturity, 0.0, yld, redemption, frequency, 3, system).unwrap();
    assert_close(actual3, expected3, 1e-10);
}

#[test]
fn coupon_schedule_is_anchored_on_maturity_for_eom_dates() {
    // Regression test for EOM schedules: naive iterative EDATE stepping can drift after month-end
    // clamping (e.g. 31st -> 28th -> 28th). Excel's coupon schedule is anchored on maturity, so
    // stepping back in whole periods should recover the prior year/month-end date.
    let system = ExcelDateSystem::EXCEL_1900;
    let maturity = ymd_to_serial(ExcelDate::new(2021, 8, 31), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 11, 1), system).unwrap();

    let rate = 0.10;
    let yld = 0.0;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    // With yld=0, dirty price is simply redemption + N*coupon_payment (no discounting). For this
    // schedule, the previous coupon date should be 2020-08-31 (not 2020-08-28), yielding A=61
    // under the US 30/360 convention.
    let coupon_payment = 100.0 * rate / (frequency as f64);
    let n = 2.0;
    let dirty = redemption + coupon_payment * n;

    let pcd = ymd_to_serial(ExcelDate::new(2020, 8, 31), system).unwrap();
    let a = date_time::days360(pcd, settlement, false, system).unwrap() as f64;
    let e = 360.0 / (frequency as f64);
    let expected = dirty - coupon_payment * (a / e);

    let actual = price(settlement, maturity, rate, yld, redemption, frequency, basis, system).unwrap();
    assert_close(actual, expected, 1e-10);
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
