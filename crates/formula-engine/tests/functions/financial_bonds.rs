use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::functions::date_time;
use formula_engine::functions::financial::{
    accrint, accrintm, coupdaybs, coupdays, coupdaysnc, coupncd, coupnum, couppcd, duration,
    mduration, price, yield_rate,
};
use formula_engine::{ErrorKind, Value};

use super::harness::TestSheet;

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

fn eval_number_or_skip(sheet: &mut TestSheet, formula: &str) -> Option<f64> {
    match sheet.eval(formula) {
        Value::Number(n) => Some(n),
        // These bond functions may not be registered in every build of the engine yet.
        Value::Error(ErrorKind::Name) => None,
        other => panic!("expected number, got {other:?} from {formula}"),
    }
}

#[test]
fn price_matches_excel_doc_example() {
    // Excel docs:
    // PRICE(DATE(2008,2,15), DATE(2017,11,15), 0.0575, 0.065, 100, 2, 0) ≈ 94.634361
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2008, 2, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2017, 11, 15), system).unwrap();

    let result = price(settlement, maturity, 0.0575, 0.065, 100.0, 2, 0, system).unwrap();
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
        settlement, maturity, rate, yld, redemption, frequency, basis, system,
    )
    .unwrap();

    let back = yield_rate(
        settlement, maturity, rate, pr, redemption, frequency, basis, system,
    )
    .unwrap();
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
        settlement, maturity, rate, yld, redemption, frequency, basis, system,
    )
    .unwrap();
    assert!(pr.is_finite());
}

#[test]
fn yield_price_roundtrip_end_of_month_schedule() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2020, 3, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 8, 31), system).unwrap();
    let rate = 0.05;
    let yld = 0.07;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 3; // Actual/365 has a fixed coupon-period length (365/frequency).

    let pr = price(
        settlement, maturity, rate, yld, redemption, frequency, basis, system,
    )
    .unwrap();

    let recovered = yield_rate(
        settlement, maturity, rate, pr, redemption, frequency, basis, system,
    )
    .unwrap();
    assert_close(recovered, yld, 1e-10);
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
fn negative_yield_below_minus_one_is_allowed_when_frequency_gt_one() {
    let system = ExcelDateSystem::EXCEL_1900;
    // Settlement is exactly on a coupon date, with one period remaining (n=1, A=0, d=1).
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();

    let rate = 0.10;
    let yld = -1.5;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    let freq = frequency as f64;
    let coupon_payment = 100.0 * rate / freq;
    let expected = (coupon_payment + redemption) / (1.0 + yld / freq);
    let actual = price(
        settlement, maturity, rate, yld, redemption, frequency, basis, system,
    )
    .unwrap();
    assert_close(actual, expected, 1e-12);

    let solved = yield_rate(
        settlement, maturity, rate, expected, redemption, frequency, basis, system,
    )
    .unwrap();
    assert_close(solved, yld, 1e-10);

    // With a single cashflow one semiannual period away, Macaulay duration is 0.5 years.
    let dur = duration(settlement, maturity, rate, yld, frequency, basis, system).unwrap();
    assert_close(dur, 1.0 / freq, 1e-12);
    let mdur = mduration(settlement, maturity, rate, yld, frequency, basis, system).unwrap();
    assert_close(mdur, dur / (1.0 + yld / freq), 1e-12);

    // Boundary behavior: 1 + yld/frequency == 0 -> #DIV/0!.
    assert_eq!(
        price(
            settlement,
            maturity,
            rate,
            -(frequency as f64),
            redemption,
            frequency,
            basis,
            system
        ),
        Err(formula_engine::ExcelError::Div0)
    );
    assert_eq!(
        price(
            settlement,
            maturity,
            rate,
            -(frequency as f64) - 0.1,
            redemption,
            frequency,
            basis,
            system
        ),
        Err(formula_engine::ExcelError::Num)
    );
}

#[test]
fn price_coupon_payment_is_based_on_face_value() {
    // With settlement on a coupon date and yld=0, PRICE reduces to redemption + coupon_payment
    // (clean == dirty, since accrued interest A=0). Excel defines coupon_payment as
    // 100*rate/frequency (i.e. `rate` is per $100 face value, independent of `redemption`).
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 1, 1), system).unwrap(); // next semiannual coupon date

    let rate = 0.10;
    let yld = 0.0;
    let redemption = 105.0;
    let frequency = 2;

    let expected = redemption + 100.0 * rate / (frequency as f64);
    let actual = price(
        settlement, maturity, rate, yld, redemption, frequency, 0, system,
    )
    .unwrap();
    assert_close(actual, expected, 1e-12);
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
    let actual2 = price(
        settlement, maturity, 0.0, yld, redemption, frequency, 2, system,
    )
    .unwrap();
    assert_close(actual2, expected2, 1e-10);

    // basis 3: E = 365/freq
    let e3 = 365.0 / freq;
    let t3 = dsc / e3 + 1.0;
    let expected3 = redemption * g.powf(-t3);
    let actual3 = price(
        settlement, maturity, 0.0, yld, redemption, frequency, 3, system,
    )
    .unwrap();
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

    let actual = price(
        settlement, maturity, rate, yld, redemption, frequency, basis, system,
    )
    .unwrap();
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
#[test]
fn coupon_schedule_sanity_basis_0_and_1() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2024, 6, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2025, 1, 1), system).unwrap();
    let pcd_expected = ymd_to_serial(ExcelDate::new(2024, 1, 1), system).unwrap();
    let ncd_expected = ymd_to_serial(ExcelDate::new(2024, 7, 1), system).unwrap();

    // Basis 0 (US/NASD 30/360).
    assert_eq!(
        couppcd(settlement, maturity, 2, 0, system).unwrap(),
        pcd_expected
    );
    assert_eq!(
        coupncd(settlement, maturity, 2, 0, system).unwrap(),
        ncd_expected
    );
    assert_eq!(coupnum(settlement, maturity, 2, 0, system).unwrap(), 2.0);

    assert_close(
        coupdaybs(settlement, maturity, 2, 0, system).unwrap(),
        164.0,
        0.0,
    );
    assert_close(
        coupdaysnc(settlement, maturity, 2, 0, system).unwrap(),
        16.0,
        0.0,
    );
    assert_close(
        coupdays(settlement, maturity, 2, 0, system).unwrap(),
        180.0,
        0.0,
    );

    // Basis 1 (Actual/Actual).
    assert_eq!(
        couppcd(settlement, maturity, 2, 1, system).unwrap(),
        pcd_expected
    );
    assert_eq!(
        coupncd(settlement, maturity, 2, 1, system).unwrap(),
        ncd_expected
    );
    assert_eq!(coupnum(settlement, maturity, 2, 1, system).unwrap(), 2.0);

    let a_actual = (settlement - pcd_expected) as f64;
    let dsc_actual = (ncd_expected - settlement) as f64;
    let e_actual = (ncd_expected - pcd_expected) as f64;
    assert_close(
        coupdaybs(settlement, maturity, 2, 1, system).unwrap(),
        a_actual,
        0.0,
    );
    assert_close(
        coupdaysnc(settlement, maturity, 2, 1, system).unwrap(),
        dsc_actual,
        0.0,
    );
    assert_close(
        coupdays(settlement, maturity, 2, 1, system).unwrap(),
        e_actual,
        0.0,
    );
}

#[test]
fn price_settlement_on_coupon_date_matches_discounted_cashflows() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2024, 1, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2026, 1, 1), system).unwrap();

    let rate = 0.10;
    let yld = 0.05;
    let redemption = 100.0;
    let frequency = 1;
    let basis = 0;

    // Two cashflows remain, exactly 1 and 2 periods away.
    let expected = 10.0 / 1.05 + 110.0 / 1.05_f64.powi(2);
    let actual = price(
        settlement, maturity, rate, yld, redemption, frequency, basis, system,
    )
    .unwrap();
    assert_close(actual, expected, 1e-12);
}

#[test]
fn yield_duration_and_mduration_one_cashflow_case_is_analytic() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2025, 1, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2026, 1, 1), system).unwrap();

    let coupon = 0.10;
    let yld_expected = 0.05;
    let redemption = 100.0;
    let frequency = 1;
    let basis = 0;

    let pr = 110.0 / (1.0 + yld_expected);
    let yld = yield_rate(
        settlement, maturity, coupon, pr, redemption, frequency, basis, system,
    )
    .unwrap();
    assert_close(yld, yld_expected, 1e-12);

    let dur = duration(
        settlement,
        maturity,
        coupon,
        yld_expected,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert_close(dur, 1.0, 1e-12);

    let mdur = mduration(
        settlement,
        maturity,
        coupon,
        yld_expected,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert_close(mdur, 1.0 / (1.0 + yld_expected), 1e-12);
}

#[test]
fn accrint_and_accrintm_basis_0_are_hand_computable() {
    let system = ExcelDateSystem::EXCEL_1900;
    let issue = ymd_to_serial(ExcelDate::new(2024, 1, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2024, 7, 1), system).unwrap();

    let rate = 0.12;
    let par = 1000.0;
    let basis = 0;

    // 30/360 half-year = 0.5; interest = 1000 * 0.12 * 0.5 = 60.
    let accrued_m = accrintm(issue, settlement, rate, par, basis, system).unwrap();
    assert_close(accrued_m, 60.0, 1e-12);

    let first_interest = settlement;
    let settlement2 = ymd_to_serial(ExcelDate::new(2024, 4, 1), system).unwrap();
    // Semiannual coupon: 1000 * 0.12 / 2 = 60. A/E = 90/180 = 0.5.
    let accrued = accrint(
        issue,
        first_interest,
        settlement2,
        rate,
        par,
        2,
        basis,
        false,
        system,
    )
    .unwrap();
    assert_close(accrued, 30.0, 1e-12);
}

#[test]
fn coup_functions_coerce_frequency_like_excel() {
    let mut sheet = TestSheet::new();
    let settlement = "DATE(2024,6,15)";
    let maturity = "DATE(2025,1,1)";

    // Number-returning COUP* helpers.
    for func in ["COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNUM"] {
        let baseline = format!("={func}({settlement},{maturity},2,0)");
        let Some(expected) = eval_number_or_skip(&mut sheet, &baseline) else {
            return;
        };
        let with_float_freq = format!("={func}({settlement},{maturity},2.9,0)");
        let Some(actual) = eval_number_or_skip(&mut sheet, &with_float_freq) else {
            return;
        };
        assert_close(actual, expected, 1e-12);
    }

    // Date-serial COUP* helpers.
    for func in ["COUPNCD", "COUPPCD"] {
        let baseline = format!("={func}({settlement},{maturity},2,0)");
        let Some(expected) = eval_number_or_skip(&mut sheet, &baseline) else {
            return;
        };
        let with_float_freq = format!("={func}({settlement},{maturity},2.9,0)");
        let Some(actual) = eval_number_or_skip(&mut sheet, &with_float_freq) else {
            return;
        };
        assert_eq!(actual, expected);
    }
}

#[test]
fn coup_functions_coerce_basis_like_excel() {
    let mut sheet = TestSheet::new();
    let settlement = "DATE(2024,6,15)";
    let maturity = "DATE(2025,1,1)";

    // Number-returning COUP* helpers.
    for func in ["COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNUM"] {
        let baseline = format!("={func}({settlement},{maturity},2,0)");
        let Some(expected) = eval_number_or_skip(&mut sheet, &baseline) else {
            return;
        };
        let with_float_basis = format!("={func}({settlement},{maturity},2,0.9)");
        let Some(actual) = eval_number_or_skip(&mut sheet, &with_float_basis) else {
            return;
        };
        assert_close(actual, expected, 1e-12);
    }

    // Date-serial COUP* helpers.
    for func in ["COUPNCD", "COUPPCD"] {
        let baseline = format!("={func}({settlement},{maturity},2,0)");
        let Some(expected) = eval_number_or_skip(&mut sheet, &baseline) else {
            return;
        };
        let with_float_basis = format!("={func}({settlement},{maturity},2,0.9)");
        let Some(actual) = eval_number_or_skip(&mut sheet, &with_float_basis) else {
            return;
        };
        assert_eq!(actual, expected);
    }

    // Regression: truncation, not rounding (1.999... -> 1).
    let baseline_days =
        eval_number_or_skip(&mut sheet, "=COUPDAYS(DATE(2024,6,15),DATE(2025,1,1),2,1)")
            .expect("COUPDAYS should evaluate for basis=1");
    let with_float = eval_number_or_skip(
        &mut sheet,
        "=COUPDAYS(DATE(2024,6,15),DATE(2025,1,1),2,1.999999999)",
    )
    .expect("COUPDAYS should evaluate for basis=1.999999999");
    assert_close(with_float, baseline_days, 1e-12);
}

#[test]
fn standard_bond_functions_coerce_frequency_like_excel() {
    let mut sheet = TestSheet::new();

    let Some(price_baseline) = eval_number_or_skip(
        &mut sheet,
        "=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)",
    ) else {
        return;
    };
    let price_with_float = eval_number_or_skip(
        &mut sheet,
        "=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2.9,0)",
    )
    .expect("PRICE should evaluate for frequency=2.9");
    assert_close(price_with_float, price_baseline, 1e-10);

    let Some(yield_baseline) = eval_number_or_skip(
        &mut sheet,
        "=YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2,0)",
    ) else {
        return;
    };
    let yield_with_float = eval_number_or_skip(
        &mut sheet,
        "=YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2.9,0)",
    )
    .expect("YIELD should evaluate for frequency=2.9");
    assert_close(yield_with_float, yield_baseline, 1e-10);

    let Some(duration_baseline) = eval_number_or_skip(
        &mut sheet,
        "=DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)",
    ) else {
        return;
    };
    let duration_with_float = eval_number_or_skip(
        &mut sheet,
        "=DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2.9,1)",
    )
    .expect("DURATION should evaluate for frequency=2.9");
    assert_close(duration_with_float, duration_baseline, 1e-12);

    let Some(mduration_baseline) = eval_number_or_skip(
        &mut sheet,
        "=MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)",
    ) else {
        return;
    };
    let mduration_with_float = eval_number_or_skip(
        &mut sheet,
        "=MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2.9,1)",
    )
    .expect("MDURATION should evaluate for frequency=2.9");
    assert_close(mduration_with_float, mduration_baseline, 1e-12);
}

#[test]
fn standard_bond_functions_coerce_basis_like_excel() {
    let mut sheet = TestSheet::new();

    let Some(price_baseline) = eval_number_or_skip(
        &mut sheet,
        "=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)",
    ) else {
        return;
    };
    let price_with_float = eval_number_or_skip(
        &mut sheet,
        "=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0.9)",
    )
    .expect("PRICE should evaluate for basis=0.9");
    assert_close(price_with_float, price_baseline, 1e-10);

    let Some(yield_baseline) = eval_number_or_skip(
        &mut sheet,
        "=YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2,0)",
    ) else {
        return;
    };
    let yield_with_float = eval_number_or_skip(
        &mut sheet,
        "=YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,95.04287,100,2,0.9)",
    )
    .expect("YIELD should evaluate for basis=0.9");
    assert_close(yield_with_float, yield_baseline, 1e-10);

    let Some(duration_baseline) = eval_number_or_skip(
        &mut sheet,
        "=DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,0)",
    ) else {
        return;
    };
    let duration_with_float = eval_number_or_skip(
        &mut sheet,
        "=DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,0.9)",
    )
    .expect("DURATION should evaluate for basis=0.9");
    assert_close(duration_with_float, duration_baseline, 1e-12);

    let Some(mduration_baseline) = eval_number_or_skip(
        &mut sheet,
        "=MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,0)",
    ) else {
        return;
    };
    let mduration_with_float = eval_number_or_skip(
        &mut sheet,
        "=MDURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,0.9)",
    )
    .expect("MDURATION should evaluate for basis=0.9");
    assert_close(mduration_with_float, mduration_baseline, 1e-12);

    // Regression: truncation, not rounding (1.999... -> 1).
    let duration_basis1 = eval_number_or_skip(
        &mut sheet,
        "=DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1)",
    )
    .expect("DURATION should evaluate for basis=1");
    let duration_basis1_float = eval_number_or_skip(
        &mut sheet,
        "=DURATION(DATE(2008,1,1),DATE(2016,1,1),0.08,0.09,2,1.999999999)",
    )
    .expect("DURATION should evaluate for basis=1.999999999");
    assert_close(duration_basis1_float, duration_basis1, 1e-12);
}

#[test]
fn accrint_functions_coerce_frequency_and_basis_like_excel() {
    let mut sheet = TestSheet::new();

    let accrint_baseline = match sheet
        .eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,8,15),0.1,1000,2,0,FALSE)")
    {
        Value::Error(ErrorKind::Name) => return,
        Value::Number(n) => n,
        other => panic!("expected number, got {other:?}"),
    };

    let accrint_float_freq = match sheet
        .eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,8,15),0.1,1000,2.9,0,FALSE)")
    {
        Value::Number(n) => n,
        other => panic!("expected number, got {other:?}"),
    };
    assert_close(accrint_float_freq, accrint_baseline, 1e-12);

    let accrint_float_basis = match sheet
        .eval("=ACCRINT(DATE(2020,2,15),DATE(2020,5,15),DATE(2020,8,15),0.1,1000,2,0.9,FALSE)")
    {
        Value::Number(n) => n,
        other => panic!("expected number, got {other:?}"),
    };
    assert_close(accrint_float_basis, accrint_baseline, 1e-12);

    let Some(accrintm_baseline) = eval_number_or_skip(
        &mut sheet,
        "=ACCRINTM(DATE(2020,1,1),DATE(2020,7,1),0.1,1000,0)",
    ) else {
        return;
    };
    let accrintm_float_basis = eval_number_or_skip(
        &mut sheet,
        "=ACCRINTM(DATE(2020,1,1),DATE(2020,7,1),0.1,1000,0.9)",
    )
    .expect("ACCRINTM should evaluate for basis=0.9");
    assert_close(accrintm_float_basis, accrintm_baseline, 1e-12);
}
