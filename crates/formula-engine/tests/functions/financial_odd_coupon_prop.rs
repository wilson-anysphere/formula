#![cfg(not(target_arch = "wasm32"))]

use super::harness::TestSheet;
use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::functions::date_time::edate;
use formula_engine::{ErrorKind, Value};
use proptest::prelude::*;
use proptest::test_runner::{Config, RngAlgorithm, TestRng, TestRunner};

const SYSTEM: ExcelDateSystem = ExcelDateSystem::EXCEL_1900;
const BASIS: i32 = 0;
const REDEMPTION: f64 = 100.0;
const TOLERANCE: f64 = 1e-7;
const CASES: u32 = 64;

const ODDF_SEED: [u8; 32] = [0x6f; 32]; // "o" (odd-first)
const ODDL_SEED: [u8; 32] = [0x6c; 32]; // "l" (odd-last)

#[derive(Debug, Clone)]
struct OddFirstCase {
    settlement: i32,
    maturity: i32,
    issue: i32,
    first_coupon: i32,
    rate: f64,
    yld: f64,
    frequency: i32,
}

#[derive(Debug, Clone)]
struct OddLastCase {
    settlement: i32,
    maturity: i32,
    last_interest: i32,
    rate: f64,
    yld: f64,
    frequency: i32,
}

fn arb_frequency() -> impl Strategy<Value = i32> {
    prop_oneof![Just(1), Just(2), Just(4)]
}

fn arb_rate_0_to_0_2() -> impl Strategy<Value = f64> {
    // Use fixed-point micros to keep test inputs deterministic and avoid NaNs/infinities.
    (0u32..=200_000u32).prop_map(|micros| micros as f64 / 1_000_000.0)
}

fn arb_oddf_case() -> impl Strategy<Value = OddFirstCase> {
    arb_frequency().prop_flat_map(|frequency| {
        let months_per_period = 12 / frequency;
        (
            2000i32..=2030,
            1u8..=12,
            1i32..=20,  // n_coupons
            2i32..=120, // issue_offset_days (>=2 so settlement can be strictly between)
            arb_rate_0_to_0_2(),
            arb_rate_0_to_0_2(),
        )
            .prop_flat_map(
                move |(year, month, n_coupons, issue_offset_days, rate, yld)| {
                    let first_coupon =
                        ymd_to_serial(ExcelDate::new(year, month, 15), SYSTEM).unwrap();
                    let maturity = edate(
                        first_coupon,
                        months_per_period * n_coupons,
                        SYSTEM,
                    )
                    .unwrap();

                    let issue = first_coupon - issue_offset_days;

                    // settlement_offset_days ∈ [1, issue_offset_days-1]
                    (1i32..issue_offset_days).prop_map(move |settle_offset_days| OddFirstCase {
                        settlement: issue + settle_offset_days,
                        maturity,
                        issue,
                        first_coupon,
                        rate,
                        yld,
                        frequency,
                    })
                },
            )
    })
}

fn arb_oddl_case() -> impl Strategy<Value = OddLastCase> {
    arb_frequency().prop_flat_map(|frequency| {
        let months_per_period = 12 / frequency;
        (
            2000i32..=2030,
            1u8..=12,
            arb_rate_0_to_0_2(),
            arb_rate_0_to_0_2(),
        )
            .prop_flat_map(move |(year, month, rate, yld)| {
                let last_interest = ymd_to_serial(ExcelDate::new(year, month, 15), SYSTEM).unwrap();

                // Ensure maturity falls inside the next regular coupon period (short stub)
                // to avoid edge cases around long stubs and schedule ambiguity.
                let next_coupon = edate(last_interest, months_per_period, SYSTEM).unwrap();
                let period_days = next_coupon - last_interest;

                // maturity_offset_days ∈ [2, period_days-1]
                (2i32..period_days).prop_flat_map(move |maturity_offset_days| {
                    let maturity = last_interest + maturity_offset_days;

                    // settlement_offset_days ∈ [1, maturity_offset_days-1]
                    (1i32..maturity_offset_days).prop_map(move |settle_offset_days| OddLastCase {
                        settlement: last_interest + settle_offset_days,
                        maturity,
                        last_interest,
                        rate,
                        yld,
                        frequency,
                    })
                })
            })
    })
}

fn oddf_available(sheet: &mut TestSheet) -> bool {
    // Use a fixed, valid-ish input set; if the function isn't registered yet, Excel semantics are
    // to return #NAME?.
    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 1), SYSTEM).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 2, 1), SYSTEM).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 15), SYSTEM).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2022, 7, 15), SYSTEM).unwrap();

    let formula = format!(
        "=ODDFPRICE({settlement},{maturity},{issue},{first_coupon},0.05,0.05,{REDEMPTION},2,{BASIS})"
    );
    match sheet.eval(&formula) {
        Value::Error(ErrorKind::Name) => false,
        _ => true,
    }
}

fn oddl_available(sheet: &mut TestSheet) -> bool {
    let last_interest = ymd_to_serial(ExcelDate::new(2020, 1, 15), SYSTEM).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 2, 15), SYSTEM).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2020, 3, 15), SYSTEM).unwrap();

    let formula = format!(
        "=ODDLPRICE({settlement},{maturity},{last_interest},0.05,0.05,{REDEMPTION},2,{BASIS})"
    );
    match sheet.eval(&formula) {
        Value::Error(ErrorKind::Name) => false,
        _ => true,
    }
}

fn unwrap_number(v: &Value) -> Option<f64> {
    match v {
        Value::Number(n) => Some(*n),
        _ => None,
    }
}

#[test]
fn prop_oddf_yield_price_roundtrip_basis0() {
    let mut sheet = TestSheet::new();
    sheet.set_date_system(SYSTEM);
    if !oddf_available(&mut sheet) {
        return;
    }
    let sheet = std::cell::RefCell::new(sheet);

    let mut runner = TestRunner::new_with_rng(
        Config {
            cases: CASES,
            ..Config::default()
        },
        TestRng::from_seed(RngAlgorithm::ChaCha, &ODDF_SEED),
    );

    runner
        .run(&arb_oddf_case(), |case| {
            let mut sheet = sheet.borrow_mut();
            let rate = format!("{:.6}", case.rate);
            let yld = format!("{:.6}", case.yld);

            let price_formula = format!(
                "=ODDFPRICE({s},{m},{i},{fc},{rate},{yld},{red},{freq},{basis})",
                s = case.settlement,
                m = case.maturity,
                i = case.issue,
                fc = case.first_coupon,
                red = REDEMPTION,
                freq = case.frequency,
                basis = BASIS
            );

            let price_val = sheet.eval(&price_formula);
            let Some(price) = unwrap_number(&price_val) else {
                prop_assert!(
                    false,
                    "expected ODDFPRICE to return a number, got {price_val:?} (case={case:?})"
                );
                return Ok(());
            };
            prop_assert!(price.is_finite(), "non-finite ODDFPRICE {price} (case={case:?})");

            sheet.set("A1", price);
            let yield_formula = format!(
                "=ODDFYIELD({s},{m},{i},{fc},{rate},A1,{red},{freq},{basis})",
                s = case.settlement,
                m = case.maturity,
                i = case.issue,
                fc = case.first_coupon,
                red = REDEMPTION,
                freq = case.frequency,
                basis = BASIS,
            );

            let yld_val = sheet.eval(&yield_formula);
            let Some(yld_out) = unwrap_number(&yld_val) else {
                prop_assert!(
                    false,
                    "expected ODDFYIELD to return a number, got {yld_val:?} (case={case:?})"
                );
                return Ok(());
            };

            prop_assert!(yld_out.is_finite(), "non-finite ODDFYIELD {yld_out} (case={case:?})");
            prop_assert!(
                (yld_out - case.yld).abs() <= TOLERANCE,
                "ODDF roundtrip failed: yld_in={} yld_out={yld_out} price={price} case={case:?}",
                case.yld
            );
            Ok(())
        })
        .unwrap();
}

#[test]
fn prop_oddl_yield_price_roundtrip_basis0() {
    let mut sheet = TestSheet::new();
    sheet.set_date_system(SYSTEM);
    if !oddl_available(&mut sheet) {
        return;
    }
    let sheet = std::cell::RefCell::new(sheet);

    let mut runner = TestRunner::new_with_rng(
        Config {
            cases: CASES,
            ..Config::default()
        },
        TestRng::from_seed(RngAlgorithm::ChaCha, &ODDL_SEED),
    );

    runner
        .run(&arb_oddl_case(), |case| {
            let mut sheet = sheet.borrow_mut();
            let rate = format!("{:.6}", case.rate);
            let yld = format!("{:.6}", case.yld);

            let price_formula = format!(
                "=ODDLPRICE({s},{m},{li},{rate},{yld},{red},{freq},{basis})",
                s = case.settlement,
                m = case.maturity,
                li = case.last_interest,
                red = REDEMPTION,
                freq = case.frequency,
                basis = BASIS
            );

            let price_val = sheet.eval(&price_formula);
            let Some(price) = unwrap_number(&price_val) else {
                prop_assert!(
                    false,
                    "expected ODDLPRICE to return a number, got {price_val:?} (case={case:?})"
                );
                return Ok(());
            };
            prop_assert!(price.is_finite(), "non-finite ODDLPRICE {price} (case={case:?})");

            sheet.set("A1", price);
            let yield_formula = format!(
                "=ODDLYIELD({s},{m},{li},{rate},A1,{red},{freq},{basis})",
                s = case.settlement,
                m = case.maturity,
                li = case.last_interest,
                red = REDEMPTION,
                freq = case.frequency,
                basis = BASIS,
            );

            let yld_val = sheet.eval(&yield_formula);
            let Some(yld_out) = unwrap_number(&yld_val) else {
                prop_assert!(
                    false,
                    "expected ODDLYIELD to return a number, got {yld_val:?} (case={case:?})"
                );
                return Ok(());
            };

            prop_assert!(yld_out.is_finite(), "non-finite ODDLYIELD {yld_out} (case={case:?})");
            prop_assert!(
                (yld_out - case.yld).abs() <= TOLERANCE,
                "ODDL roundtrip failed: yld_in={} yld_out={yld_out} price={price} case={case:?}",
                case.yld
            );
            Ok(())
        })
        .unwrap();
}
