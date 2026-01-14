#![cfg(not(target_arch = "wasm32"))]

use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::functions::date_time::edate;
use formula_engine::functions::financial::{oddfprice, oddfyield, oddlprice, oddlyield};
use proptest::prelude::*;
use proptest::test_runner::{Config, RngAlgorithm, TestRng, TestRunner};

const SYSTEM: ExcelDateSystem = ExcelDateSystem::EXCEL_1900;
const BASIS: i32 = 0;
const YIELD_TOLERANCE: f64 = 1e-7;
const PRICE_TOLERANCE: f64 = 1e-5;
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
    redemption: f64,
    frequency: i32,
}

#[derive(Debug, Clone)]
struct OddLastCase {
    settlement: i32,
    maturity: i32,
    last_interest: i32,
    rate: f64,
    yld: f64,
    redemption: f64,
    frequency: i32,
}

fn arb_frequency() -> impl Strategy<Value = i32> {
    prop_oneof![Just(1), Just(2), Just(4)]
}

fn arb_rate_0_to_0_2() -> impl Strategy<Value = f64> {
    // Use fixed-point micros to keep test inputs deterministic and avoid NaNs/infinities.
    (0u32..=200_000u32).prop_map(|micros| micros as f64 / 1_000_000.0)
}

fn arb_yld_0_to_0_2() -> impl Strategy<Value = f64> {
    // Use fixed-point micros to keep test inputs deterministic and avoid NaNs/infinities.
    (0u32..=200_000u32).prop_map(|micros| micros as f64 / 1_000_000.0)
}

fn arb_redemption_50_to_150() -> impl Strategy<Value = f64> {
    // Keep redemption near par for solver stability and determinism.
    // Use cents for stable shrinking.
    (5_000u32..=15_000u32).prop_map(|cents| cents as f64 / 100.0)
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
            arb_yld_0_to_0_2(),
            arb_redemption_50_to_150(),
        )
            .prop_flat_map(
                move |(year, month, n_coupons, issue_offset_days, rate, yld, redemption)| {
                    let first_coupon =
                        ymd_to_serial(ExcelDate::new(year, month, 15), SYSTEM).unwrap();
                    let maturity =
                        edate(first_coupon, months_per_period * n_coupons, SYSTEM).unwrap();

                    let issue = first_coupon - issue_offset_days;

                    // settlement_offset_days ∈ [1, issue_offset_days-1]
                    (1i32..issue_offset_days).prop_map(move |settle_offset_days| OddFirstCase {
                        settlement: issue + settle_offset_days,
                        maturity,
                        issue,
                        first_coupon,
                        rate,
                        yld,
                        redemption,
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
            arb_yld_0_to_0_2(),
            arb_redemption_50_to_150(),
        )
            .prop_flat_map(move |(year, month, rate, yld, redemption)| {
                let last_interest = ymd_to_serial(ExcelDate::new(year, month, 15), SYSTEM).unwrap();

                // Ensure maturity falls inside the next regular coupon period (short stub)
                // to avoid edge cases around long stubs and schedule ambiguity.
                let next_coupon = edate(last_interest, months_per_period, SYSTEM).unwrap();
                let period_days = next_coupon - last_interest;

                // maturity_offset_days ∈ [2, min(period_days-1, 120)] (keep cases fast/stable)
                let max_stub_exclusive = period_days.min(121);
                (2i32..max_stub_exclusive).prop_flat_map(move |maturity_offset_days| {
                    let maturity = last_interest + maturity_offset_days;

                    // settlement_offset_days ∈ [1, maturity_offset_days-1]
                    (1i32..maturity_offset_days).prop_map(move |settle_offset_days| OddLastCase {
                        settlement: last_interest + settle_offset_days,
                        maturity,
                        last_interest,
                        rate,
                        yld,
                        redemption,
                        frequency,
                    })
                })
            })
    })
}

#[test]
fn prop_oddf_yield_price_roundtrip_basis0() {
    let mut runner = TestRunner::new_with_rng(
        Config {
            cases: CASES,
            failure_persistence: None,
            ..Config::default()
        },
        TestRng::from_seed(RngAlgorithm::ChaCha, &ODDF_SEED),
    );

    runner
        .run(&arb_oddf_case(), |case| {
            let price = oddfprice(
                case.settlement,
                case.maturity,
                case.issue,
                case.first_coupon,
                case.rate,
                case.yld,
                case.redemption,
                case.frequency,
                BASIS,
                SYSTEM,
            )
            .map_err(|e| TestCaseError::fail(format!("ODDFPRICE errored: {e:?} case={case:?}")))?;
            prop_assert!(price.is_finite(), "non-finite ODDFPRICE {price} (case={case:?})");
            prop_assert!(price > 0.0, "non-positive ODDFPRICE {price} (case={case:?})");

            let yld_out = oddfyield(
                case.settlement,
                case.maturity,
                case.issue,
                case.first_coupon,
                case.rate,
                price,
                case.redemption,
                case.frequency,
                BASIS,
                SYSTEM,
            )
            .map_err(|e| {
                TestCaseError::fail(format!(
                    "ODDFYIELD errored: {e:?} yld_in={} price={price} case={case:?}",
                    case.yld
                ))
            })?;

            prop_assert!(yld_out.is_finite(), "non-finite ODDFYIELD {yld_out} (case={case:?})");
            prop_assert!(
                (yld_out - case.yld).abs() <= YIELD_TOLERANCE,
                "ODDF roundtrip failed: yld_in={} yld_out={yld_out} price={price} case={case:?}",
                case.yld
            );

            // Secondary invariant: pricing at the recovered yield should reproduce the price.
            let price_roundtrip = oddfprice(
                case.settlement,
                case.maturity,
                case.issue,
                case.first_coupon,
                case.rate,
                yld_out,
                case.redemption,
                case.frequency,
                BASIS,
                SYSTEM,
            )
            .map_err(|e| {
                TestCaseError::fail(format!(
                    "ODDFPRICE(yld_out) errored: {e:?} yld_out={yld_out} price={price} case={case:?}"
                ))
            })?;
            prop_assert!(
                price_roundtrip.is_finite(),
                "non-finite ODDFPRICE {price_roundtrip} (case={case:?})"
            );
            prop_assert!(
                price_roundtrip > 0.0,
                "non-positive ODDFPRICE {price_roundtrip} (case={case:?})"
            );
            prop_assert!(
                (price_roundtrip - price).abs() <= PRICE_TOLERANCE,
                "ODDF price roundtrip failed: price_in={price} price_out={price_roundtrip} yld_out={yld_out} case={case:?}",
            );

            // Monotonicity sanity check: higher yield => lower price (for positive cashflows).
            let y_lo = (case.yld - 0.01).max(0.0);
            let y_hi = (case.yld + 0.01).min(0.2);
            if y_hi > y_lo {
                let p_lo = oddfprice(
                    case.settlement,
                    case.maturity,
                    case.issue,
                    case.first_coupon,
                    case.rate,
                    y_lo,
                    case.redemption,
                    case.frequency,
                    BASIS,
                    SYSTEM,
                )
                .map_err(|e| {
                    TestCaseError::fail(format!(
                        "ODDFPRICE(y_lo) errored: {e:?} y_lo={y_lo} case={case:?}"
                    ))
                })?;
                let p_hi = oddfprice(
                    case.settlement,
                    case.maturity,
                    case.issue,
                    case.first_coupon,
                    case.rate,
                    y_hi,
                    case.redemption,
                    case.frequency,
                    BASIS,
                    SYSTEM,
                )
                .map_err(|e| {
                    TestCaseError::fail(format!(
                        "ODDFPRICE(y_hi) errored: {e:?} y_hi={y_hi} case={case:?}"
                    ))
                })?;

                prop_assert!(p_lo.is_finite() && p_hi.is_finite());
                prop_assert!(p_lo > 0.0 && p_hi > 0.0);
                prop_assert!(
                    p_hi <= p_lo + 1e-8,
                    "ODDF monotonicity failed: y_lo={y_lo} p_lo={p_lo} y_hi={y_hi} p_hi={p_hi} case={case:?}",
                );
            }

            Ok(())
        })
        .unwrap();
}

#[test]
fn prop_oddl_yield_price_roundtrip_basis0() {
    let mut runner = TestRunner::new_with_rng(
        Config {
            cases: CASES,
            failure_persistence: None,
            ..Config::default()
        },
        TestRng::from_seed(RngAlgorithm::ChaCha, &ODDL_SEED),
    );

    runner
        .run(&arb_oddl_case(), |case| {
            let price = oddlprice(
                case.settlement,
                case.maturity,
                case.last_interest,
                case.rate,
                case.yld,
                case.redemption,
                case.frequency,
                BASIS,
                SYSTEM,
            )
            .map_err(|e| TestCaseError::fail(format!("ODDLPRICE errored: {e:?} case={case:?}")))?;
            prop_assert!(price.is_finite(), "non-finite ODDLPRICE {price} (case={case:?})");
            prop_assert!(price > 0.0, "non-positive ODDLPRICE {price} (case={case:?})");

            let yld_out = oddlyield(
                case.settlement,
                case.maturity,
                case.last_interest,
                case.rate,
                price,
                case.redemption,
                case.frequency,
                BASIS,
                SYSTEM,
            )
            .map_err(|e| {
                TestCaseError::fail(format!(
                    "ODDLYIELD errored: {e:?} yld_in={} price={price} case={case:?}",
                    case.yld
                ))
            })?;

            prop_assert!(yld_out.is_finite(), "non-finite ODDLYIELD {yld_out} (case={case:?})");
            prop_assert!(
                (yld_out - case.yld).abs() <= YIELD_TOLERANCE,
                "ODDL roundtrip failed: yld_in={} yld_out={yld_out} price={price} case={case:?}",
                case.yld
            );

            let price_roundtrip = oddlprice(
                case.settlement,
                case.maturity,
                case.last_interest,
                case.rate,
                yld_out,
                case.redemption,
                case.frequency,
                BASIS,
                SYSTEM,
            )
            .map_err(|e| {
                TestCaseError::fail(format!(
                    "ODDLPRICE(yld_out) errored: {e:?} yld_out={yld_out} price={price} case={case:?}"
                ))
            })?;
            prop_assert!(
                price_roundtrip.is_finite(),
                "non-finite ODDLPRICE {price_roundtrip} (case={case:?})"
            );
            prop_assert!(
                price_roundtrip > 0.0,
                "non-positive ODDLPRICE {price_roundtrip} (case={case:?})"
            );
            prop_assert!(
                (price_roundtrip - price).abs() <= PRICE_TOLERANCE,
                "ODDL price roundtrip failed: price_in={price} price_out={price_roundtrip} yld_out={yld_out} case={case:?}",
            );

            let y_lo = (case.yld - 0.01).max(0.0);
            let y_hi = (case.yld + 0.01).min(0.2);
            if y_hi > y_lo {
                let p_lo = oddlprice(
                    case.settlement,
                    case.maturity,
                    case.last_interest,
                    case.rate,
                    y_lo,
                    case.redemption,
                    case.frequency,
                    BASIS,
                    SYSTEM,
                )
                .map_err(|e| {
                    TestCaseError::fail(format!(
                        "ODDLPRICE(y_lo) errored: {e:?} y_lo={y_lo} case={case:?}"
                    ))
                })?;
                let p_hi = oddlprice(
                    case.settlement,
                    case.maturity,
                    case.last_interest,
                    case.rate,
                    y_hi,
                    case.redemption,
                    case.frequency,
                    BASIS,
                    SYSTEM,
                )
                .map_err(|e| {
                    TestCaseError::fail(format!(
                        "ODDLPRICE(y_hi) errored: {e:?} y_hi={y_hi} case={case:?}"
                    ))
                })?;

                prop_assert!(p_lo.is_finite() && p_hi.is_finite());
                prop_assert!(p_lo > 0.0 && p_hi > 0.0);
                prop_assert!(
                    p_hi <= p_lo + 1e-8,
                    "ODDL monotonicity failed: y_lo={y_lo} p_lo={p_lo} y_hi={y_hi} p_hi={p_hi} case={case:?}",
                );
            }

            Ok(())
        })
        .unwrap();
}
