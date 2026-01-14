use formula_engine::{ErrorKind, Value};

use super::harness::TestSheet;

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

fn eval_number(sheet: &mut TestSheet, formula: &str) -> f64 {
    match sheet.eval(formula) {
        Value::Number(n) => n,
        other => panic!("expected number, got {other:?} from {formula}"),
    }
}

pub(super) fn eval_number_or_skip(sheet: &mut TestSheet, formula: &str) -> Option<f64> {
    match sheet.eval(formula) {
        Value::Number(n) => Some(n),
        // These bond functions may not be registered in every build of the engine yet.
        Value::Error(ErrorKind::Name) => None,
        other => panic!("expected number, got {other:?} from {formula}"),
    }
}

fn coupon_date_from_maturity(maturity: &str, months_per_period: i32, periods_back: i32) -> String {
    // Coupon schedules are maturity-anchored.
    //
    // IMPORTANT: Coupon schedules are derived as `EDATE(maturity, -k*m)` rather than by stepping
    // backwards one period at a time. `EDATE` month-stepping is not invertible due to end-of-month
    // clamping, and iteratively stepping can cause the day-of-month to drift (e.g. 31st -> 30th),
    // producing dates that Excel's COUP* helpers do not treat as coupon boundaries.
    //
    // Excel also has an end-of-month (EOM) rule: if `maturity` is the last day of its month,
    // the coupon schedule is pinned to month-end (including leap-year February).
    debug_assert!(months_per_period > 0);
    debug_assert!(periods_back >= 0);
    let months_back = months_per_period
        .checked_mul(periods_back)
        .expect("months_back fits in i32");
    if months_back == 0 {
        maturity.to_string()
    } else {
        format!(
            "IF({maturity}=EOMONTH({maturity},0),EOMONTH(EDATE({maturity},-{months_back}),0),EDATE({maturity},-{months_back}))"
        )
    }
}

#[test]
fn coup_invariants_when_settlement_is_coupon_date() {
    let mut sheet = TestSheet::new();

    // Skip this entire test module if the COUP* functions aren't available yet.
    let check = "=COUPDAYBS(DATE(2020,7,1),DATE(2021,1,1),2,0)";
    if eval_number_or_skip(&mut sheet, check).is_none()
        || eval_number_or_skip(&mut sheet, "=COUPDAYSNC(DATE(2020,7,1),DATE(2021,1,1),2,0)")
            .is_none()
        || eval_number_or_skip(&mut sheet, "=COUPDAYS(DATE(2020,7,1),DATE(2021,1,1),2,0)").is_none()
    {
        return;
    }

    let maturities = [
        "DATE(2030,12,31)",
        "DATE(2031,2,28)",
        "DATE(2030,8,30)",
        "DATE(2030,7,15)",
    ];
    let frequencies = [1, 2, 4];
    let bases = [0, 1, 2, 3, 4];

    for maturity in maturities {
        for &frequency in &frequencies {
            let months_per_period = 12 / frequency;
            for k in 1..=6 {
                let settlement = coupon_date_from_maturity(maturity, months_per_period, k);

                for &basis in &bases {
                    let daybs = eval_number(
                        &mut sheet,
                        &format!("=COUPDAYBS({settlement},{maturity},{frequency},{basis})"),
                    );
                    assert_close(daybs, 0.0, 1e-12);

                    let daysnc = eval_number(
                        &mut sheet,
                        &format!("=COUPDAYSNC({settlement},{maturity},{frequency},{basis})"),
                    );
                    let days = eval_number(
                        &mut sheet,
                        &format!("=COUPDAYS({settlement},{maturity},{frequency},{basis})"),
                    );

                    // For bases 2 (Actual/360) and 3 (Actual/365), Excel treats the coupon-period
                    // length E as a fixed fraction of a 360/365-day year, while DSC remains an
                    // actual day count. That means DSC is not necessarily equal to E even when
                    // settlement is a coupon date.
                    match basis {
                        0 | 1 | 4 => assert_close(daysnc, days, 1e-12),
                        2 => {
                            assert_close(days, 360.0 / (frequency as f64), 1e-12);
                            let ncd = coupon_date_from_maturity(maturity, months_per_period, k - 1);
                            let expected =
                                eval_number(&mut sheet, &format!("=({ncd})-({settlement})"));
                            assert_close(daysnc, expected, 1e-12);
                        }
                        3 => {
                            assert_close(days, 365.0 / (frequency as f64), 1e-12);
                            let ncd = coupon_date_from_maturity(maturity, months_per_period, k - 1);
                            let expected =
                                eval_number(&mut sheet, &format!("=({ncd})-({settlement})"));
                            assert_close(daysnc, expected, 1e-12);
                        }
                        _ => unreachable!(),
                    }

                    if matches!(basis, 0 | 4) {
                        assert_close(daybs + daysnc, days, 1e-12);
                    }
                }
            }
        }
    }
}

#[test]
fn coup_days_additivity_for_30_360_bases() {
    let mut sheet = TestSheet::new();

    // Skip if the COUP* helpers aren't registered yet.
    if eval_number_or_skip(
        &mut sheet,
        "=COUPDAYBS(DATE(2020,7,15),DATE(2021,7,15),2,0)",
    )
    .is_none()
        || eval_number_or_skip(
            &mut sheet,
            "=COUPDAYSNC(DATE(2020,7,15),DATE(2021,7,15),2,0)",
        )
        .is_none()
        || eval_number_or_skip(&mut sheet, "=COUPDAYS(DATE(2020,7,15),DATE(2021,7,15),2,0)")
            .is_none()
    {
        return;
    }

    let maturities = [
        "DATE(2030,12,31)",
        "DATE(2031,2,28)",
        "DATE(2030,8,30)",
        "DATE(2030,7,15)",
    ];
    let frequencies = [1, 2, 4];
    let bases = [0, 4];
    let deltas = [1, 15, 30];

    for maturity in maturities {
        for &frequency in &frequencies {
            let months_per_period = 12 / frequency;

            // Ensure there's room to step back and still have a valid next coupon date.
            for k in 1..=6 {
                let pcd = coupon_date_from_maturity(maturity, months_per_period, k);

                for &delta in &deltas {
                    let settlement = format!("({pcd}+{delta})");

                    for &basis in &bases {
                        let daybs = eval_number(
                            &mut sheet,
                            &format!("=COUPDAYBS({settlement},{maturity},{frequency},{basis})"),
                        );
                        let daysnc = eval_number(
                            &mut sheet,
                            &format!("=COUPDAYSNC({settlement},{maturity},{frequency},{basis})"),
                        );
                        let days = eval_number(
                            &mut sheet,
                            &format!("=COUPDAYS({settlement},{maturity},{frequency},{basis})"),
                        );
                        assert_close(daybs + daysnc, days, 1e-12);
                    }
                }
            }
        }
    }
}

#[test]
fn coup_schedule_roundtrips_when_settlement_is_coupon_date() {
    let mut sheet = TestSheet::new();

    // Skip if the COUP date helpers aren't registered yet.
    if eval_number_or_skip(&mut sheet, "=COUPPCD(DATE(2020,7,1),DATE(2021,1,1),2,0)").is_none()
        || eval_number_or_skip(&mut sheet, "=COUPNCD(DATE(2020,7,1),DATE(2021,1,1),2,0)").is_none()
        || eval_number_or_skip(&mut sheet, "=COUPNUM(DATE(2020,7,1),DATE(2021,1,1),2,0)").is_none()
    {
        return;
    }

    let maturities = [
        "DATE(2030,12,31)",
        "DATE(2031,2,28)",
        "DATE(2030,8,30)",
        "DATE(2030,7,15)",
    ];
    let frequencies = [1, 2, 4];
    let bases = [0, 1, 2, 3, 4];

    for maturity in maturities {
        for &frequency in &frequencies {
            let months_per_period = 12 / frequency;
            for k in 1..=6 {
                let settlement = coupon_date_from_maturity(maturity, months_per_period, k);
                let expected_ncd = coupon_date_from_maturity(maturity, months_per_period, k - 1);

                for &basis in &bases {
                    let pcd = eval_number(
                        &mut sheet,
                        &format!("=COUPPCD({settlement},{maturity},{frequency},{basis})"),
                    );
                    let settlement_serial = eval_number(&mut sheet, &format!("={settlement}"));
                    assert_close(pcd, settlement_serial, 0.0);

                    let ncd = eval_number(
                        &mut sheet,
                        &format!("=COUPNCD({settlement},{maturity},{frequency},{basis})"),
                    );
                    let expected_ncd_serial = eval_number(&mut sheet, &format!("={expected_ncd}"));
                    assert_close(ncd, expected_ncd_serial, 0.0);

                    let n = eval_number(
                        &mut sheet,
                        &format!("=COUPNUM({settlement},{maturity},{frequency},{basis})"),
                    );
                    assert_close(n, k as f64, 0.0);
                }
            }
        }
    }
}

#[test]
fn price_yield_roundtrip_consistency() {
    let mut sheet = TestSheet::new();

    // Skip if PRICE/YIELD aren't registered yet.
    if eval_number_or_skip(
        &mut sheet,
        "=PRICE(DATE(2020,7,1),DATE(2021,1,1),0.05,0.04,100,2,0)",
    )
    .is_none()
        || eval_number_or_skip(
            &mut sheet,
            "=YIELD(DATE(2020,7,1),DATE(2021,1,1),0.05,100,100,2,0)",
        )
        .is_none()
    {
        return;
    }

    let maturities = [
        "DATE(2030,12,31)",
        "DATE(2031,2,28)",
        "DATE(2030,8,30)",
        "DATE(2030,7,15)",
    ];
    let frequencies = [1, 2, 4];
    let bases = [0, 1, 2, 3, 4];
    let rates = [0.03, 0.065];
    let yields = [0.01, 0.045, 0.11];
    let redemptions = [100.0, 105.0];

    for maturity in maturities {
        for &frequency in &frequencies {
            let months_per_period = 12 / frequency;
            for k in 1..=5 {
                let settlement = coupon_date_from_maturity(maturity, months_per_period, k);

                for &basis in &bases {
                    for &rate in &rates {
                        for &yld in &yields {
                            for &redemption in &redemptions {
                                let recovered = eval_number(
                                    &mut sheet,
                                    &format!(
                                        "=LET(pr,PRICE({settlement},{maturity},{rate},{yld},{redemption},{frequency},{basis}),YIELD({settlement},{maturity},{rate},pr,{redemption},{frequency},{basis}))",
                                    ),
                                );
                                assert_close(recovered, yld, 1e-7);
                            }
                        }
                    }
                }
            }
        }
    }
}

#[test]
fn price_matches_pv_when_settlement_is_coupon_date() {
    let mut sheet = TestSheet::new();

    // Skip if PRICE isn't registered yet.
    if eval_number_or_skip(
        &mut sheet,
        "=PRICE(DATE(2020,7,1),DATE(2021,1,1),0.05,0.04,100,2,0)",
    )
    .is_none()
    {
        return;
    }

    // When `settlement` is exactly the previous coupon date (A=0), `PRICE` should reduce to the
    // standard time-value PV of N periods of coupon payments + final redemption.
    //
    // This is a stronger cross-check than PRICE/YIELD roundtripping because it compares against
    // the independent `PV` implementation.
    let maturities = [
        "DATE(2030,12,31)",
        "DATE(2031,2,28)",
        "DATE(2030,8,30)",
        "DATE(2030,7,15)",
    ];
    let frequencies = [1, 2, 4];
    // For bases 2 and 3, `PRICE` does not reduce to an integer-period `PV` when settlement is a
    // coupon date because the coupon-period length `E` is fixed (360/freq or 365/freq) while `DSC`
    // remains an actual day count. Restrict this invariant to bases where settlement aligns with a
    // true period boundary (0, 1, 4).
    let bases = [0, 1, 4];
    let rates = [0.0, 0.03, 0.065];
    let yields = [0.01, 0.045, 0.11];
    let redemptions = [100.0, 105.0];

    for maturity in maturities {
        for &frequency in &frequencies {
            let months_per_period = 12 / frequency;
            for k in 1..=5 {
                let settlement = coupon_date_from_maturity(maturity, months_per_period, k);

                for &basis in &bases {
                    for &rate in &rates {
                        for &yld in &yields {
                            for &redemption in &redemptions {
                                let price = eval_number(
                                    &mut sheet,
                                    &format!(
                                        "=PRICE({settlement},{maturity},{rate},{yld},{redemption},{frequency},{basis})"
                                    ),
                                );

                                let pv = eval_number(
                                    &mut sheet,
                                    &format!(
                                        // Excel's `rate` is defined per $100 face value (not scaled by
                                        // `redemption`, which is the amount repaid per $100 at maturity).
                                        "=LET(n,{k},c,100*({rate})/{frequency},r,({yld})/{frequency},PV(r,n,-c,-{redemption}))"
                                    ),
                                );

                                if (price - pv).abs() > 1e-7 {
                                    let settlement_serial =
                                        eval_number(&mut sheet, &format!("={settlement}"));
                                    panic!(
                                        "PRICE/PV mismatch: maturity={maturity} settlement={settlement} (serial={settlement_serial}) frequency={frequency} basis={basis} rate={rate} yld={yld} redemption={redemption}: expected {pv}, got {price}"
                                    );
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

#[test]
fn duration_n1_equals_time_to_maturity() {
    let mut sheet = TestSheet::new();

    // Skip if DURATION isn't registered yet.
    if eval_number_or_skip(
        &mut sheet,
        "=DURATION(DATE(2020,7,2),DATE(2021,1,1),0.05,0.04,2,0)",
    )
    .is_none()
    {
        return;
    }

    // For N=1 (only one remaining cash flow date), duration collapses to the time in years until
    // maturity:
    //   DURATION = (DSC / E) / frequency
    // This should be independent of coupon and yield (there's a single cash flow).
    let maturities = [
        "DATE(2030,12,31)",
        "DATE(2031,2,28)",
        "DATE(2030,8,30)",
        "DATE(2030,7,15)",
    ];
    let frequencies = [1, 2, 4];
    let bases = [0, 1, 2, 3, 4];
    let deltas = [1, 10, 30];
    let coupons = [0.0, 0.025, 0.08];
    let yields = [0.01, 0.05, 0.12];

    for maturity in maturities {
        for &frequency in &frequencies {
            let months_per_period = 12 / frequency;

            for &delta in &deltas {
                let pcd = coupon_date_from_maturity(maturity, months_per_period, 1);
                let settlement = format!("({pcd}+{delta})");

                for &basis in &bases {
                    // For N=1, NCD == maturity and PCD is the coupon date one period prior.
                    //
                    // Compute expected DURATION using the same day-count definitions as the bond
                    // schedule (`coupon_schedule` in `bonds.rs`):
                    // - basis 0: `E` is modeled as 360/frequency and `DSC = E - A` (with `A`
                    //   computed via US/NASD DAYS360)
                    // - basis 4: `E` is modeled as 360/frequency and `DSC = E - A` (European 30E/360
                    //   day counts for `A`, but a fixed coupon-period length for `E`)
                    // - basis 2: `E = 360/frequency`, `DSC` is an actual day count
                    // - basis 3: `E = 365/frequency`, `DSC` is an actual day count
                    // - basis 1: `E` is the actual length of the coupon period, and `DSC` is an
                    //   actual day count
                    let expected = eval_number(
                        &mut sheet,
                        &format!(
                            "=LET(pcd,{pcd},a,IF({basis}=0,DAYS360(pcd,{settlement},FALSE),IF({basis}=4,DAYS360(pcd,{settlement},TRUE),{settlement}-pcd)),e,IF(OR({basis}=0,{basis}=2,{basis}=4),360/{frequency},IF({basis}=3,365/{frequency},{maturity}-pcd)),dsc,IF(OR({basis}=0,{basis}=4),e-a,{maturity}-{settlement}),(dsc/e)/{frequency})"
                        ),
                    );

                    for &coupon in &coupons {
                        for &yld in &yields {
                            let dur = eval_number(
                                &mut sheet,
                                &format!(
                                    "=DURATION({settlement},{maturity},{coupon},{yld},{frequency},{basis})"
                                ),
                            );
                            assert_close(dur, expected, 1e-7);
                        }
                    }
                }
            }
        }
    }
}

#[test]
fn mduration_matches_duration_identity() {
    let mut sheet = TestSheet::new();

    // Skip if DURATION/MDURATION aren't registered yet.
    if eval_number_or_skip(
        &mut sheet,
        "=DURATION(DATE(2020,7,1),DATE(2021,1,1),0.05,0.04,2,0)",
    )
    .is_none()
        || eval_number_or_skip(
            &mut sheet,
            "=MDURATION(DATE(2020,7,1),DATE(2021,1,1),0.05,0.04,2,0)",
        )
        .is_none()
    {
        return;
    }

    let maturities = ["DATE(2030,12,31)", "DATE(2031,2,28)", "DATE(2030,8,30)"];
    let frequencies = [1, 2, 4];
    let bases = [0, 1, 2, 3, 4];
    let coupons = [0.025, 0.08];
    let yields = [0.02, 0.055, 0.12];

    for maturity in maturities {
        for &frequency in &frequencies {
            let months_per_period = 12 / frequency;
            for k in 1..=6 {
                let settlement = coupon_date_from_maturity(maturity, months_per_period, k);

                for &basis in &bases {
                    for &coupon in &coupons {
                        for &yld in &yields {
                            let dur = eval_number(
                                &mut sheet,
                                &format!(
                                    "=DURATION({settlement},{maturity},{coupon},{yld},{frequency},{basis})"
                                ),
                            );
                            let mdur = eval_number(
                                &mut sheet,
                                &format!(
                                    "=MDURATION({settlement},{maturity},{coupon},{yld},{frequency},{basis})"
                                ),
                            );
                            let expected = dur / (1.0 + yld / (frequency as f64));
                            assert_close(mdur, expected, 1e-7);
                        }
                    }
                }
            }
        }
    }
}
