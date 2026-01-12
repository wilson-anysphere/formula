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

fn eval_number_or_skip(sheet: &mut TestSheet, formula: &str) -> Option<f64> {
    match sheet.eval(formula) {
        Value::Number(n) => Some(n),
        // These bond functions may not be registered in every build of the engine yet.
        Value::Error(ErrorKind::Name) => None,
        other => panic!("expected number, got {other:?} from {formula}"),
    }
}

#[test]
fn coup_invariants_when_settlement_is_coupon_date() {
    let mut sheet = TestSheet::new();

    // Skip this entire test module if the COUP* functions aren't available yet.
    let check = "=COUPDAYBS(DATE(2020,7,1),DATE(2021,1,1),2,0)";
    if eval_number_or_skip(&mut sheet, check).is_none()
        || eval_number_or_skip(
            &mut sheet,
            "=COUPDAYSNC(DATE(2020,7,1),DATE(2021,1,1),2,0)",
        )
        .is_none()
        || eval_number_or_skip(&mut sheet, "=COUPDAYS(DATE(2020,7,1),DATE(2021,1,1),2,0)")
            .is_none()
    {
        return;
    }

    let maturities = ["DATE(2030,12,31)", "DATE(2031,2,28)", "DATE(2030,7,15)"];
    let frequencies = [1, 2, 4];
    let bases = [0, 1, 2, 3, 4];

    for maturity in maturities {
        for &frequency in &frequencies {
            let months_per_period = 12 / frequency;
            for k in 1..=6 {
                let months_back = k * months_per_period;
                let settlement = format!("EDATE({maturity},-{months_back})");

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
                    assert_close(daysnc, days, 1e-12);

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
        || eval_number_or_skip(
            &mut sheet,
            "=COUPDAYS(DATE(2020,7,15),DATE(2021,7,15),2,0)",
        )
        .is_none()
    {
        return;
    }

    let maturities = ["DATE(2030,12,31)", "DATE(2031,2,28)", "DATE(2030,7,15)"];
    let frequencies = [1, 2, 4];
    let bases = [0, 4];
    let deltas = [1, 15, 30];

    for maturity in maturities {
        for &frequency in &frequencies {
            let months_per_period = 12 / frequency;

            // Ensure there's room to step back and still have a valid next coupon date.
            for k in 1..=6 {
                let months_back = k * months_per_period;
                let pcd = format!("EDATE({maturity},-{months_back})");

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
    if eval_number_or_skip(
        &mut sheet,
        "=COUPPCD(DATE(2020,7,1),DATE(2021,1,1),2,0)",
    )
    .is_none()
        || eval_number_or_skip(
            &mut sheet,
            "=COUPNCD(DATE(2020,7,1),DATE(2021,1,1),2,0)",
        )
        .is_none()
        || eval_number_or_skip(
            &mut sheet,
            "=COUPNUM(DATE(2020,7,1),DATE(2021,1,1),2,0)",
        )
        .is_none()
    {
        return;
    }

    let maturities = ["DATE(2030,12,31)", "DATE(2031,2,28)", "DATE(2030,7,15)"];
    let frequencies = [1, 2, 4];
    let bases = [0, 1, 2, 3, 4];

    for maturity in maturities {
        for &frequency in &frequencies {
            let months_per_period = 12 / frequency;
            for k in 1..=6 {
                let months_back = k * months_per_period;
                let settlement = format!("EDATE({maturity},-{months_back})");
                let expected_ncd = format!("EDATE({settlement},{months_per_period})");

                for &basis in &bases {
                    let pcd = eval_number(
                        &mut sheet,
                        &format!("=COUPPCD({settlement},{maturity},{frequency},{basis})"),
                    );
                    let settlement_serial =
                        eval_number(&mut sheet, &format!("={settlement}"));
                    assert_close(pcd, settlement_serial, 0.0);

                    let ncd = eval_number(
                        &mut sheet,
                        &format!("=COUPNCD({settlement},{maturity},{frequency},{basis})"),
                    );
                    let expected_ncd_serial =
                        eval_number(&mut sheet, &format!("={expected_ncd}"));
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

    let maturities = ["DATE(2030,12,31)", "DATE(2030,7,15)"];
    let frequencies = [1, 2, 4];
    let bases = [0, 1, 2, 3, 4];
    let rates = [0.03, 0.065];
    let yields = [0.01, 0.045, 0.11];
    let redemptions = [100.0, 105.0];

    for maturity in maturities {
        for &frequency in &frequencies {
            let months_per_period = 12 / frequency;
            for k in 1..=5 {
                let months_back = k * months_per_period;
                let settlement = format!("EDATE({maturity},-{months_back})");

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

    let maturities = ["DATE(2030,12,31)", "DATE(2031,2,28)"];
    let frequencies = [1, 2, 4];
    let bases = [0, 1, 2, 3, 4];
    let coupons = [0.025, 0.08];
    let yields = [0.02, 0.055, 0.12];

    for maturity in maturities {
        for &frequency in &frequencies {
            let months_per_period = 12 / frequency;
            for k in 1..=6 {
                let months_back = k * months_per_period;
                let settlement = format!("EDATE({maturity},-{months_back})");

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
