use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::error::ExcelError;
use formula_engine::functions::financial::{oddfprice, oddfyield, oddlprice, oddlyield};
use formula_engine::{ErrorKind, Value};

use super::harness::TestSheet;

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

#[test]
fn oddlprice_basis1_uses_prev_coupon_period_for_e() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2023, 2, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2023, 5, 15), system).unwrap();
    let last_interest = ymd_to_serial(ExcelDate::new(2023, 1, 31), system).unwrap();

    // Zero-coupon: price depends only on dsc/E exponent.
    let rate = 0.0;
    let yld = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 1;

    let price = oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();

    // For basis=1 with end-of-month coupon dates, E differs depending on whether you look
    // forward or backward from last_interest (Jan31->Jul31 is 181 days; Jul31->Jan31 is 184).
    //
    // The odd-coupon implementation mirrors Excel's `COUP*` day count conventions, which use the
    // *previous* regular coupon period length (`prev_coupon` -> `last_interest`) as `E`.
    let e = 184.0;
    let dsc = 89.0; // 2023-02-15 -> 2023-05-15
    let frac = dsc / e;
    let denom = 1.0 + yld / (frequency as f64);
    let expected = redemption / denom.powf(frac);
    assert_close(price, expected, 1e-12);

    // Also validate formula evaluation wiring.
    let mut sheet = TestSheet::new();
    let value =
        sheet.eval("=ODDLPRICE(DATE(2023,2,15),DATE(2023,5,15),DATE(2023,1,31),0,0.05,100,2,1)");
    match value {
        Value::Number(n) => assert_close(n, expected, 1e-12),
        other => panic!("expected number, got {other:?}"),
    }
}

#[test]
fn oddfprice_basis1_uses_prev_coupon_period_for_e() {
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = ymd_to_serial(ExcelDate::new(2022, 12, 20), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2024, 7, 31), system).unwrap();
    let issue = ymd_to_serial(ExcelDate::new(2022, 12, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2023, 1, 31), system).unwrap();

    let rate = 0.0;
    let yld = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 1;

    let price = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();

    // For basis=1 with end-of-month coupon dates:
    // prev_quasi = 2022-07-31, first_coupon = 2023-01-31 => E = 184 days.
    let e = 184.0;
    let dsc = 42.0; // 2022-12-20 -> 2023-01-31
    let frac = dsc / e;
    let n_coupons = 4.0; // 2023-01-31, 2023-07-31, 2024-01-31, 2024-07-31
    let exponent = frac + (n_coupons - 1.0);
    let denom = 1.0 + yld / (frequency as f64);
    let expected = redemption / denom.powf(exponent);
    assert_close(price, expected, 1e-12);

    let mut sheet = TestSheet::new();
    let value = sheet.eval("=ODDFPRICE(DATE(2022,12,20),DATE(2024,7,31),DATE(2022,12,15),DATE(2023,1,31),0,0.05,100,2,1)");
    match value {
        Value::Number(n) => assert_close(n, expected, 1e-12),
        other => panic!("expected number, got {other:?}"),
    }
}

#[test]
fn odd_coupon_coupon_payment_is_based_on_face_value() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Excel's odd-coupon bond functions return a price per $100 face value, and the periodic coupon
    // amount is computed from the $100 face value (not from the `redemption` amount).
    let rate = 0.10;
    let yld = 0.0;
    let redemption = 105.0;
    let frequency = 2;
    let basis = 0;

    let c = 100.0 * rate / (frequency as f64);

    // ODDF*: settlement == first_coupon is allowed.
    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let settlement = first_coupon;
    let maturity = ymd_to_serial(ExcelDate::new(2021, 7, 1), system).unwrap();

    // With `yld=0` and settlement on the first coupon date, the clean price is just the sum of the
    // remaining cashflows: two regular coupons plus redemption.
    let expected_oddf = redemption + 2.0 * c;
    let price_oddf = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert_close(price_oddf, expected_oddf, 1e-12);
    let recovered_yld = oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        price_oddf,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert_close(recovered_yld, yld, 1e-10);

    // ODDL*: settlement == last_interest is allowed.
    let last_interest = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let settlement2 = last_interest;
    let maturity2 = ymd_to_serial(ExcelDate::new(2021, 1, 1), system).unwrap();

    // With `yld=0` and a single odd-last period that is exactly one regular coupon period, the
    // clean price is redemption + one regular coupon.
    let expected_oddl = redemption + c;
    let price_oddl = oddlprice(
        settlement2,
        maturity2,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert_close(price_oddl, expected_oddl, 1e-12);
    let recovered_yld = oddlyield(
        settlement2,
        maturity2,
        last_interest,
        rate,
        price_oddl,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert_close(recovered_yld, yld, 1e-10);

    // Worksheet wiring.
    let mut sheet = TestSheet::new();
    let v = sheet.eval(
        "=ODDFPRICE(DATE(2020,7,1),DATE(2021,7,1),DATE(2020,1,1),DATE(2020,7,1),0.1,0,105,2,0)",
    );
    match v {
        Value::Number(n) => assert_close(n, expected_oddf, 1e-12),
        other => panic!("expected number, got {other:?}"),
    }
    let v = sheet.eval("=ODDLPRICE(DATE(2020,7,1),DATE(2021,1,1),DATE(2020,7,1),0.1,0,105,2,0)");
    match v {
        Value::Number(n) => assert_close(n, expected_oddl, 1e-12),
        other => panic!("expected number, got {other:?}"),
    }
}

#[test]
fn odd_coupon_settlement_boundary_behavior() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Pinned by current engine behavior; verify against real Excel via
    // tools/excel-oracle/run-excel-oracle.ps1 (Task 393).
    //
    // Settlement ordering rules:
    //
    // - ODDL*: settlement < maturity, maturity > last_interest (settlement may be <= last_interest).
    //   If settlement < last_interest, the engine PVs remaining regular coupons through last_interest
    //   plus the final odd stub cashflow at maturity.
    // - ODDF*: issue <= settlement <= first_coupon <= maturity, with the additional strict
    //   constraints `issue < first_coupon` and `settlement < maturity`.
    //
    // ODDF* rejects settlement > first_coupon, issue > settlement, issue == first_coupon, and
    // settlement == maturity (see `crates/formula-engine/tests/odd_coupon_date_boundaries.rs` for
    // pinned boundary behavior).

    // ODDL*: settlement == last_interest should be accepted.
    let maturity = ymd_to_serial(ExcelDate::new(2023, 5, 15), system).unwrap();
    let last_interest = ymd_to_serial(ExcelDate::new(2023, 1, 31), system).unwrap();
    let settlement_eq_last = last_interest;
    let yld_in = 0.06;
    let pr = oddlprice(
        settlement_eq_last,
        maturity,
        last_interest,
        0.05,
        yld_in,
        100.0,
        2,
        0,
        system,
    )
    .expect("ODDLPRICE should accept settlement == last_interest");
    let yld_out = oddlyield(
        settlement_eq_last,
        maturity,
        last_interest,
        0.05,
        pr,
        100.0,
        2,
        0,
        system,
    )
    .expect("ODDLYIELD should converge when settlement == last_interest");
    assert_close(yld_out, yld_in, 1e-6);

    // ODDL*: settlement < last_interest should also be accepted (multiple regular coupons remain).
    let settlement_before_last = ymd_to_serial(ExcelDate::new(2022, 11, 1), system).unwrap();
    let pr = oddlprice(
        settlement_before_last,
        maturity,
        last_interest,
        0.05,
        yld_in,
        100.0,
        2,
        0,
        system,
    )
    .expect("ODDLPRICE should accept settlement < last_interest");
    assert!(
        pr.is_finite() && pr > 0.0,
        "expected positive finite price, got {pr}"
    );
    let yld_out = oddlyield(
        settlement_before_last,
        maturity,
        last_interest,
        0.05,
        pr,
        100.0,
        2,
        0,
        system,
    )
    .expect("ODDLYIELD should converge when settlement < last_interest");
    assert_close(yld_out, yld_in, 1e-6);
    // ODDF*: settlement == first_coupon is allowed (settlement on the first coupon date; see
    // `crates/formula-engine/tests/odd_coupon_date_boundaries.rs`).
    let issue = ymd_to_serial(ExcelDate::new(2022, 12, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2023, 1, 31), system).unwrap();
    let maturity2 = ymd_to_serial(ExcelDate::new(2024, 7, 31), system).unwrap();
    let settlement_eq_first = first_coupon;
    let pr_oddf_eq_first = oddfprice(
        settlement_eq_first,
        maturity2,
        issue,
        first_coupon,
        0.05,
        yld_in,
        100.0,
        2,
        0,
        system,
    )
    .expect("ODDFPRICE should accept settlement == first_coupon");
    assert!(
        pr_oddf_eq_first.is_finite() && pr_oddf_eq_first > 0.0,
        "expected positive finite price, got {pr_oddf_eq_first}"
    );
    let yld_out = oddfyield(
        settlement_eq_first,
        maturity2,
        issue,
        first_coupon,
        0.05,
        pr_oddf_eq_first,
        100.0,
        2,
        0,
        system,
    )
    .expect("ODDFYIELD should converge when settlement == first_coupon");
    assert_close(yld_out, yld_in, 1e-6);

    // ODDF*: settlement > first_coupon => #NUM! (excel-oracle cases:
    // - `oddfprice_invalid_schedule_settlement_after_first_1bc49b18aff1`
    // - `oddfyield_invalid_schedule_settlement_after_first_938531ed93bc`)
    let settlement_after_first = ymd_to_serial(ExcelDate::new(2023, 2, 1), system).unwrap();
    let result = oddfprice(
        settlement_after_first,
        maturity2,
        issue,
        first_coupon,
        0.05,
        0.06,
        100.0,
        2,
        0,
        system,
    );
    assert_eq!(result, Err(ExcelError::Num));
    let result = oddfyield(
        settlement_after_first,
        maturity2,
        issue,
        first_coupon,
        0.05,
        99.0,
        100.0,
        2,
        0,
        system,
    );
    assert_eq!(result, Err(ExcelError::Num));

    let mut sheet = TestSheet::new();
    let v =
        sheet.eval("=ODDLPRICE(DATE(2023,1,31),DATE(2023,5,15),DATE(2023,1,31),0.05,0.06,100,2,0)");
    assert!(
        matches!(v, Value::Number(n) if n.is_finite()),
        "expected finite number for worksheet ODDLPRICE when settlement == last_interest, got {v:?}"
    );
    let v = sheet.eval("=ODDLYIELD(DATE(2023,1,31),DATE(2023,5,15),DATE(2023,1,31),0.05,ODDLPRICE(DATE(2023,1,31),DATE(2023,5,15),DATE(2023,1,31),0.05,0.06,100,2,0),100,2,0)");
    assert!(
        matches!(v, Value::Number(n) if (n - 0.06).abs() <= 1e-6),
        "expected yield ~0.06 for worksheet ODDLYIELD when settlement == last_interest, got {v:?}"
    );
    // Settlement < last_interest is allowed (regular coupons remain before `last_interest`); ensure
    // worksheet functions evaluate too.
    let v =
        sheet.eval("=ODDLPRICE(DATE(2022,11,1),DATE(2023,5,15),DATE(2023,1,31),0.05,0.06,100,2,0)");
    assert!(
        matches!(v, Value::Number(n) if n.is_finite()),
        "expected finite number for worksheet ODDLPRICE when settlement < last_interest, got {v:?}"
    );
    let v = sheet.eval("=ODDLYIELD(DATE(2022,11,1),DATE(2023,5,15),DATE(2023,1,31),0.05,ODDLPRICE(DATE(2022,11,1),DATE(2023,5,15),DATE(2023,1,31),0.05,0.06,100,2,0),100,2,0)");
    match v {
        Value::Number(n) => assert_close(n, 0.06, 1e-9),
        other => panic!("expected number for worksheet ODDLYIELD when settlement < last_interest, got {other:?}"),
    }
    let v = sheet.eval("=ODDFPRICE(DATE(2023,1,31),DATE(2024,7,31),DATE(2022,12,15),DATE(2023,1,31),0.05,0.06,100,2,0)");
    match v {
        Value::Number(n) => assert_close(n, pr_oddf_eq_first, 1e-9),
        other => panic!(
            "expected number for worksheet ODDFPRICE when settlement == first_coupon, got {other:?}"
        ),
    }
    let v = sheet.eval("=ODDFYIELD(DATE(2023,1,31),DATE(2024,7,31),DATE(2022,12,15),DATE(2023,1,31),0.05,ODDFPRICE(DATE(2023,1,31),DATE(2024,7,31),DATE(2022,12,15),DATE(2023,1,31),0.05,0.06,100,2,0),100,2,0)");
    assert!(
        matches!(v, Value::Number(n) if (n - 0.06).abs() <= 1e-6),
        "expected yield ~0.06 for worksheet ODDFYIELD when settlement == first_coupon, got {v:?}"
    );
    let v = sheet.eval("=ODDFPRICE(DATE(2023,2,1),DATE(2024,7,31),DATE(2022,12,15),DATE(2023,1,31),0.05,0.06,100,2,0)");
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDFPRICE when settlement > first_coupon, got {v:?}"
    );
    let v = sheet.eval("=ODDFYIELD(DATE(2023,2,1),DATE(2024,7,31),DATE(2022,12,15),DATE(2023,1,31),0.05,99,100,2,0)");
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDFYIELD when settlement > first_coupon, got {v:?}"
    );
}

#[test]
fn oddf_issue_equal_settlement_boundary_roundtrip() {
    let system = ExcelDateSystem::EXCEL_1900;

    // ODDF*: issue == settlement is allowed (zero accrued interest).
    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let settlement = issue;
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 7, 1), system).unwrap();

    let rate = 0.05;
    let yld_in = 0.06;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    let pr = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld_in,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDFPRICE should allow issue == settlement");
    assert!(
        pr.is_finite() && pr > 0.0,
        "expected positive finite price, got {pr}"
    );

    let yld_out = oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        pr,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDFYIELD should converge when issue == settlement");
    assert_close(yld_out, yld_in, 1e-6);

    // Also validate worksheet functions.
    let mut sheet = TestSheet::new();
    let v = sheet.eval(
        "=ODDFPRICE(DATE(2020,1,1),DATE(2021,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.06,100,2,0)",
    );
    assert!(
        matches!(v, Value::Number(n) if n.is_finite() && n > 0.0),
        "expected positive finite number for worksheet ODDFPRICE when issue == settlement, got {v:?}"
    );
    let v = sheet.eval("=ODDFYIELD(DATE(2020,1,1),DATE(2021,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,ODDFPRICE(DATE(2020,1,1),DATE(2021,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.06,100,2,0),100,2,0)");
    assert!(
        matches!(v, Value::Number(n) if (n - yld_in).abs() <= 1e-6),
        "expected yield ~0.06 for worksheet ODDFYIELD when issue == settlement, got {v:?}"
    );
}

#[test]
fn odd_coupon_yield_price_roundtrip() {
    let system = ExcelDateSystem::EXCEL_1900;

    // ODDL* roundtrip.
    let settlement = ymd_to_serial(ExcelDate::new(2023, 2, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2023, 5, 15), system).unwrap();
    let last_interest = ymd_to_serial(ExcelDate::new(2023, 1, 31), system).unwrap();
    let rate = 0.05;
    let yld_in = 0.06;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    let pr = oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        yld_in,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    let yld_out = oddlyield(
        settlement,
        maturity,
        last_interest,
        rate,
        pr,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert_close(yld_out, yld_in, 1e-6);

    // ODDF* roundtrip.
    let settlement = ymd_to_serial(ExcelDate::new(2022, 12, 20), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2024, 7, 31), system).unwrap();
    let issue = ymd_to_serial(ExcelDate::new(2022, 12, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2023, 1, 31), system).unwrap();
    let pr = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        yld_in,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    let yld_out = oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        pr,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert_close(yld_out, yld_in, 1e-6);
}

#[test]
fn odd_coupon_internal_functions_reject_non_finite_numeric_inputs() {
    let system = ExcelDateSystem::EXCEL_1900;
    // Use separate valid schedules for ODDF* and ODDL* so these assertions exercise the
    // non-finite numeric guards (not chronology errors).
    let oddf_settlement = ymd_to_serial(ExcelDate::new(2020, 3, 1), system).unwrap();
    let oddf_maturity = ymd_to_serial(ExcelDate::new(2023, 7, 1), system).unwrap();
    let oddf_issue = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let oddf_first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();

    let oddl_settlement = ymd_to_serial(ExcelDate::new(2022, 11, 1), system).unwrap();
    let oddl_maturity = ymd_to_serial(ExcelDate::new(2023, 3, 1), system).unwrap();
    let oddl_last_interest = ymd_to_serial(ExcelDate::new(2022, 7, 1), system).unwrap();

    // ODDFPRICE
    assert_eq!(
        oddfprice(
            oddf_settlement,
            oddf_maturity,
            oddf_issue,
            oddf_first_coupon,
            f64::INFINITY,
            0.05,
            100.0,
            1,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    assert_eq!(
        oddfprice(
            oddf_settlement,
            oddf_maturity,
            oddf_issue,
            oddf_first_coupon,
            0.06,
            f64::NAN,
            100.0,
            1,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    assert_eq!(
        oddfprice(
            oddf_settlement,
            oddf_maturity,
            oddf_issue,
            oddf_first_coupon,
            0.06,
            0.05,
            f64::NEG_INFINITY,
            1,
            0,
            system
        ),
        Err(ExcelError::Num)
    );

    // ODDFYIELD
    assert_eq!(
        oddfyield(
            oddf_settlement,
            oddf_maturity,
            oddf_issue,
            oddf_first_coupon,
            f64::NAN,
            95.0,
            100.0,
            1,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    assert_eq!(
        oddfyield(
            oddf_settlement,
            oddf_maturity,
            oddf_issue,
            oddf_first_coupon,
            0.06,
            f64::INFINITY,
            100.0,
            1,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    assert_eq!(
        oddfyield(
            oddf_settlement,
            oddf_maturity,
            oddf_issue,
            oddf_first_coupon,
            0.06,
            95.0,
            f64::NAN,
            1,
            0,
            system
        ),
        Err(ExcelError::Num)
    );

    // ODDLPRICE
    assert_eq!(
        oddlprice(
            oddl_settlement,
            oddl_maturity,
            oddl_last_interest,
            f64::NEG_INFINITY,
            0.05,
            100.0,
            1,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    assert_eq!(
        oddlprice(
            oddl_settlement,
            oddl_maturity,
            oddl_last_interest,
            0.06,
            f64::INFINITY,
            100.0,
            1,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    assert_eq!(
        oddlprice(
            oddl_settlement,
            oddl_maturity,
            oddl_last_interest,
            0.06,
            0.05,
            f64::NAN,
            1,
            0,
            system
        ),
        Err(ExcelError::Num)
    );

    // ODDLYIELD
    assert_eq!(
        oddlyield(
            oddl_settlement,
            oddl_maturity,
            oddl_last_interest,
            f64::INFINITY,
            95.0,
            100.0,
            1,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    assert_eq!(
        oddlyield(
            oddl_settlement,
            oddl_maturity,
            oddl_last_interest,
            0.06,
            f64::NAN,
            100.0,
            1,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    assert_eq!(
        oddlyield(
            oddl_settlement,
            oddl_maturity,
            oddl_last_interest,
            0.06,
            95.0,
            f64::NEG_INFINITY,
            1,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
}

#[test]
fn odd_coupon_invalid_schedule_inputs_return_num_errors() {
    // These cases are mirrored in the Excel-oracle corpus (tagged `odd_coupon` + `invalid_schedule`).
    //
    // NOTE: The pinned Excel dataset in CI is currently a synthetic baseline generated from the
    // engine. Once Task 486 lands (real Excel patching), these expected errors should be validated
    // against real Excel and updated if needed.
    let system = ExcelDateSystem::EXCEL_1900;

    // ---------------------------------------------------------------------
    // ODDF*: first_coupon not reachable from maturity-anchored schedule.
    // ---------------------------------------------------------------------
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 20), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 8, 30), system).unwrap();
    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 2, 28), system).unwrap();

    assert_eq!(
        oddfprice(
            settlement,
            maturity,
            issue,
            first_coupon,
            0.08,
            0.075,
            100.0,
            2,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    assert_eq!(
        oddfyield(
            settlement,
            maturity,
            issue,
            first_coupon,
            0.08,
            98.0,
            100.0,
            2,
            0,
            system
        ),
        Err(ExcelError::Num)
    );

    // Worksheet evaluation should surface the same #NUM!.
    let mut sheet = TestSheet::new();
    let v = sheet.eval("=ODDFPRICE(DATE(2020,1,20),DATE(2021,8,30),DATE(2020,1,15),DATE(2020,2,28),0.08,0.075,100,2,0)");
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDFPRICE invalid schedule, got {v:?}"
    );
    let v = sheet.eval("=ODDFYIELD(DATE(2020,1,20),DATE(2021,8,30),DATE(2020,1,15),DATE(2020,2,28),0.08,98,100,2,0)");
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDFYIELD invalid schedule, got {v:?}"
    );

    // ---------------------------------------------------------------------
    // ODDF*: EOM schedule mismatch (maturity EOM vs first_coupon not EOM, and vice versa).
    // ---------------------------------------------------------------------
    let settlement = ymd_to_serial(ExcelDate::new(2022, 12, 20), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2024, 7, 31), system).unwrap(); // EOM
    let issue = ymd_to_serial(ExcelDate::new(2022, 12, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2023, 1, 30), system).unwrap(); // not EOM

    assert_eq!(
        oddfprice(
            settlement,
            maturity,
            issue,
            first_coupon,
            0.05,
            0.06,
            100.0,
            2,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    assert_eq!(
        oddfyield(
            settlement,
            maturity,
            issue,
            first_coupon,
            0.05,
            98.0,
            100.0,
            2,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    let v = sheet.eval("=ODDFPRICE(DATE(2022,12,20),DATE(2024,7,31),DATE(2022,12,15),DATE(2023,1,30),0.05,0.06,100,2,0)");
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDFPRICE invalid schedule (maturity EOM, first_coupon not), got {v:?}"
    );
    let v = sheet.eval("=ODDFYIELD(DATE(2022,12,20),DATE(2024,7,31),DATE(2022,12,15),DATE(2023,1,30),0.05,98,100,2,0)");
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDFYIELD invalid schedule (maturity EOM, first_coupon not), got {v:?}"
    );

    let maturity = ymd_to_serial(ExcelDate::new(2024, 7, 30), system).unwrap(); // not EOM
    let first_coupon = ymd_to_serial(ExcelDate::new(2023, 1, 31), system).unwrap(); // EOM
    assert_eq!(
        oddfprice(
            settlement,
            maturity,
            issue,
            first_coupon,
            0.05,
            0.06,
            100.0,
            2,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    assert_eq!(
        oddfyield(
            settlement,
            maturity,
            issue,
            first_coupon,
            0.05,
            98.0,
            100.0,
            2,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    let v = sheet.eval("=ODDFPRICE(DATE(2022,12,20),DATE(2024,7,30),DATE(2022,12,15),DATE(2023,1,31),0.05,0.06,100,2,0)");
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDFPRICE invalid schedule (maturity not EOM, first_coupon EOM), got {v:?}"
    );
    let v = sheet.eval("=ODDFYIELD(DATE(2022,12,20),DATE(2024,7,30),DATE(2022,12,15),DATE(2023,1,31),0.05,98,100,2,0)");
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDFYIELD invalid schedule (maturity not EOM, first_coupon EOM), got {v:?}"
    );

    // ---------------------------------------------------------------------
    // ODDF*: minimal schedule ordering invalidity: settlement after first_coupon.
    // ---------------------------------------------------------------------
    let settlement = ymd_to_serial(ExcelDate::new(2020, 8, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2025, 1, 1), system).unwrap();
    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    assert_eq!(
        oddfprice(
            settlement,
            maturity,
            issue,
            first_coupon,
            0.05,
            0.04,
            100.0,
            2,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    assert_eq!(
        oddfyield(
            settlement,
            maturity,
            issue,
            first_coupon,
            0.05,
            98.0,
            100.0,
            2,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    let v = sheet.eval(
        "=ODDFPRICE(DATE(2020,8,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0)",
    );
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDFPRICE when settlement > first_coupon, got {v:?}"
    );
    let v = sheet.eval(
        "=ODDFYIELD(DATE(2020,8,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,98,100,2,0)",
    );
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDFYIELD when settlement > first_coupon, got {v:?}"
    );

    // ---------------------------------------------------------------------
    // ODDL*: last_interest must be strictly before maturity.
    // ---------------------------------------------------------------------
    let settlement = ymd_to_serial(ExcelDate::new(2024, 7, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2025, 1, 1), system).unwrap();
    let last_interest = ymd_to_serial(ExcelDate::new(2025, 1, 2), system).unwrap();

    assert_eq!(
        oddlprice(
            settlement,
            maturity,
            last_interest,
            0.05,
            0.04,
            100.0,
            2,
            0,
            system
        ),
        Err(ExcelError::Num)
    );
    assert_eq!(
        oddlyield(
            settlement,
            maturity,
            last_interest,
            0.05,
            99.0,
            100.0,
            2,
            0,
            system
        ),
        Err(ExcelError::Num)
    );

    let v =
        sheet.eval("=ODDLPRICE(DATE(2024,7,1),DATE(2025,1,1),DATE(2025,1,2),0.05,0.04,100,2,0)");
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDLPRICE invalid schedule, got {v:?}"
    );
    let v = sheet.eval("=ODDLYIELD(DATE(2024,7,1),DATE(2025,1,1),DATE(2025,1,2),0.05,99,100,2,0)");
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDLYIELD invalid schedule, got {v:?}"
    );
}

#[test]
fn odd_coupon_misaligned_last_interest_schedule_is_currently_accepted() {
    // These cases are mirrored in the Excel-oracle corpus (tagged `odd_coupon` + `invalid_schedule`).
    //
    // NOTE: The pinned Excel dataset in CI is currently a synthetic baseline generated from the
    // engine. Once Task 486 lands (real Excel patching), these numeric expectations should be
    // validated against real Excel and updated if needed.
    let system = ExcelDateSystem::EXCEL_1900;
    let mut sheet = TestSheet::new();

    // ---------------------------------------------------------------------
    // ODDL*: maturity is EOM but last_interest is not (basis=0).
    // ---------------------------------------------------------------------
    let settlement = ymd_to_serial(ExcelDate::new(2024, 8, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2025, 1, 31), system).unwrap();
    let last_interest = ymd_to_serial(ExcelDate::new(2024, 7, 30), system).unwrap();

    let pr = oddlprice(
        settlement,
        maturity,
        last_interest,
        0.05,
        0.04,
        100.0,
        2,
        0,
        system,
    )
    .expect("ODDLPRICE should currently accept misaligned last_interest schedule");
    assert_close(pr, 100.476_307_189_542_48, 1e-9);
    let y = oddlyield(
        settlement,
        maturity,
        last_interest,
        0.05,
        99.0,
        100.0,
        2,
        0,
        system,
    )
    .expect("ODDLYIELD should currently accept misaligned last_interest schedule");
    assert_close(y, 0.070_416_608_219_943_12, 1e-9);

    let v =
        sheet.eval("=ODDLPRICE(DATE(2024,8,1),DATE(2025,1,31),DATE(2024,7,30),0.05,0.04,100,2,0)");
    match v {
        Value::Number(n) => assert_close(n, pr, 1e-9),
        other => {
            panic!("expected number for worksheet ODDLPRICE misaligned schedule, got {other:?}")
        }
    }
    let v =
        sheet.eval("=ODDLYIELD(DATE(2024,8,1),DATE(2025,1,31),DATE(2024,7,30),0.05,99,100,2,0)");
    match v {
        Value::Number(n) => assert_close(n, y, 1e-9),
        other => {
            panic!("expected number for worksheet ODDLYIELD misaligned schedule, got {other:?}")
        }
    }

    // ---------------------------------------------------------------------
    // ODDL*: EOM mismatch variants under basis=1.
    // ---------------------------------------------------------------------
    let settlement = ymd_to_serial(ExcelDate::new(2024, 8, 15), system).unwrap();

    // maturity is EOM but last_interest is not.
    let maturity = ymd_to_serial(ExcelDate::new(2025, 1, 31), system).unwrap();
    let last_interest = ymd_to_serial(ExcelDate::new(2024, 7, 30), system).unwrap();
    let pr = oddlprice(
        settlement,
        maturity,
        last_interest,
        0.05,
        0.04,
        100.0,
        2,
        1,
        system,
    )
    .expect(
        "ODDLPRICE should currently accept basis=1 EOM mismatch (maturity EOM, last_interest not)",
    );
    assert_close(pr, 100.453_115_102_272_87, 1e-9);
    let y = oddlyield(
        settlement,
        maturity,
        last_interest,
        0.05,
        99.0,
        100.0,
        2,
        1,
        system,
    )
    .expect(
        "ODDLYIELD should currently accept basis=1 EOM mismatch (maturity EOM, last_interest not)",
    );
    assert_close(y, 0.072_192_898_126_549_61, 1e-9);

    // maturity is not EOM but last_interest is EOM.
    let maturity = ymd_to_serial(ExcelDate::new(2025, 1, 30), system).unwrap();
    let last_interest = ymd_to_serial(ExcelDate::new(2024, 7, 31), system).unwrap();
    let pr = oddlprice(
        settlement,
        maturity,
        last_interest,
        0.05,
        0.04,
        100.0,
        2,
        1,
        system,
    )
    .expect("ODDLPRICE should currently accept basis=1 EOM mismatch (maturity not EOM, last_interest EOM)");
    assert_close(pr, 100.450_830_831_707_98, 1e-9);
    let y = oddlyield(
        settlement,
        maturity,
        last_interest,
        0.05,
        99.0,
        100.0,
        2,
        1,
        system,
    )
    .expect("ODDLYIELD should currently accept basis=1 EOM mismatch (maturity not EOM, last_interest EOM)");
    assert_close(y, 0.072_339_574_523_340_2, 1e-9);
}
