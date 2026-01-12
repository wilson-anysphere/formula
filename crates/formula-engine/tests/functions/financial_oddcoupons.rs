use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::functions::financial::{oddfprice, oddfyield, oddlprice, oddlyield};
use formula_engine::{ErrorKind, ExcelError, Value};

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
fn odd_coupon_settlement_equal_coupon_dates_are_allowed() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Pinned by current engine behavior; verify against real Excel via
    // tools/excel-oracle/run-excel-oracle.ps1 (Task 393).
    //
    // Settlement must fall strictly inside the odd coupon period:
    // - ODDL*: last_interest < settlement < maturity
    // - ODDF*: issue < settlement < first_coupon <= maturity

    // ODDL*: settlement == last_interest => #NUM!
    let maturity = ymd_to_serial(ExcelDate::new(2023, 5, 15), system).unwrap();
    let last_interest = ymd_to_serial(ExcelDate::new(2023, 1, 31), system).unwrap();
    let settlement_eq_last = last_interest;
    let result = oddlprice(
        settlement_eq_last,
        maturity,
        last_interest,
        0.05,
        0.06,
        100.0,
        2,
        0,
        system,
    );
    assert!(
        matches!(result, Ok(n) if n.is_finite()),
        "expected finite number for ODDLPRICE when last_interest == settlement, got {result:?}"
    );

    // ODDL*: settlement < last_interest => #NUM! (excel-oracle case id prefix: fin_oddlprice_settle_before_last_b0_*)
    let settlement_before_last = ymd_to_serial(ExcelDate::new(2022, 11, 1), system).unwrap();
    let result = oddlprice(
        settlement_before_last,
        maturity,
        last_interest,
        0.05,
        0.06,
        100.0,
        2,
        0,
        system,
    );
    assert_eq!(result, Err(ExcelError::Num));
    let result = oddlyield(
        settlement_before_last,
        maturity,
        last_interest,
        0.05,
        99.0,
        100.0,
        2,
        0,
        system,
    );
    assert_eq!(result, Err(ExcelError::Num));

    // ODDF*: settlement == first_coupon => #NUM!
    let issue = ymd_to_serial(ExcelDate::new(2022, 12, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2023, 1, 31), system).unwrap();
    let maturity2 = ymd_to_serial(ExcelDate::new(2024, 7, 31), system).unwrap();
    let result = oddfprice(
        first_coupon, // settlement == first_coupon (outside odd first period)
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
    assert!(
        matches!(result, Ok(n) if n.is_finite()),
        "expected finite number for ODDFPRICE when settlement == first_coupon, got {result:?}"
    );

    // ODDF*: settlement > first_coupon => #NUM! (excel-oracle case id prefix: fin_oddfprice_settle_after_first_b0_*)
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
        "expected finite number for worksheet ODDLPRICE when last_interest == settlement, got {v:?}"
    );
    let v =
        sheet.eval("=ODDLPRICE(DATE(2022,11,1),DATE(2023,5,15),DATE(2023,1,31),0.05,0.06,100,2,0)");
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDLPRICE when settlement < last_interest, got {v:?}"
    );
    let v = sheet.eval("=ODDLYIELD(DATE(2022,11,1),DATE(2023,5,15),DATE(2023,1,31),0.05,99,100,2,0)");
    assert!(
        matches!(v, Value::Error(ErrorKind::Num)),
        "expected #NUM! for worksheet ODDLYIELD when settlement < last_interest, got {v:?}"
    );
    let v = sheet.eval("=ODDFPRICE(DATE(2023,1,31),DATE(2024,7,31),DATE(2022,12,15),DATE(2023,1,31),0.05,0.06,100,2,0)");
    assert!(
        matches!(v, Value::Number(n) if n.is_finite()),
        "expected finite number for worksheet ODDFPRICE when settlement == first_coupon, got {v:?}"
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
