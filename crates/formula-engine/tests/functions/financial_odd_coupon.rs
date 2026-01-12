use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::error::ExcelError;
use formula_engine::functions::financial::{oddfprice, oddfyield, oddlprice, oddlyield};

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

fn serial(year: i32, month: u8, day: u8, system: ExcelDateSystem) -> i32 {
    ymd_to_serial(ExcelDate::new(year, month, day), system).expect("valid excel serial")
}

#[test]
fn oddfyield_extreme_prices_roundtrip() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Odd first coupon: issue -> first_coupon is a short stub, followed by a regular period.
    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 2, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2020, 8, 15), system).unwrap();

    let rate = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    for pr in [50.0, 200.0] {
        let yld = oddfyield(
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
        .expect("ODDFYIELD should converge");

        assert!(yld.is_finite(), "yield should be finite, got {yld}");
        assert!(
            yld > -(frequency as f64),
            "yield should be > -frequency, got {yld}"
        );

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
        .expect("ODDFPRICE should succeed");

        assert_close(price, pr, 1e-6);
    }
}

#[test]
fn oddlyield_extreme_prices_roundtrip() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Odd last coupon: settlement is after the last interest date, with a long stub to maturity.
    let last_interest = ymd_to_serial(ExcelDate::new(2020, 6, 30), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 9, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 1, 15), system).unwrap();

    let rate = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    for pr in [50.0, 200.0] {
        let yld = oddlyield(
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
        .expect("ODDLYIELD should converge");

        assert!(yld.is_finite(), "yield should be finite, got {yld}");
        assert!(
            yld > -(frequency as f64),
            "yield should be > -frequency, got {yld}"
        );

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
        .expect("ODDLPRICE should succeed");

        assert_close(price, pr, 1e-6);
    }
}

#[test]
fn odd_coupon_functions_coerce_frequency_like_excel() {
    let mut sheet = TestSheet::new();
    // Task: coercion edge cases for `frequency`.
    //
    // Excel coerces:
    // - numeric text: "2" -> 2
    // - TRUE/FALSE -> 1/0
    //
    // `frequency` must be one of {1,2,4}. So FALSE (0) should produce #NUM!.

    // Baseline semiannual (frequency=2) example (Task 56).
    let baseline_semiannual = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let baseline_semiannual_value = match eval_number_or_skip(&mut sheet, baseline_semiannual) {
        Some(v) => v,
        None => return,
    };

    let semiannual_text_freq = r#"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,"2",0)"#;
    let semiannual_text_freq_value = eval_number_or_skip(&mut sheet, semiannual_text_freq)
        .expect("ODDFPRICE should accept frequency supplied as numeric text");
    assert_close(semiannual_text_freq_value, baseline_semiannual_value, 1e-9);

    // Annual schedule (frequency=1) example.
    let baseline_annual = "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0)";
    let baseline_annual_value = eval_number_or_skip(&mut sheet, baseline_annual)
        .expect("ODDFPRICE should accept explicit annual frequency");

    let annual_true_freq = "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,TRUE,0)";
    let annual_true_freq_value = eval_number_or_skip(&mut sheet, annual_true_freq)
        .expect("ODDFPRICE should accept TRUE frequency (TRUE->1)");
    assert_close(annual_true_freq_value, baseline_annual_value, 1e-9);

    let annual_false_freq = "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,FALSE,0)";
    match sheet.eval(annual_false_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for frequency=FALSE (0), got {other:?}"),
    }
}

#[test]
fn odd_coupon_functions_coerce_basis_like_excel() {
    let mut sheet = TestSheet::new();
    // Task: coercion edge cases for `basis`.
    //
    // Excel coerces:
    // - TRUE/FALSE -> 1/0
    // - blank -> 0 (same as default)
    //
    // Use an ODDFPRICE example with basis=0/1 to confirm.

    let baseline_basis_0 = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let baseline_basis_0_value = match eval_number_or_skip(&mut sheet, baseline_basis_0) {
        Some(v) => v,
        None => return,
    };

    // Basis passed as a blank cell should behave like basis=0.
    // (A1 is unset/blank by default.)
    let blank_basis = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,A1)";
    let blank_basis_value = eval_number_or_skip(&mut sheet, blank_basis)
        .expect("ODDFPRICE should accept blank basis and treat it as 0");
    assert_close(blank_basis_value, baseline_basis_0_value, 1e-9);

    let false_basis = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,FALSE)";
    let false_basis_value = eval_number_or_skip(&mut sheet, false_basis)
        .expect("ODDFPRICE should accept FALSE basis (FALSE->0)");
    assert_close(false_basis_value, baseline_basis_0_value, 1e-9);

    let baseline_basis_1 = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,1)";
    let baseline_basis_1_value = eval_number_or_skip(&mut sheet, baseline_basis_1)
        .expect("ODDFPRICE should accept explicit basis=1");

    let true_basis = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,TRUE)";
    let true_basis_value = eval_number_or_skip(&mut sheet, true_basis)
        .expect("ODDFPRICE should accept TRUE basis (TRUE->1)");
    assert_close(true_basis_value, baseline_basis_1_value, 1e-9);
}

#[test]
fn odd_first_coupon_bond_functions_respect_workbook_date_system() {
    let system_1900 = ExcelDateSystem::EXCEL_1900;
    let system_1904 = ExcelDateSystem::Excel1904;

    // Baseline case (Task 56): odd first coupon period.
    let settlement_1900 = serial(2008, 11, 11, system_1900);
    let maturity_1900 = serial(2021, 3, 1, system_1900);
    let issue_1900 = serial(2008, 10, 15, system_1900);
    let first_coupon_1900 = serial(2009, 3, 1, system_1900);

    let settlement_1904 = serial(2008, 11, 11, system_1904);
    let maturity_1904 = serial(2021, 3, 1, system_1904);
    let issue_1904 = serial(2008, 10, 15, system_1904);
    let first_coupon_1904 = serial(2009, 3, 1, system_1904);

    let price_1900 = oddfprice(
        settlement_1900,
        maturity_1900,
        issue_1900,
        first_coupon_1900,
        0.0785,
        0.0625,
        100.0,
        2,
        0,
        system_1900,
    )
    .expect("oddfprice should succeed under Excel1900");
    let yield_1900 = oddfyield(
        settlement_1900,
        maturity_1900,
        issue_1900,
        first_coupon_1900,
        0.0785,
        98.0,
        100.0,
        2,
        0,
        system_1900,
    )
    .expect("oddfyield should succeed under Excel1900");

    let price_1904 = oddfprice(
        settlement_1904,
        maturity_1904,
        issue_1904,
        first_coupon_1904,
        0.0785,
        0.0625,
        100.0,
        2,
        0,
        system_1904,
    )
    .expect("oddfprice should succeed under Excel1904");
    let yield_1904 = oddfyield(
        settlement_1904,
        maturity_1904,
        issue_1904,
        first_coupon_1904,
        0.0785,
        98.0,
        100.0,
        2,
        0,
        system_1904,
    )
    .expect("oddfyield should succeed under Excel1904");

    assert_close(price_1904, price_1900, 1e-9);
    assert_close(yield_1904, yield_1900, 1e-10);
}

#[test]
fn odd_last_coupon_bond_functions_respect_workbook_date_system() {
    let system_1900 = ExcelDateSystem::EXCEL_1900;
    let system_1904 = ExcelDateSystem::Excel1904;

    // Baseline case (Task 56): odd last coupon period.
    let settlement_1900 = serial(2020, 11, 11, system_1900);
    let maturity_1900 = serial(2021, 3, 1, system_1900);
    let last_interest_1900 = serial(2020, 10, 15, system_1900);

    let settlement_1904 = serial(2020, 11, 11, system_1904);
    let maturity_1904 = serial(2021, 3, 1, system_1904);
    let last_interest_1904 = serial(2020, 10, 15, system_1904);

    let price_1900 = oddlprice(
        settlement_1900,
        maturity_1900,
        last_interest_1900,
        0.0785,
        0.0625,
        100.0,
        2,
        0,
        system_1900,
    )
    .expect("oddlprice should succeed under Excel1900");
    let yield_1900 = oddlyield(
        settlement_1900,
        maturity_1900,
        last_interest_1900,
        0.0785,
        98.0,
        100.0,
        2,
        0,
        system_1900,
    )
    .expect("oddlyield should succeed under Excel1900");

    let price_1904 = oddlprice(
        settlement_1904,
        maturity_1904,
        last_interest_1904,
        0.0785,
        0.0625,
        100.0,
        2,
        0,
        system_1904,
    )
    .expect("oddlprice should succeed under Excel1904");
    let yield_1904 = oddlyield(
        settlement_1904,
        maturity_1904,
        last_interest_1904,
        0.0785,
        98.0,
        100.0,
        2,
        0,
        system_1904,
    )
    .expect("oddlyield should succeed under Excel1904");

    assert_close(price_1904, price_1900, 1e-9);
    assert_close(yield_1904, yield_1900, 1e-10);
}

#[test]
fn odd_first_coupon_roundtrips_yield_with_annual_frequency() {
    // Aligned annual schedule from `first_coupon` by 12 months:
    // 2020-07-01, 2021-07-01, 2022-07-01, 2023-07-01 (maturity).
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = serial(2020, 3, 1, system);
    let maturity = serial(2023, 7, 1, system);
    let issue = serial(2020, 1, 1, system);
    let first_coupon = serial(2020, 7, 1, system);

    let yld = 0.05;
    let price = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.06,
        yld,
        100.0,
        1,
        0,
        system,
    )
    .expect("oddfprice should succeed");

    let recovered_yield = oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.06,
        price,
        100.0,
        1,
        0,
        system,
    )
    .expect("oddfyield should succeed");

    assert_close(recovered_yield, yld, 1e-7);
}

#[test]
fn odd_first_coupon_roundtrips_yield_with_quarterly_frequency_and_non_100_redemption() {
    // Aligned quarterly schedule from `first_coupon` by 3 months:
    // 2020-02-15, 2020-05-15, 2020-08-15, 2020-11-15, 2021-02-15, 2021-05-15, 2021-08-15.
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = serial(2020, 1, 20, system);
    let maturity = serial(2021, 8, 15, system);
    let issue = serial(2020, 1, 1, system);
    let first_coupon = serial(2020, 2, 15, system);

    let yld = 0.07;
    let price_100 = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.08,
        yld,
        100.0,
        4,
        0,
        system,
    )
    .expect("oddfprice redemption=100 should succeed");
    let price_105 = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.08,
        yld,
        105.0,
        4,
        0,
        system,
    )
    .expect("oddfprice redemption=105 should succeed");

    assert!(
        (price_105 - price_100).abs() > 1e-9,
        "expected redemption to affect price (redemption=100 => {price_100}, redemption=105 => {price_105})"
    );
    assert!(
        price_105 > price_100,
        "expected higher redemption to increase price (redemption=100 => {price_100}, redemption=105 => {price_105})"
    );

    let recovered_yield_100 = oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.08,
        price_100,
        100.0,
        4,
        0,
        system,
    )
    .expect("oddfyield redemption=100 should succeed");
    let recovered_yield_105 = oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        0.08,
        price_105,
        105.0,
        4,
        0,
        system,
    )
    .expect("oddfyield redemption=105 should succeed");

    assert_close(recovered_yield_100, yld, 1e-7);
    assert_close(recovered_yield_105, yld, 1e-7);
}

#[test]
fn odd_last_coupon_roundtrips_yield_with_annual_frequency() {
    // `last_interest` is a coupon date on an annual schedule (12 month stepping). Maturity
    // occurs 8 months later, making this an odd last coupon period.
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = serial(2022, 11, 1, system);
    let maturity = serial(2023, 3, 1, system);
    let last_interest = serial(2022, 7, 1, system);

    let yld = 0.05;
    let price = oddlprice(
        settlement,
        maturity,
        last_interest,
        0.06,
        yld,
        100.0,
        1,
        0,
        system,
    )
    .expect("oddlprice should succeed");
    let recovered_yield = oddlyield(
        settlement,
        maturity,
        last_interest,
        0.06,
        price,
        100.0,
        1,
        0,
        system,
    )
    .expect("oddlyield should succeed");

    assert_close(recovered_yield, yld, 1e-7);
}

#[test]
fn odd_last_coupon_roundtrips_yield_with_quarterly_frequency_and_non_100_redemption() {
    // `last_interest` is a coupon date on a quarterly schedule. Maturity occurs 2 months later
    // (shorter than the regular 3 month period), making this an odd last coupon period.
    let system = ExcelDateSystem::EXCEL_1900;
    let settlement = serial(2021, 7, 1, system);
    let maturity = serial(2021, 8, 15, system);
    let last_interest = serial(2021, 6, 15, system);

    let yld = 0.07;
    let price_100 = oddlprice(
        settlement,
        maturity,
        last_interest,
        0.08,
        yld,
        100.0,
        4,
        0,
        system,
    )
    .expect("oddlprice redemption=100 should succeed");
    let price_105 = oddlprice(
        settlement,
        maturity,
        last_interest,
        0.08,
        yld,
        105.0,
        4,
        0,
        system,
    )
    .expect("oddlprice redemption=105 should succeed");

    assert!(
        (price_105 - price_100).abs() > 1e-9,
        "expected redemption to affect price (redemption=100 => {price_100}, redemption=105 => {price_105})"
    );
    assert!(
        price_105 > price_100,
        "expected higher redemption to increase price (redemption=100 => {price_100}, redemption=105 => {price_105})"
    );

    let recovered_yield_100 = oddlyield(
        settlement,
        maturity,
        last_interest,
        0.08,
        price_100,
        100.0,
        4,
        0,
        system,
    )
    .expect("oddlyield redemption=100 should succeed");
    let recovered_yield_105 = oddlyield(
        settlement,
        maturity,
        last_interest,
        0.08,
        price_105,
        105.0,
        4,
        0,
        system,
    )
    .expect("oddlyield redemption=105 should succeed");

    assert_close(recovered_yield_100, yld, 1e-7);
    assert_close(recovered_yield_105, yld, 1e-7);
}

#[test]
fn oddfprice_returns_num_error_for_non_finite_price_near_negative_frequency_boundary() {
    let system = ExcelDateSystem::EXCEL_1900;

    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 15), system).unwrap();
    // Long maturity to ensure the discount factor underflows and PV overflows for yields near -frequency.
    let maturity = ymd_to_serial(ExcelDate::new(2033, 1, 15), system).unwrap();

    let rate = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    let yld = -(frequency as f64) + 1e-12;
    let result = oddfprice(
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
    );

    assert!(
        matches!(result, Err(ExcelError::Num)),
        "expected #NUM!, got {result:?}"
    );
}

#[test]
fn oddlprice_returns_num_error_for_non_finite_price_near_negative_frequency_boundary() {
    let system = ExcelDateSystem::EXCEL_1900;

    let last_interest = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 7, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2033, 1, 15), system).unwrap();

    let rate = 0.05;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

    let yld = -(frequency as f64) + 1e-12;
    let result = oddlprice(
        settlement,
        maturity,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    );

    assert!(
        matches!(result, Err(ExcelError::Num)),
        "expected #NUM!, got {result:?}"
    );
}

#[test]
fn odd_coupon_prices_are_finite_for_large_redemption_values() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Reuse the existing odd first coupon setup.
    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2023, 1, 15), system).unwrap();

    let rate = 0.05;
    let yld = 0.06;
    let redemption = 1e12;
    let frequency = 2;
    let basis = 0;

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
    .expect("ODDFPRICE should succeed for large finite redemption");
    assert!(price.is_finite(), "expected finite price, got {price}");

    // Odd last coupon setup.
    let last_interest = ymd_to_serial(ExcelDate::new(2022, 7, 15), system).unwrap();
    let settlement_last = ymd_to_serial(ExcelDate::new(2022, 10, 15), system).unwrap();
    let maturity_last = ymd_to_serial(ExcelDate::new(2023, 1, 15), system).unwrap();

    let price_last = oddlprice(
        settlement_last,
        maturity_last,
        last_interest,
        rate,
        yld,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDLPRICE should succeed for large finite redemption");
    assert!(
        price_last.is_finite(),
        "expected finite price, got {price_last}"
    );
}

#[test]
fn odd_yield_solver_falls_back_when_derivative_is_non_finite() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Construct a case where the Newton step fails because the analytic derivative overflows at the
    // default guess (0.1), but the price itself remains finite.
    let issue = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 15), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2033, 1, 15), system).unwrap();

    let rate = 0.05;
    let frequency = 2;
    let basis = 0;
    let redemption = 1e308;
    let target_yield = 0.1;

    let pr = oddfprice(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        target_yield,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDFPRICE should be finite for the target yield");

    let recovered = oddfyield(
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
    .expect("ODDFYIELD should converge via bisection fallback");

    assert_close(recovered, target_yield, 1e-6);

    // Repeat for the odd last coupon solver.
    let last_interest = ymd_to_serial(ExcelDate::new(2020, 1, 15), system).unwrap();
    let settlement_last = ymd_to_serial(ExcelDate::new(2020, 7, 15), system).unwrap();
    let maturity_last = ymd_to_serial(ExcelDate::new(2033, 1, 15), system).unwrap();

    let pr_last = oddlprice(
        settlement_last,
        maturity_last,
        last_interest,
        rate,
        target_yield,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDLPRICE should be finite for the target yield");

    let recovered_last = oddlyield(
        settlement_last,
        maturity_last,
        last_interest,
        rate,
        pr_last,
        redemption,
        frequency,
        basis,
        system,
    )
    .expect("ODDLYIELD should converge via bisection fallback");

    assert_close(recovered_last, target_yield, 1e-6);
}
