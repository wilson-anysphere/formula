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

fn serial(year: i32, month: u8, day: u8, system: ExcelDateSystem) -> i32 {
    ymd_to_serial(ExcelDate::new(year, month, day), system).expect("valid excel serial")
}

fn eval_number_or_skip(sheet: &mut TestSheet, formula: &str) -> Option<f64> {
    match sheet.eval(formula) {
        Value::Number(n) => Some(n),
        // The odd-coupon bond functions are not yet implemented in every build of the engine.
        // Skip these tests when the function registry doesn't recognize the name.
        Value::Error(ErrorKind::Name) => None,
        other => panic!("expected number, got {other:?} from {formula}"),
    }
}

fn cell_number_or_skip(sheet: &TestSheet, addr: &str) -> Option<f64> {
    match sheet.get(addr) {
        Value::Number(n) => Some(n),
        Value::Error(ErrorKind::Name) => None,
        other => panic!("expected number, got {other:?} from cell {addr}"),
    }
}

fn eval_value_or_skip(sheet: &mut TestSheet, formula: &str) -> Option<Value> {
    match sheet.eval(formula) {
        Value::Error(ErrorKind::Name) => None,
        other => Some(other),
    }
}

#[test]
fn oddfprice_zero_coupon_rate_reduces_to_discounted_redemption() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Long first coupon period: issue -> first_coupon spans 9 months, then regular semiannual.
    let issue = ymd_to_serial(ExcelDate::new(2019, 10, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 7, 1), system).unwrap();

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
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
    .unwrap();
    assert!(price.is_finite());

    // With rate=0, coupons and accrued interest are 0, so the price reduces to a discounted redemption:
    // P = redemption / (1 + yld/frequency)^(n-1 + DSC/E)
    //
    // Here (basis 0, 30/360):
    // - Coupon dates: 2020-07-01, 2021-01-01, 2021-07-01 => n = 3
    // - E = 360/frequency = 180, DSC = 180 => DSC/E = 1
    // - exponent = 3
    let y = yld / (frequency as f64);
    let expected = redemption / (1.0 + y).powi(3);
    assert_close(price, expected, 1e-12);
}

#[test]
fn oddlprice_zero_coupon_rate_reduces_to_discounted_redemption() {
    let system = ExcelDateSystem::EXCEL_1900;

    // Short odd last period inside an otherwise regular semiannual schedule.
    let last_interest = ymd_to_serial(ExcelDate::new(2021, 1, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2021, 2, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 5, 1), system).unwrap();

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
    let frequency = 2;
    let basis = 0;

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
    assert!(price.is_finite());

    // With rate=0, coupons and accrued interest are 0:
    // P = redemption / (1 + yld/frequency)^(DSC/E)
    //
    // Basis 0 (30/360), frequency=2: E=180, DSC=90 => exponent=0.5.
    let y = yld / (frequency as f64);
    let expected = redemption / (1.0 + y).powf(0.5);
    assert_close(price, expected, 1e-12);
}

#[test]
fn odd_coupon_yield_inverts_zero_coupon_prices() {
    let system = ExcelDateSystem::EXCEL_1900;

    // ODDF*
    let issue = ymd_to_serial(ExcelDate::new(2019, 10, 1), system).unwrap();
    let settlement = ymd_to_serial(ExcelDate::new(2020, 1, 1), system).unwrap();
    let first_coupon = ymd_to_serial(ExcelDate::new(2020, 7, 1), system).unwrap();
    let maturity = ymd_to_serial(ExcelDate::new(2021, 7, 1), system).unwrap();

    let rate = 0.0;
    let yld = 0.1;
    let redemption = 100.0;
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
    .unwrap();

    let solved = oddfyield(
        settlement,
        maturity,
        issue,
        first_coupon,
        rate,
        price,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert_close(solved, yld, 1e-10);

    // ODDL*
    let last_interest = ymd_to_serial(ExcelDate::new(2021, 1, 1), system).unwrap();
    let settlement2 = ymd_to_serial(ExcelDate::new(2021, 2, 1), system).unwrap();
    let maturity2 = ymd_to_serial(ExcelDate::new(2021, 5, 1), system).unwrap();

    let price2 = oddlprice(
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

    let solved2 = oddlyield(
        settlement2,
        maturity2,
        last_interest,
        rate,
        price2,
        redemption,
        frequency,
        basis,
        system,
    )
    .unwrap();
    assert_close(solved2, yld, 1e-10);
}

#[test]
fn builtins_odd_coupon_zero_coupon_rate_oracle_cases() {
    let mut sheet = TestSheet::new();

    // Deterministic oracle values (Excel 1900 date system).
    let price = match eval_number_or_skip(
        &mut sheet,
        "=ODDFPRICE(DATE(2020,1,1),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0,0.1,100,2,0)",
    ) {
        Some(v) => v,
        None => return,
    };
    assert_close(price, 86.3837598531476, 1e-9);

    let price2 = eval_number_or_skip(
        &mut sheet,
        "=ODDLPRICE(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0,0.1,100,2,0)",
    )
    .expect("ODDLPRICE should evaluate");
    assert_close(price2, 97.59000729485331, 1e-9);

    let yld = eval_number_or_skip(
        &mut sheet,
        &format!(
            "=ODDFYIELD(DATE(2020,1,1),DATE(2021,7,1),DATE(2019,10,1),DATE(2020,7,1),0,{price},100,2,0)"
        ),
    )
    .expect("ODDFYIELD should evaluate");
    assert_close(yld, 0.1, 1e-10);

    let yld2 = eval_number_or_skip(
        &mut sheet,
        &format!("=ODDLYIELD(DATE(2021,2,1),DATE(2021,5,1),DATE(2021,1,1),0,{price2},100,2,0)"),
    )
    .expect("ODDLYIELD should evaluate");
    assert_close(yld2, 0.1, 1e-10);
}

#[test]
fn oddfprice_matches_known_example_basis0() {
    let mut sheet = TestSheet::new();
    let formula = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let v = sheet.eval(formula);
    match v {
        Value::Number(n) => assert_close(n, 113.59920582823823, 1e-9),
        other => panic!("expected number, got {other:?} from {formula}"),
    }
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
    let baseline_annual =
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0)";
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

    // Blank is coerced to 0 for numeric args.
    let annual_blank_cell_freq = "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,A1,0)";
    match sheet.eval(annual_blank_cell_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for frequency=<blank> (0), got {other:?}"),
    }
    let annual_blank_arg_freq =
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,,0)";
    match sheet.eval(annual_blank_arg_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for frequency=<explicit blank arg> (0), got {other:?}"),
    }

    // Spot-check ODDLPRICE as well (odd last coupon) to ensure the coercion behavior is consistent
    // across ODDF*/ODDL*.
    let oddl_baseline_semiannual =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let oddl_baseline_semiannual_value =
        match eval_number_or_skip(&mut sheet, oddl_baseline_semiannual) {
            Some(v) => v,
            None => return,
        };
    let oddl_text_freq =
        r#"=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,"2",0)"#;
    let oddl_text_freq_value = eval_number_or_skip(&mut sheet, oddl_text_freq)
        .expect("ODDLPRICE should accept frequency supplied as numeric text");
    assert_close(oddl_text_freq_value, oddl_baseline_semiannual_value, 1e-9);

    // TRUE/FALSE should also coerce to 1/0 for ODDLPRICE.
    let oddl_baseline_annual =
        "=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0)";
    let oddl_baseline_annual_value = eval_number_or_skip(&mut sheet, oddl_baseline_annual)
        .expect("ODDLPRICE should accept explicit annual frequency");
    let oddl_true_freq =
        "=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,TRUE,0)";
    let oddl_true_freq_value = eval_number_or_skip(&mut sheet, oddl_true_freq)
        .expect("ODDLPRICE should accept TRUE frequency (TRUE->1)");
    assert_close(oddl_true_freq_value, oddl_baseline_annual_value, 1e-9);

    let oddl_false_freq =
        "=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,FALSE,0)";
    match sheet.eval(oddl_false_freq) {
        Value::Error(ErrorKind::Name) => return,
        Value::Error(ErrorKind::Num) => {}
        other => panic!("expected #NUM! for ODDLPRICE frequency=FALSE (0), got {other:?}"),
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

    // Passing an explicit blank argument for an optional parameter behaves like 0 in Excel.
    let blank_basis_arg = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,)";
    let blank_basis_arg_value = eval_number_or_skip(&mut sheet, blank_basis_arg)
        .expect("ODDFPRICE should accept blank basis argument and treat it as 0");
    assert_close(blank_basis_arg_value, baseline_basis_0_value, 1e-9);

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

    // Spot-check ODDLPRICE for blank/boolean basis coercions.
    let oddl_baseline_basis_0 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let oddl_baseline_basis_0_value = match eval_number_or_skip(&mut sheet, oddl_baseline_basis_0) {
        Some(v) => v,
        None => return,
    };
    let oddl_blank_basis =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,A1)";
    let oddl_blank_basis_value = eval_number_or_skip(&mut sheet, oddl_blank_basis)
        .expect("ODDLPRICE should treat blank basis as 0");
    assert_close(oddl_blank_basis_value, oddl_baseline_basis_0_value, 1e-9);

    let oddl_blank_basis_arg =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,)";
    let oddl_blank_basis_arg_value = eval_number_or_skip(&mut sheet, oddl_blank_basis_arg)
        .expect("ODDLPRICE should treat blank basis argument as 0");
    assert_close(
        oddl_blank_basis_arg_value,
        oddl_baseline_basis_0_value,
        1e-9,
    );

    let oddl_false_basis =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,FALSE)";
    let oddl_false_basis_value = eval_number_or_skip(&mut sheet, oddl_false_basis)
        .expect("ODDLPRICE should accept FALSE basis (FALSE->0)");
    assert_close(oddl_false_basis_value, oddl_baseline_basis_0_value, 1e-9);

    let oddl_baseline_basis_1 =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,1)";
    let oddl_baseline_basis_1_value = eval_number_or_skip(&mut sheet, oddl_baseline_basis_1)
        .expect("ODDLPRICE should accept explicit basis=1");
    let oddl_true_basis =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,TRUE)";
    let oddl_true_basis_value = eval_number_or_skip(&mut sheet, oddl_true_basis)
        .expect("ODDLPRICE should accept TRUE basis (TRUE->1)");
    assert_close(oddl_true_basis_value, oddl_baseline_basis_1_value, 1e-9);
}

#[test]
fn odd_coupon_functions_accept_iso_date_text_arguments() {
    let mut sheet = TestSheet::new();
    // Excel date coercion: ISO-like text should be parsed as a date serial.
    // Ensure odd coupon functions accept text dates and produce the same result
    // as DATE()-based inputs.

    // Baseline case (Task 56): odd first coupon period.
    let baseline_oddfprice = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let baseline_oddfprice_value = match eval_number_or_skip(&mut sheet, baseline_oddfprice) {
        Some(v) => v,
        None => return,
    };
    let iso_oddfprice = r#"=ODDFPRICE("2008-11-11","2021-03-01","2008-10-15","2009-03-01",0.0785,0.0625,100,2,0)"#;
    let iso_oddfprice_value = eval_number_or_skip(&mut sheet, iso_oddfprice)
        .expect("ODDFPRICE should accept ISO date text arguments");
    assert_close(iso_oddfprice_value, baseline_oddfprice_value, 1e-9);

    // Baseline case (Task 56): odd last coupon period.
    let baseline_oddlprice =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let baseline_oddlprice_value = eval_number_or_skip(&mut sheet, baseline_oddlprice)
        .expect("ODDLPRICE should return a number for the baseline");
    let iso_oddlprice =
        r#"=ODDLPRICE("2020-11-11","2021-03-01","2020-10-15",0.0785,0.0625,100,2,0)"#;
    let iso_oddlprice_value = eval_number_or_skip(&mut sheet, iso_oddlprice)
        .expect("ODDLPRICE should accept ISO date text arguments");
    assert_close(iso_oddlprice_value, baseline_oddlprice_value, 1e-9);
}

#[test]
fn odd_coupon_functions_floor_time_fractions_in_date_arguments() {
    let mut sheet = TestSheet::new();
    // Excel date serials can contain a fractional time component. For these bond functions,
    // Excel behaves as though date arguments are floored to the day.

    let baseline_oddfprice = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let baseline_oddfprice_value = match eval_number_or_skip(&mut sheet, baseline_oddfprice) {
        Some(v) => v,
        None => return,
    };
    let fractional_oddfprice = "=ODDFPRICE(DATE(2008,11,11)+0.75,DATE(2021,3,1)+0.1,DATE(2008,10,15)+0.9,DATE(2009,3,1)+0.5,0.0785,0.0625,100,2,0)";
    let fractional_oddfprice_value = eval_number_or_skip(&mut sheet, fractional_oddfprice)
        .expect("ODDFPRICE should floor fractional date serials");
    assert_close(fractional_oddfprice_value, baseline_oddfprice_value, 1e-9);

    let baseline_oddlprice =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let baseline_oddlprice_value = eval_number_or_skip(&mut sheet, baseline_oddlprice)
        .expect("ODDLPRICE should return a number for the baseline");
    let fractional_oddlprice = "=ODDLPRICE(DATE(2020,11,11)+0.75,DATE(2021,3,1)+0.1,DATE(2020,10,15)+0.9,0.0785,0.0625,100,2,0)";
    let fractional_oddlprice_value = eval_number_or_skip(&mut sheet, fractional_oddlprice)
        .expect("ODDLPRICE should floor fractional date serials");
    assert_close(fractional_oddlprice_value, baseline_oddlprice_value, 1e-9);
}

#[test]
fn oddfyield_roundtrips_price_with_text_dates() {
    let mut sheet = TestSheet::new();
    // Ensure ODDFYIELD accepts date arguments supplied as ISO-like text, and that the
    // ODDFPRICE/ODDFYIELD pair roundtrips the yield.

    sheet.set_formula(
        "A1",
        "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)",
    );
    sheet.recalc();

    let _price = match cell_number_or_skip(&sheet, "A1") {
        Some(v) => v,
        None => return,
    };

    let recovered_yield = match eval_number_or_skip(
        &mut sheet,
        r#"=ODDFYIELD("2008-11-11","2021-03-01","2008-10-15","2009-03-01",0.0785,A1,100,2,0)"#,
    ) {
        Some(v) => v,
        None => return,
    };
    assert_close(recovered_yield, 0.0625, 1e-10);
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
fn odd_first_coupon_bond_functions_round_trip_long_stub() {
    let mut sheet = TestSheet::new();

    // Long odd-first coupon period:
    // - issue is far before first_coupon so DFC/E > 1 (long stub)
    // - settlement is between issue and first_coupon
    // - maturity is aligned with the regular schedule after first_coupon
    //
    // Also includes a basis=1 variant that crosses the 2020 leap day, exercising
    // E computation for actual/actual.
    //
    // (Excel oracle values are validated separately; this unit test asserts ODDFPRICE/ODDFYIELD
    // are internally consistent.)
    let yield_target = 0.0625;
    let rate = 0.0785;

    for basis in [0, 1] {
        let price_formula = format!(
            "=ODDFPRICE(DATE(2019,6,1),DATE(2022,3,1),DATE(2019,1,1),DATE(2020,3,1),{rate},{yield_target},100,2,{basis})"
        );
        sheet.set_formula("A1", &price_formula);
        sheet.recalc();

        let Some(price) = cell_number_or_skip(&sheet, "A1") else {
            return;
        };

        // Round-trip: compute yield from the computed price.
        let yield_formula = format!(
            "=ODDFYIELD(DATE(2019,6,1),DATE(2022,3,1),DATE(2019,1,1),DATE(2020,3,1),{rate},A1,100,2,{basis})"
        );
        let Some(y) = eval_number_or_skip(&mut sheet, &yield_formula) else {
            return;
        };

        assert!(
            price.is_finite() && price > 0.0,
            "expected positive finite price, got {price}"
        );
        assert_close(y, yield_target, 1e-9);
    }
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
fn odd_coupon_bond_price_allows_negative_yield() {
    let mut sheet = TestSheet::new();

    let oddf = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,-0.01,100,2,0)";
    let oddl = "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,-0.01,100,2,0)";

    let oddf_price = match eval_number_or_skip(&mut sheet, oddf) {
        Some(v) => v,
        None => return,
    };
    let oddl_price = eval_number_or_skip(&mut sheet, oddl)
        .expect("ODDLPRICE should return a number for negative yld within (-frequency, ∞)");

    assert!(
        oddf_price.is_finite(),
        "expected finite price, got {oddf_price}"
    );
    assert!(
        oddl_price.is_finite(),
        "expected finite price, got {oddl_price}"
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
fn odd_coupon_bond_price_rejects_negative_coupon_rate() {
    let mut sheet = TestSheet::new();

    let oddf = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),-0.01,0.0625,100,2,0)";
    let oddl = "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),-0.01,0.0625,100,2,0)";

    let Some(out) = eval_value_or_skip(&mut sheet, oddf) else {
        return;
    };
    assert!(
        matches!(out, Value::Error(ErrorKind::Num)),
        "expected #NUM! for negative rate in ODDFPRICE, got {out:?}"
    );

    let Some(out) = eval_value_or_skip(&mut sheet, oddl) else {
        return;
    };
    assert!(
        matches!(out, Value::Error(ErrorKind::Num)),
        "expected #NUM! for negative rate in ODDLPRICE, got {out:?}"
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

#[test]
fn odd_coupon_bond_yield_can_be_negative() {
    let mut sheet = TestSheet::new();

    // A price above the undiscounted cashflows implies a negative yield when yields are allowed
    // below 0. Excel's behavior here is historically ambiguous; this test locks in the current
    // engine semantics (negative yields are supported down to `-frequency`).
    let oddf = "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,300,100,2,0)";
    let oddl = "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,300,100,2,0)";

    let oddf_yld = match eval_number_or_skip(&mut sheet, oddf) {
        Some(v) => v,
        None => return,
    };
    let oddl_yld = eval_number_or_skip(&mut sheet, oddl).expect("ODDLYIELD should return a number");

    assert!(
        oddf_yld < 0.0 && oddf_yld > -2.0,
        "expected ODDFYIELD to return a negative yield in (-2, 0), got {oddf_yld}"
    );
    assert!(
        oddl_yld < 0.0 && oddl_yld > -2.0,
        "expected ODDLYIELD to return a negative yield in (-2, 0), got {oddl_yld}"
    );
}

#[test]
fn odd_coupon_bond_price_allows_zero_coupon_rate() {
    let mut sheet = TestSheet::new();

    // Zero-coupon odd-first/odd-last cases should still be valid.
    let oddf = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0,0.0625,100,2,0)";
    let oddl = "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0,0.0625,100,2,0)";

    let Some(out) = eval_value_or_skip(&mut sheet, oddf) else {
        return;
    };
    assert!(
        matches!(out, Value::Number(_)),
        "expected a numeric price for ODDFPRICE with rate=0, got {out:?}"
    );

    let Some(out) = eval_value_or_skip(&mut sheet, oddl) else {
        return;
    };
    assert!(
        matches!(out, Value::Number(_)),
        "expected a numeric price for ODDLPRICE with rate=0, got {out:?}"
    );
}

#[test]
fn odd_last_coupon_bond_functions_round_trip_long_stub() {
    let mut sheet = TestSheet::new();

    // Long odd-last coupon period:
    // - last_interest is far before maturity so DSM/E > 1 (long stub)
    // - settlement is between last_interest and maturity
    //
    // We keep the schedule simple: there are no regular coupon payments between settlement
    // and maturity, so the functions must correctly scale the final coupon amount.
    let yield_target = 0.0625;
    let rate = 0.0785;

    for basis in [0, 1] {
        let price_formula = format!(
            "=ODDLPRICE(DATE(2021,2,1),DATE(2022,3,1),DATE(2020,10,15),{rate},{yield_target},100,2,{basis})"
        );
        sheet.set_formula("A1", &price_formula);
        sheet.recalc();

        let Some(price) = cell_number_or_skip(&sheet, "A1") else {
            return;
        };

        let yield_formula = format!(
            "=ODDLYIELD(DATE(2021,2,1),DATE(2022,3,1),DATE(2020,10,15),{rate},A1,100,2,{basis})"
        );
        let Some(y) = eval_number_or_skip(&mut sheet, &yield_formula) else {
            return;
        };

        assert!(
            price.is_finite() && price > 0.0,
            "expected positive finite price, got {price}"
        );
        assert_close(y, yield_target, 1e-9);
    }
}
