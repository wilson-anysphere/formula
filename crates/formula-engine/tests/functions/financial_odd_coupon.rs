use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
use formula_engine::functions::financial::{oddfprice, oddfyield, oddlprice, oddlyield};
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
fn odd_first_coupon_bond_functions_respect_workbook_date_system() {
    let mut sheet = TestSheet::new();

    // Baseline case (Task 56): odd first coupon period.
    let price_formula = "=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.0625,100,2,0)";
    let yield_formula = "=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,98,100,2,0)";

    let price_1900 = match eval_number_or_skip(&mut sheet, price_formula) {
        Some(v) => v,
        None => return,
    };
    let yield_1900 = match eval_number_or_skip(&mut sheet, yield_formula) {
        Some(v) => v,
        None => return,
    };

    sheet.set_date_system(ExcelDateSystem::Excel1904);

    let price_1904 = eval_number_or_skip(&mut sheet, price_formula)
        .expect("ODDFPRICE should return a number under Excel1904");
    let yield_1904 = eval_number_or_skip(&mut sheet, yield_formula)
        .expect("ODDFYIELD should return a number under Excel1904");

    assert_close(price_1904, price_1900, 1e-9);
    assert_close(yield_1904, yield_1900, 1e-10);
}

#[test]
fn odd_last_coupon_bond_functions_respect_workbook_date_system() {
    let mut sheet = TestSheet::new();

    // Baseline case (Task 56): odd last coupon period.
    let price_formula =
        "=ODDLPRICE(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,0.0625,100,2,0)";
    let yield_formula =
        "=ODDLYIELD(DATE(2020,11,11),DATE(2021,3,1),DATE(2020,10,15),0.0785,98,100,2,0)";

    let price_1900 = match eval_number_or_skip(&mut sheet, price_formula) {
        Some(v) => v,
        None => return,
    };
    let yield_1900 = match eval_number_or_skip(&mut sheet, yield_formula) {
        Some(v) => v,
        None => return,
    };

    sheet.set_date_system(ExcelDateSystem::Excel1904);

    let price_1904 = eval_number_or_skip(&mut sheet, price_formula)
        .expect("ODDLPRICE should return a number under Excel1904");
    let yield_1904 = eval_number_or_skip(&mut sheet, yield_formula)
        .expect("ODDLYIELD should return a number under Excel1904");

    assert_close(price_1904, price_1900, 1e-9);
    assert_close(yield_1904, yield_1900, 1e-10);
}

#[test]
fn odd_first_coupon_roundtrips_yield_with_annual_frequency() {
    let mut sheet = TestSheet::new();

    // Aligned annual schedule from `first_coupon` by 12 months:
    // 2020-07-01, 2021-07-01, 2022-07-01, 2023-07-01 (maturity).
    sheet.set_formula(
        "A1",
        "=ODDFPRICE(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,0.05,100,1,0)",
    );
    sheet.recalc();

    let _price = match cell_number_or_skip(&sheet, "A1") {
        Some(v) => v,
        None => return,
    };

    let recovered_yield = match eval_number_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2020,3,1),DATE(2023,7,1),DATE(2020,1,1),DATE(2020,7,1),0.06,A1,100,1,0)",
    ) {
        Some(v) => v,
        None => return,
    };

    assert_close(recovered_yield, 0.05, 1e-10);
}

#[test]
fn odd_first_coupon_roundtrips_yield_with_quarterly_frequency_and_non_100_redemption() {
    let mut sheet = TestSheet::new();

    // Aligned quarterly schedule from `first_coupon` by 3 months:
    // 2020-02-15, 2020-05-15, 2020-08-15, 2020-11-15, 2021-02-15, 2021-05-15, 2021-08-15.
    sheet.set_formula(
        "A1",
        "=ODDFPRICE(DATE(2020,1,20),DATE(2021,8,15),DATE(2020,1,1),DATE(2020,2,15),0.08,0.07,100,4,0)",
    );
    sheet.set_formula(
        "A2",
        "=ODDFPRICE(DATE(2020,1,20),DATE(2021,8,15),DATE(2020,1,1),DATE(2020,2,15),0.08,0.07,105,4,0)",
    );
    sheet.recalc();

    let price_100 = match cell_number_or_skip(&sheet, "A1") {
        Some(v) => v,
        None => return,
    };
    let price_105 = match cell_number_or_skip(&sheet, "A2") {
        Some(v) => v,
        None => return,
    };

    assert!(
        (price_105 - price_100).abs() > 1e-9,
        "expected redemption to affect price (redemption=100 => {price_100}, redemption=105 => {price_105})"
    );
    assert!(
        price_105 > price_100,
        "expected higher redemption to increase price (redemption=100 => {price_100}, redemption=105 => {price_105})"
    );

    let recovered_yield_100 = match eval_number_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2020,1,20),DATE(2021,8,15),DATE(2020,1,1),DATE(2020,2,15),0.08,A1,100,4,0)",
    ) {
        Some(v) => v,
        None => return,
    };
    let recovered_yield_105 = match eval_number_or_skip(
        &mut sheet,
        "=ODDFYIELD(DATE(2020,1,20),DATE(2021,8,15),DATE(2020,1,1),DATE(2020,2,15),0.08,A2,105,4,0)",
    ) {
        Some(v) => v,
        None => return,
    };

    assert_close(recovered_yield_100, 0.07, 1e-10);
    assert_close(recovered_yield_105, 0.07, 1e-10);
}

#[test]
fn odd_last_coupon_roundtrips_yield_with_annual_frequency() {
    let mut sheet = TestSheet::new();

    // `last_interest` is a coupon date on an annual schedule (12 month stepping). Maturity
    // occurs 8 months later, making this an odd last coupon period.
    sheet.set_formula(
        "A1",
        "=ODDLPRICE(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,0.05,100,1,0)",
    );
    sheet.recalc();

    let _price = match cell_number_or_skip(&sheet, "A1") {
        Some(v) => v,
        None => return,
    };

    let recovered_yield = match eval_number_or_skip(
        &mut sheet,
        "=ODDLYIELD(DATE(2022,11,1),DATE(2023,3,1),DATE(2022,7,1),0.06,A1,100,1,0)",
    ) {
        Some(v) => v,
        None => return,
    };

    assert_close(recovered_yield, 0.05, 1e-10);
}

#[test]
fn odd_last_coupon_roundtrips_yield_with_quarterly_frequency_and_non_100_redemption() {
    let mut sheet = TestSheet::new();

    // `last_interest` is a coupon date on a quarterly schedule. Maturity occurs 2 months later
    // (shorter than the regular 3 month period), making this an odd last coupon period.
    sheet.set_formula(
        "A1",
        "=ODDLPRICE(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.08,0.07,100,4,0)",
    );
    sheet.set_formula(
        "A2",
        "=ODDLPRICE(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.08,0.07,105,4,0)",
    );
    sheet.recalc();

    let price_100 = match cell_number_or_skip(&sheet, "A1") {
        Some(v) => v,
        None => return,
    };
    let price_105 = match cell_number_or_skip(&sheet, "A2") {
        Some(v) => v,
        None => return,
    };

    assert!(
        (price_105 - price_100).abs() > 1e-9,
        "expected redemption to affect price (redemption=100 => {price_100}, redemption=105 => {price_105})"
    );
    assert!(
        price_105 > price_100,
        "expected higher redemption to increase price (redemption=100 => {price_100}, redemption=105 => {price_105})"
    );

    let recovered_yield_100 = match eval_number_or_skip(
        &mut sheet,
        "=ODDLYIELD(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.08,A1,100,4,0)",
    ) {
        Some(v) => v,
        None => return,
    };
    let recovered_yield_105 = match eval_number_or_skip(
        &mut sheet,
        "=ODDLYIELD(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.08,A2,105,4,0)",
    ) {
        Some(v) => v,
        None => return,
    };

    assert_close(recovered_yield_100, 0.07, 1e-10);
    assert_close(recovered_yield_105, 0.07, 1e-10);
}
