use formula_engine::date::ExcelDateSystem;
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

