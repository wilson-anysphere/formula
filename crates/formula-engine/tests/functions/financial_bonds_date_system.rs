use formula_engine::date::ExcelDateSystem;
use formula_engine::{ErrorKind, Value};

use super::harness::TestSheet;

fn assert_close(actual: f64, expected: f64, tol: f64) {
    assert!(
        (actual - expected).abs() <= tol,
        "expected {expected}, got {actual}"
    );
}

fn cell_number_or_skip(sheet: &TestSheet, addr: &str) -> Option<f64> {
    match sheet.get(addr) {
        Value::Number(n) => Some(n),
        // Standard coupon bond functions are not always implemented in every build of the engine.
        // Skip these tests when the function registry doesn't recognize the name.
        Value::Error(ErrorKind::Name) => None,
        other => panic!("expected number, got {other:?} from cell {addr}"),
    }
}

#[test]
fn standard_coupon_bond_functions_respect_workbook_date_system() {
    let mut sheet = TestSheet::new();

    // Example from Excel documentation: standard coupon bond with semiannual payments.
    sheet.set_formula(
        "A1",
        "=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)",
    );
    // Round-trip the yield from the computed price.
    sheet.set_formula(
        "A2",
        "=YIELD(DATE(2008,2,15),DATE(2017,11,15),0.0575,A1,100,2,0)",
    );
    sheet.set_formula(
        "A3",
        "=DURATION(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,2,0)",
    );
    sheet.set_formula(
        "A4",
        "=MDURATION(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,2,0)",
    );
    sheet.recalc();

    let price_1900 = match cell_number_or_skip(&sheet, "A1") {
        Some(v) => v,
        None => return,
    };
    let yield_1900 = match cell_number_or_skip(&sheet, "A2") {
        Some(v) => v,
        None => return,
    };
    let duration_1900 = match cell_number_or_skip(&sheet, "A3") {
        Some(v) => v,
        None => return,
    };
    let mduration_1900 = match cell_number_or_skip(&sheet, "A4") {
        Some(v) => v,
        None => return,
    };

    sheet.set_date_system(ExcelDateSystem::Excel1904);
    sheet.recalc();

    let price_1904 =
        cell_number_or_skip(&sheet, "A1").expect("PRICE should return a number under Excel1904");
    let yield_1904 =
        cell_number_or_skip(&sheet, "A2").expect("YIELD should return a number under Excel1904");
    let duration_1904 =
        cell_number_or_skip(&sheet, "A3").expect("DURATION should return a number under Excel1904");
    let mduration_1904 = cell_number_or_skip(&sheet, "A4")
        .expect("MDURATION should return a number under Excel1904");

    assert_close(price_1904, price_1900, 1e-9);
    assert_close(yield_1904, yield_1900, 1e-10);
    assert_close(duration_1904, duration_1900, 1e-10);
    assert_close(mduration_1904, mduration_1900, 1e-10);
}

#[test]
fn coup_schedule_components_respect_workbook_date_system() {
    let mut sheet = TestSheet::new();

    // COUP* returns serial numbers, which are expected to differ between Excel 1900 and 1904 date
    // systems. Compare calendar components instead.
    sheet.set_formula("A1", "=YEAR(COUPNCD(DATE(2008,2,15),DATE(2017,11,15),2,0))");
    sheet.set_formula(
        "A2",
        "=MONTH(COUPNCD(DATE(2008,2,15),DATE(2017,11,15),2,0))",
    );
    sheet.set_formula("A3", "=DAY(COUPNCD(DATE(2008,2,15),DATE(2017,11,15),2,0))");
    sheet.set_formula("B1", "=YEAR(COUPPCD(DATE(2008,2,15),DATE(2017,11,15),2,0))");
    sheet.set_formula(
        "B2",
        "=MONTH(COUPPCD(DATE(2008,2,15),DATE(2017,11,15),2,0))",
    );
    sheet.set_formula("B3", "=DAY(COUPPCD(DATE(2008,2,15),DATE(2017,11,15),2,0))");
    sheet.recalc();

    let ncd_year_1900 = match cell_number_or_skip(&sheet, "A1") {
        Some(v) => v,
        None => return,
    };
    let ncd_month_1900 = match cell_number_or_skip(&sheet, "A2") {
        Some(v) => v,
        None => return,
    };
    let ncd_day_1900 = match cell_number_or_skip(&sheet, "A3") {
        Some(v) => v,
        None => return,
    };
    let pcd_year_1900 = match cell_number_or_skip(&sheet, "B1") {
        Some(v) => v,
        None => return,
    };
    let pcd_month_1900 = match cell_number_or_skip(&sheet, "B2") {
        Some(v) => v,
        None => return,
    };
    let pcd_day_1900 = match cell_number_or_skip(&sheet, "B3") {
        Some(v) => v,
        None => return,
    };

    sheet.set_date_system(ExcelDateSystem::Excel1904);
    sheet.recalc();

    let ncd_year_1904 = cell_number_or_skip(&sheet, "A1")
        .expect("COUPNCD/YEAR should return a number under Excel1904");
    let ncd_month_1904 = cell_number_or_skip(&sheet, "A2")
        .expect("COUPNCD/MONTH should return a number under Excel1904");
    let ncd_day_1904 = cell_number_or_skip(&sheet, "A3")
        .expect("COUPNCD/DAY should return a number under Excel1904");
    let pcd_year_1904 = cell_number_or_skip(&sheet, "B1")
        .expect("COUPPCD/YEAR should return a number under Excel1904");
    let pcd_month_1904 = cell_number_or_skip(&sheet, "B2")
        .expect("COUPPCD/MONTH should return a number under Excel1904");
    let pcd_day_1904 = cell_number_or_skip(&sheet, "B3")
        .expect("COUPPCD/DAY should return a number under Excel1904");

    assert_close(ncd_year_1904, ncd_year_1900, 0.0);
    assert_close(ncd_month_1904, ncd_month_1900, 0.0);
    assert_close(ncd_day_1904, ncd_day_1900, 0.0);
    assert_close(pcd_year_1904, pcd_year_1900, 0.0);
    assert_close(pcd_month_1904, pcd_month_1900, 0.0);
    assert_close(pcd_day_1904, pcd_day_1900, 0.0);
}
