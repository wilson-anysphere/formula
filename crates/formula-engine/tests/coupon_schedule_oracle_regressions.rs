use formula_engine::{Engine, Value};

fn eval_formula(formula: &str) -> Value {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", formula)
        .expect("set formula");
    engine.recalculate_single_threaded();
    engine.get_cell_value("Sheet1", "A1")
}

fn assert_number_close(v: Value, expected: f64, tol: f64) {
    match v {
        Value::Number(n) => assert!((n - expected).abs() <= tol, "expected {expected}, got {n}"),
        other => panic!("expected number, got {other:?}"),
    }
}

#[test]
fn coupdays_eom_apr30_basis1_matches_excel_oracle() {
    // Case id: coupdays_eom_apr30_f63ac6d593a1
    let v = eval_formula("=COUPDAYS(DATE(2020,2,15),DATE(2020,4,30),4,1)");
    assert_number_close(v, 90.0, 0.0);
}

#[test]
fn couppcd_eom_feb28_basis1_matches_excel_oracle() {
    // Case id: couppcd_eom_feb28_8ebc724a184e
    let v = eval_formula("=COUPPCD(DATE(2020,11,15),DATE(2021,2,28),2,1)");
    // Serial date for 2020-08-31 in the Excel 1900 system.
    assert_number_close(v, 44074.0, 0.0);
}

#[test]
fn coupdays_eom_feb28_basis1_matches_excel_oracle() {
    // Case id: coupdays_eom_feb28_69d7ccd41d35
    let v = eval_formula("=COUPDAYS(DATE(2020,11,15),DATE(2021,2,28),2,1)");
    assert_number_close(v, 181.0, 0.0);
}

#[test]
fn coupdaybs_basis4_eom_feb28_matches_excel_oracle() {
    // Case id: coupdaybs_b4_eom_feb28_baf6f7ea73cf
    let v = eval_formula("=COUPDAYBS(DATE(2020,11,15),DATE(2021,2,28),2,4)");
    assert_number_close(v, 75.0, 0.0);
}

#[test]
fn coupdays_basis4_eom_feb28_matches_excel_oracle() {
    // Case id: coupdays_b4_eom_feb28_aa1d1a6a6133
    let v = eval_formula("=COUPDAYS(DATE(2020,11,15),DATE(2021,2,28),2,4)");
    assert_number_close(v, 180.0, 0.0);
}

#[test]
fn coupdaysnc_basis4_eom_feb28_matches_excel_oracle() {
    // Case id: coupdaysnc_b4_eom_feb28_d49eb1080971
    let v = eval_formula("=COUPDAYSNC(DATE(2020,11,15),DATE(2021,2,28),2,4)");
    assert_number_close(v, 105.0, 0.0);
}
