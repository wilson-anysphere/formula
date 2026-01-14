use formula_engine::date::ExcelDateSystem;
use formula_engine::{Engine, ErrorKind, Value};

fn eval_formula_in_system(formula: &str, system: ExcelDateSystem) -> Value {
    let mut engine = Engine::new();
    engine.set_date_system(system);
    engine
        .set_cell_formula("Sheet1", "A1", formula)
        .expect("set formula");
    engine.recalculate_single_threaded();
    engine.get_cell_value("Sheet1", "A1")
}

fn eval_formula(formula: &str) -> Value {
    eval_formula_in_system(formula, ExcelDateSystem::EXCEL_1900)
}

fn eval_formula_1904(formula: &str) -> Value {
    eval_formula_in_system(formula, ExcelDateSystem::Excel1904)
}

fn assert_number_close(v: Value, expected: f64, tol: f64) {
    match v {
        Value::Number(n) => assert!((n - expected).abs() <= tol, "expected {expected}, got {n}"),
        other => panic!("expected number, got {other:?}"),
    }
}

fn assert_number_close_in_both_systems(formula: &str, expected: f64, tol: f64) {
    assert_number_close(eval_formula(formula), expected, tol);
    assert_number_close(eval_formula_1904(formula), expected, tol);
}

fn assert_num_error_in_both_systems(formula: &str) {
    let v = eval_formula(formula);
    assert_eq!(
        v,
        Value::Error(ErrorKind::Num),
        "expected #NUM! under Excel1900, got {v:?}"
    );
    let v = eval_formula_1904(formula);
    assert_eq!(
        v,
        Value::Error(ErrorKind::Num),
        "expected #NUM! under Excel1904, got {v:?}"
    );
}

#[test]
fn oddfprice_allows_issue_equal_settlement() {
    // Pinned by excel-oracle boundary cases.
    assert_number_close_in_both_systems(
        "=ODDFPRICE(DATE(2020,1,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0)",
        104.49129250312109,
        1e-6,
    );
}

#[test]
fn oddfyield_allows_issue_equal_settlement() {
    // Pinned by excel-oracle boundary cases.
    assert_number_close_in_both_systems(
        "=ODDFYIELD(DATE(2020,1,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,ODDFPRICE(DATE(2020,1,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0),100,2,0)",
        0.04000000000000014,
        1e-6,
    );
}

#[test]
fn oddfprice_allows_settlement_equal_first_coupon() {
    // Pinned by excel-oracle boundary cases.
    assert_number_close_in_both_systems(
        "=ODDFPRICE(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0)",
        104.08111835318353,
        1e-6,
    );
}

#[test]
fn oddfyield_allows_settlement_equal_first_coupon() {
    // Pinned by excel-oracle boundary cases.
    assert_number_close_in_both_systems(
        "=ODDFYIELD(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,ODDFPRICE(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0),100,2,0)",
        0.039999999999979614,
        1e-6,
    );
}

#[test]
fn oddfprice_allows_first_coupon_equal_maturity() {
    // Pinned by excel-oracle boundary cases.
    assert_number_close_in_both_systems(
        "=ODDFPRICE(DATE(2020,3,1),DATE(2020,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0)",
        100.3223801273643,
        1e-6,
    );
}

#[test]
fn oddfyield_allows_first_coupon_equal_maturity() {
    // Pinned by excel-oracle boundary cases.
    assert_number_close_in_both_systems(
        "=ODDFYIELD(DATE(2020,3,1),DATE(2020,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,ODDFPRICE(DATE(2020,3,1),DATE(2020,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0),100,2,0)",
        0.03999999999999963,
        1e-6,
    );
}

#[test]
fn oddfprice_rejects_settlement_equal_maturity() {
    // Even when `first_coupon == maturity` is valid, `settlement` must still be strictly before
    // maturity.
    assert_num_error_in_both_systems(
        "=ODDFPRICE(DATE(2020,7,1),DATE(2020,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0)",
    );
}

#[test]
fn oddfyield_rejects_settlement_equal_maturity() {
    assert_num_error_in_both_systems(
        "=ODDFYIELD(DATE(2020,7,1),DATE(2020,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,99,100,2,0)",
    );
}

#[test]
fn oddfprice_rejects_settlement_equal_maturity_long_term() {
    assert_num_error_in_both_systems(
        "=ODDFPRICE(DATE(2025,1,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2025,1,1),0.05,0.04,100,2,0)",
    );
}

#[test]
fn oddfyield_rejects_settlement_equal_maturity_long_term() {
    assert_num_error_in_both_systems(
        "=ODDFYIELD(DATE(2025,1,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2025,1,1),0.05,99,100,2,0)",
    );
}

#[test]
fn oddfprice_rejects_issue_equal_first_coupon() {
    assert_num_error_in_both_systems(
        "=ODDFPRICE(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,7,1),DATE(2020,7,1),0.05,0.04,100,2,0)",
    );
}

#[test]
fn oddfyield_rejects_issue_equal_first_coupon() {
    assert_num_error_in_both_systems(
        "=ODDFYIELD(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,7,1),DATE(2020,7,1),0.05,99,100,2,0)",
    );
}

#[test]
fn oddlprice_allows_settlement_equal_last_interest() {
    // Pinned by excel-oracle boundary cases.
    assert_number_close_in_both_systems(
        "=ODDLPRICE(DATE(2024,7,1),DATE(2025,1,1),DATE(2024,7,1),0.05,0.04,100,2,0)",
        100.49019607843137,
        1e-6,
    );
}

#[test]
fn oddlyield_allows_settlement_equal_last_interest() {
    // Pinned by excel-oracle boundary cases.
    assert_number_close_in_both_systems(
        "=ODDLYIELD(DATE(2024,7,1),DATE(2025,1,1),DATE(2024,7,1),0.05,ODDLPRICE(DATE(2024,7,1),DATE(2025,1,1),DATE(2024,7,1),0.05,0.04,100,2,0),100,2,0)",
        0.039999999999999813,
        1e-6,
    );
}

#[test]
fn oddlprice_rejects_last_interest_equal_maturity() {
    assert_num_error_in_both_systems(
        "=ODDLPRICE(DATE(2025,1,1),DATE(2025,1,1),DATE(2025,1,1),0.05,0.04,100,2,0)",
    );
}

#[test]
fn oddlyield_rejects_last_interest_equal_maturity() {
    assert_num_error_in_both_systems(
        "=ODDLYIELD(DATE(2025,1,1),DATE(2025,1,1),DATE(2025,1,1),0.05,99,100,2,0)",
    );
}
