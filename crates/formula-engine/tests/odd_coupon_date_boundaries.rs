use formula_engine::{Engine, ErrorKind, Value};

fn eval_formula(formula: &str) -> Value {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", formula)
        .expect("set formula");
    engine.recalculate_single_threaded();
    engine.get_cell_value("Sheet1", "A1")
}

#[test]
fn oddfprice_allows_issue_equal_settlement() {
    let v = eval_formula("=ODDFPRICE(DATE(2020,1,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0)");
    assert!(matches!(v, Value::Number(n) if n.is_finite()), "got {v:?}");
}

#[test]
fn oddfyield_allows_issue_equal_settlement() {
    let v = eval_formula(
        "=ODDFYIELD(DATE(2020,1,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,ODDFPRICE(DATE(2020,1,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0),100,2,0)",
    );
    assert!(
        matches!(v, Value::Number(n) if (n - 0.04).abs() <= 1e-6),
        "expected yield ~0.04, got {v:?}"
    );
}

#[test]
fn oddfprice_allows_settlement_equal_first_coupon() {
    let v = eval_formula("=ODDFPRICE(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0)");
    assert!(matches!(v, Value::Number(n) if n.is_finite()), "got {v:?}");
}

#[test]
fn oddfyield_allows_settlement_equal_first_coupon() {
    let v = eval_formula(
        "=ODDFYIELD(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,ODDFPRICE(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0),100,2,0)",
    );
    assert!(
        matches!(v, Value::Number(n) if (n - 0.04).abs() <= 1e-6),
        "expected yield ~0.04, got {v:?}"
    );
}

#[test]
fn oddfprice_allows_first_coupon_equal_maturity() {
    let v = eval_formula(
        "=ODDFPRICE(DATE(2020,3,1),DATE(2020,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0)",
    );
    assert!(matches!(v, Value::Number(n) if n.is_finite()));
}

#[test]
fn oddfyield_allows_first_coupon_equal_maturity() {
    let v = eval_formula("=ODDFYIELD(DATE(2020,3,1),DATE(2020,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,ODDFPRICE(DATE(2020,3,1),DATE(2020,7,1),DATE(2020,1,1),DATE(2020,7,1),0.05,0.04,100,2,0),100,2,0)");
    assert!(matches!(v, Value::Number(n) if n.is_finite()));
}

#[test]
fn oddfprice_rejects_issue_equal_first_coupon() {
    let v = eval_formula(
        "=ODDFPRICE(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,7,1),DATE(2020,7,1),0.05,0.04,100,2,0)",
    );
    assert_eq!(v, Value::Error(ErrorKind::Num));
}

#[test]
fn oddfyield_rejects_issue_equal_first_coupon() {
    let v = eval_formula(
        "=ODDFYIELD(DATE(2020,7,1),DATE(2025,1,1),DATE(2020,7,1),DATE(2020,7,1),0.05,99,100,2,0)",
    );
    assert_eq!(v, Value::Error(ErrorKind::Num));
}

#[test]
fn oddlprice_allows_settlement_equal_last_interest() {
    let v =
        eval_formula("=ODDLPRICE(DATE(2024,7,1),DATE(2025,1,1),DATE(2024,7,1),0.05,0.04,100,2,0)");
    assert!(matches!(v, Value::Number(n) if n.is_finite()), "got {v:?}");
}

#[test]
fn oddlyield_allows_settlement_equal_last_interest() {
    let v = eval_formula("=ODDLYIELD(DATE(2024,7,1),DATE(2025,1,1),DATE(2024,7,1),0.05,ODDLPRICE(DATE(2024,7,1),DATE(2025,1,1),DATE(2024,7,1),0.05,0.04,100,2,0),100,2,0)");
    assert!(matches!(v, Value::Number(n) if n.is_finite()), "got {v:?}");
}

#[test]
fn oddlprice_rejects_last_interest_equal_maturity() {
    let v =
        eval_formula("=ODDLPRICE(DATE(2025,1,1),DATE(2025,1,1),DATE(2025,1,1),0.05,0.04,100,2,0)");
    assert_eq!(v, Value::Error(ErrorKind::Num));
}

#[test]
fn oddlyield_rejects_last_interest_equal_maturity() {
    let v =
        eval_formula("=ODDLYIELD(DATE(2025,1,1),DATE(2025,1,1),DATE(2025,1,1),0.05,99,100,2,0)");
    assert_eq!(v, Value::Error(ErrorKind::Num));
}
