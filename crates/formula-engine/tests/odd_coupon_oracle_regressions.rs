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
fn oddfprice_basis4_matches_excel_oracle() {
    // Case id: oddfprice_basis4_0679b2004f6c
    let v = eval_formula(
        "=ODDFPRICE(DATE(2020,1,20),DATE(2021,8,30),DATE(2020,1,15),DATE(2020,2,29),0.08,0.075,100,2,4)",
    );
    assert_number_close(v, 100.75597490147861, 1e-6);
}

#[test]
fn oddfyield_basis4_matches_excel_oracle() {
    // Case id: oddfyield_basis4_b8c7df4a447f
    let v = eval_formula(
        "=ODDFYIELD(DATE(2020,1,20),DATE(2021,8,30),DATE(2020,1,15),DATE(2020,2,29),0.08,98,100,2,4)",
    );
    assert_number_close(v, 0.0937828362065177, 1e-6);
}

#[test]
fn oddlprice_basis4_matches_excel_oracle() {
    // Case id: oddlprice_basis4_546c496042bd
    let v = eval_formula(
        "=ODDLPRICE(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.06,0.055,100,4,4)",
    );
    assert_number_close(v, 100.06126030022395, 1e-6);
}

#[test]
fn oddlyield_basis4_matches_excel_oracle() {
    // Case id: oddlyield_basis4_45ff4139618f
    let v =
        eval_formula("=ODDLYIELD(DATE(2021,7,1),DATE(2021,8,15),DATE(2021,6,15),0.06,98,100,4,4)");
    assert_number_close(v, 0.23089149550037363, 1e-6);
}
