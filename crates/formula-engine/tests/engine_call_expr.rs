use formula_engine::{Engine, Value};

#[test]
fn engine_compiles_and_evaluates_lambda_invocation_call_expr() {
    let mut engine = Engine::new();
    engine
        .set_cell_formula("Sheet1", "A1", "=LAMBDA(x,x+1)(3)")
        .expect("set formula");
    engine.recalculate_single_threaded();
    assert_eq!(engine.get_cell_value("Sheet1", "A1"), Value::Number(4.0));
}
