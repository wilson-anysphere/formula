use formula_engine::{Engine, ErrorKind, NameDefinition, NameScope, Value};

fn eval(engine: &mut Engine, formula: &str) -> Value {
    engine
        .set_cell_formula("Sheet1", "A1", formula)
        .expect("set formula");
    engine.recalculate();
    engine.get_cell_value("Sheet1", "A1")
}

#[test]
fn let_binds_names_sequentially() {
    let mut engine = Engine::new();
    assert_eq!(eval(&mut engine, "=LET(x,1,x+1)"), Value::Number(2.0));
    assert_eq!(eval(&mut engine, "=LET(x,1,y,x+1,y+1)"), Value::Number(3.0));
}

#[test]
fn let_shadows_defined_names() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "X",
            NameScope::Workbook,
            NameDefinition::Constant(Value::Number(10.0)),
        )
        .expect("define name");

    assert_eq!(eval(&mut engine, "=LET(X,1,X+1)"), Value::Number(2.0));
}

#[test]
fn lambda_invocation_syntax_and_named_fallback() {
    let mut engine = Engine::new();
    assert_eq!(eval(&mut engine, "=LAMBDA(x,x+1)(3)"), Value::Number(4.0));
    assert_eq!(
        eval(&mut engine, "=LET(f, LAMBDA(x,x+1), f(3))"),
        Value::Number(4.0)
    );
}

#[test]
fn top_level_lambda_returns_calc_error() {
    let mut engine = Engine::new();
    assert_eq!(eval(&mut engine, "=LAMBDA(x,x)"), Value::Error(ErrorKind::Calc));
}

