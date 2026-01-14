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
    assert_eq!(
        eval(&mut engine, "=LAMBDA(x,x)"),
        Value::Error(ErrorKind::Calc)
    );
}

#[test]
fn lambda_used_as_number_returns_value_error() {
    let mut engine = Engine::new();
    assert_eq!(
        eval(&mut engine, "=1+LAMBDA(x,x)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn lambda_used_in_lookup_functions_returns_value_error() {
    let mut engine = Engine::new();
    engine.set_cell_value("Sheet1", "B1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "B2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "C2", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "B4", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "C4", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "B5", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "C5", 20.0).unwrap();
    engine.set_cell_value("Sheet1", "D1", 1.0).unwrap();
    engine.set_cell_value("Sheet1", "D2", 2.0).unwrap();
    engine.set_cell_value("Sheet1", "E1", 10.0).unwrap();
    engine.set_cell_value("Sheet1", "E2", 20.0).unwrap();

    assert_eq!(
        eval(&mut engine, "=MATCH(LAMBDA(x,x),{1,2},0)"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        eval(&mut engine, "=XMATCH(LAMBDA(x,x),{1,2})"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        eval(&mut engine, r#"=XLOOKUP(LAMBDA(x,x),D1:D2,E1:E2,"no")"#),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        eval(&mut engine, "=XLOOKUP(LAMBDA(x,x),{1,2},{3,4})"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        eval(&mut engine, "=VLOOKUP(LAMBDA(x,x),B1:C2,2,FALSE)"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        eval(&mut engine, "=HLOOKUP(LAMBDA(x,x),B4:C5,2,FALSE)"),
        Value::Error(ErrorKind::Value)
    );
}
