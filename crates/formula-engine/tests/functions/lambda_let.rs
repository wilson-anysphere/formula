use formula_engine::{Engine, ErrorKind, NameDefinition, NameScope, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn let_binds_values_left_to_right() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=LET(a,1,a)"), 1.0);
    assert_number(&sheet.eval("=LET(a,1,b,a+1,b)"), 2.0);
    assert_number(&sheet.eval("=LET(a,2,b,a*3,c,b+1,c)"), 7.0);
}

#[test]
fn let_rejects_non_identifier_names() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=LET(1,2,3)"), Value::Error(ErrorKind::Value));
    assert_eq!(sheet.eval("=LET(A1,2,3)"), Value::Error(ErrorKind::Value));
}

#[test]
fn lambda_can_be_called_via_let_binding() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=LET(f,LAMBDA(x,x+1),f(2))"), 3.0);
}

#[test]
fn lambda_can_be_called_inline() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=LAMBDA(x,x+1)(2)"), 3.0);
    assert_number(&sheet.eval("=LAMBDA(x,LAMBDA(y,x+y))(1)(2)"), 3.0);
}

#[test]
fn let_result_can_be_invoked_as_lambda() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=LET(f,LAMBDA(x,x+1),f)(2)"), 3.0);
}

#[test]
fn parenthesized_lambda_invocation_preserves_recursion() {
    let mut sheet = TestSheet::new();
    assert_number(
        &sheet.eval("=LET(FACT,LAMBDA(n,IF(n<=1,1,n*FACT(n-1))),(FACT)(5))"),
        120.0,
    );
}

#[test]
fn lambda_captures_lexical_env() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=LET(a,10,f,LAMBDA(x,a+x),f(5))"), 15.0);
}

#[test]
fn defined_name_lambda_can_be_called_like_function() {
    let mut engine = Engine::new();
    engine
        .define_name(
            "ADD1",
            NameScope::Workbook,
            NameDefinition::Formula("=LAMBDA(x,x+1)".to_string()),
        )
        .expect("define name");

    engine
        .set_cell_formula("Sheet1", "A1", "=ADD1(2)")
        .expect("set formula");
    engine.recalculate();
    assert_number(&engine.get_cell_value("Sheet1", "A1"), 3.0);
}

#[test]
fn lambda_supports_recursion_and_depth_guard() {
    let mut sheet = TestSheet::new();
    assert_number(
        &sheet.eval("=LET(FACT,LAMBDA(n,IF(n<=1,1,n*FACT(n-1))),FACT(5))"),
        120.0,
    );

    assert_eq!(
        sheet.eval("=LET(f,LAMBDA(x,f(x)),f(1))"),
        Value::Error(ErrorKind::Calc)
    );
}

#[test]
fn lambda_supports_omitted_parameters_with_isomitted() {
    let mut sheet = TestSheet::new();

    // Calling a LAMBDA with fewer arguments should bind missing parameters as blank,
    // while still allowing the body to detect omission via ISOMITTED.
    assert_number(
        &sheet.eval("=LET(f,LAMBDA(x,y,IF(ISOMITTED(y),x,x+y)),f(2))"),
        2.0,
    );
    assert_number(
        &sheet.eval("=LET(f,LAMBDA(x,y,IF(ISOMITTED(y),x,x+y)),f(2,3))"),
        5.0,
    );

    assert_eq!(
        sheet.eval("=LET(f,LAMBDA(x,y,ISOMITTED(y)),f(1))"),
        Value::Bool(true)
    );

    // A blank placeholder is not the same as an omitted argument.
    assert_eq!(
        sheet.eval("=LET(f,LAMBDA(x,y,ISOMITTED(y)),f(1,))"),
        Value::Bool(false)
    );
}

#[test]
fn lambda_calls_can_be_used_as_reference_arguments() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", 10.0);
    sheet.set("A2", 20.0);
    sheet.set("A3", 30.0);

    // Ensure lambdas can return references so reference-only functions like OFFSET accept the
    // result without treating it as an inlined array value.
    assert_number(
        &sheet.eval("=LET(f,LAMBDA(r,r),OFFSET(f(A1:A3),0,0,1,1))"),
        10.0,
    );
}
