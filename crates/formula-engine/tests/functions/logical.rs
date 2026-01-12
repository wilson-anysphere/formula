use formula_engine::{eval, ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

#[test]
fn true_false_support_literal_and_zero_arg_function_forms() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.bytecode_program_count(), 0);

    assert_eq!(sheet.eval("=TRUE()"), Value::Bool(true));
    assert_eq!(sheet.eval("=FALSE()"), Value::Bool(false));
    assert_eq!(
        sheet.bytecode_program_count(),
        2,
        "expected TRUE()/FALSE() to compile to bytecode"
    );

    // Literal forms still work.
    assert_eq!(sheet.eval("=TRUE"), Value::Bool(true));
    assert_eq!(sheet.eval("=FALSE"), Value::Bool(false));
    assert_eq!(
        sheet.bytecode_program_count(),
        4,
        "expected TRUE/FALSE literal forms to compile to bytecode"
    );

    // Ensure `TRUE()` parses/executes as a function call (Excel-compatible).
    assert_number(&sheet.eval("=IF(TRUE(),1,2)"), 1.0);
    assert_eq!(
        sheet.bytecode_program_count(),
        5,
        "expected IF(TRUE(),...) to compile to bytecode"
    );
}

#[test]
fn if_selects_branch_and_defaults_false() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=IF(TRUE, 1, 2)"), 1.0);
    assert_number(&sheet.eval("=IF(FALSE, 1, 2)"), 2.0);
    assert_eq!(sheet.eval("=IF(FALSE, 1)"), Value::Bool(false));
}

#[test]
fn if_propagates_logical_test_error() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=IF(#DIV/0!, 1, 2)"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn and_or_ignore_text_and_blank_in_ranges_but_error_on_scalar_text() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", Value::Blank);
    sheet.set("A2", Value::Text("x".to_string()));
    sheet.set("A3", true);
    sheet.set("A4", 0.0);

    assert_eq!(sheet.eval("=AND(A1:A3)"), Value::Bool(true));
    assert_eq!(sheet.eval("=OR(A1:A3)"), Value::Bool(true));

    assert_eq!(sheet.eval("=AND(A1:A4)"), Value::Bool(false));
    assert_eq!(sheet.eval("=OR(A4)"), Value::Bool(false));

    assert_eq!(
        sheet.eval("=AND(\"x\", TRUE)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn and_or_reject_lambda_values_in_arrays() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=AND({LAMBDA(x,x),TRUE})"),
        Value::Error(ErrorKind::Value)
    );
    assert_eq!(
        sheet.eval("=OR({FALSE,LAMBDA(x,x)})"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn and_or_propagate_errors_even_if_result_known() {
    let mut sheet = TestSheet::new();
    sheet.set("A1", false);
    sheet.set("A2", Value::Error(ErrorKind::Div0));
    assert_eq!(sheet.eval("=AND(A1:A2)"), Value::Error(ErrorKind::Div0));

    sheet.set("A1", true);
    assert_eq!(sheet.eval("=OR(A1:A2)"), Value::Error(ErrorKind::Div0));
}

#[test]
fn not_coercions() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=NOT(TRUE)"), Value::Bool(false));
    assert_eq!(sheet.eval("=NOT(0)"), Value::Bool(true));
    assert_eq!(sheet.eval("=NOT(\"FALSE\")"), Value::Bool(true));
}

#[test]
fn iferror_and_ifna() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=IFERROR(1, 2)"), 1.0);
    assert_number(&sheet.eval("=IFERROR(1/0, 2)"), 2.0);

    assert_number(&sheet.eval("=IFNA(#N/A, 9)"), 9.0);
    assert_eq!(
        sheet.eval("=IFNA(#DIV/0!, 9)"),
        Value::Error(ErrorKind::Div0)
    );
}

#[test]
fn ifs_selects_first_true_condition_and_is_lazy() {
    let mut sheet = TestSheet::new();

    assert_number(&sheet.eval("=IFS(TRUE, 1, FALSE, 2)"), 1.0);
    assert_number(&sheet.eval("=IFS(FALSE, 1, TRUE, 2)"), 2.0);
    assert_eq!(
        sheet.eval("=IFS(FALSE, 1, FALSE, 2)"),
        Value::Error(ErrorKind::NA)
    );

    // Short-circuit: unselected value expressions are not evaluated.
    assert_number(&sheet.eval("=IFS(TRUE, 1, TRUE, 1/0)"), 1.0);
    assert_number(&sheet.eval("=IFS(FALSE, 1/0, TRUE, 2)"), 2.0);

    // Argument pairs are required.
    assert_eq!(
        sheet.eval("=IFS(TRUE, 1, FALSE)"),
        Value::Error(ErrorKind::Value)
    );
}

#[test]
fn ifs_accepts_xlfn_prefix_and_spills_arrays() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=_xlfn.IFS(TRUE, 1, FALSE, 2)"), 1.0);

    sheet.set_formula("C1", "=IFS({TRUE,FALSE}, 1, TRUE, 2)");
    sheet.recalculate();
    assert_eq!(sheet.get("C1"), Value::Number(1.0));
    assert_eq!(sheet.get("D1"), Value::Number(2.0));
}

#[test]
fn switch_selects_first_matching_case_and_is_lazy() {
    let mut sheet = TestSheet::new();

    assert_eq!(
        sheet.eval("=SWITCH(2, 1, \"a\", 2, \"b\", 3, \"c\")"),
        Value::Text("b".to_string())
    );

    assert_eq!(
        sheet.eval("=SWITCH(4, 1, \"a\", 2, \"b\")"),
        Value::Error(ErrorKind::NA)
    );
    assert_eq!(
        sheet.eval("=SWITCH(4, 1, \"a\", 2, \"b\", \"none\")"),
        Value::Text("none".to_string())
    );

    // Short-circuit: unselected case/result expressions are not evaluated.
    assert_eq!(sheet.eval("=SWITCH(1, 1, 10, 2, 1/0)"), Value::Number(10.0));
    assert_eq!(sheet.eval("=SWITCH(2, 1, 1/0, 2, 20)"), Value::Number(20.0));
    assert_eq!(
        sheet.eval("=SWITCH(1, 1, \"ok\", 1/0, \"bad\")"),
        Value::Text("ok".to_string())
    );

    // Spill over array expressions.
    sheet.set_formula("A1", "=SWITCH({1,2,3}, 1, \"a\", 2, \"b\", 3, \"c\")");
    sheet.recalculate();
    assert_eq!(sheet.get("A1"), Value::Text("a".to_string()));
    assert_eq!(sheet.get("B1"), Value::Text("b".to_string()));
    assert_eq!(sheet.get("C1"), Value::Text("c".to_string()));

    // Accepts `_xlfn.` prefix.
    assert_eq!(
        sheet.eval("=_xlfn.SWITCH(2, 1, \"a\", 2, \"b\")"),
        Value::Text("b".to_string())
    );

    // Argument pairs are required; default can be supplied as a trailing arg.
    assert_eq!(sheet.eval("=SWITCH(1, 1)"), Value::Error(ErrorKind::Value));
    assert_eq!(
        sheet.eval("=SWITCH(1, 1, \"a\", 2)"),
        Value::Text("a".to_string())
    );
    assert_number(&sheet.eval("=SWITCH(2, 1, 10, 99)"), 99.0);
}

#[test]
fn na_function_returns_na_error() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=NA()"), Value::Error(ErrorKind::NA));
}

#[test]
fn na_accepts_xlfn_prefix() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=_xlfn.NA()"), Value::Error(ErrorKind::NA));
}

#[test]
fn unknown_functions_return_name_error() {
    let mut sheet = TestSheet::new();
    assert_eq!(
        sheet.eval("=NO_SUCH_FUNCTION(1)"),
        Value::Error(ErrorKind::Name)
    );
}

#[test]
fn function_lookup_is_case_insensitive() {
    let mut sheet = TestSheet::new();
    assert_number(&sheet.eval("=sUm(1,2,3)"), 6.0);
}

#[test]
fn parser_preserves_original_function_name_for_display() {
    let expr = eval::Parser::parse("=sUm(1,2)").unwrap();
    match expr {
        eval::Expr::FunctionCall {
            name,
            original_name,
            ..
        } => {
            assert_eq!(name, "SUM");
            assert_eq!(original_name, "sUm");
        }
        other => panic!("expected function call, got {other:?}"),
    }
}

#[test]
fn parser_strips_xlfn_prefix_for_lookup_but_preserves_original() {
    let expr = eval::Parser::parse("=_xlfn.XLOOKUP(1,2,3)").unwrap();
    match expr {
        eval::Expr::FunctionCall {
            name,
            original_name,
            ..
        } => {
            assert_eq!(name, "XLOOKUP");
            assert_eq!(original_name, "_xlfn.XLOOKUP");
        }
        other => panic!("expected function call, got {other:?}"),
    }
}
