use formula_engine::{eval, ErrorKind, Value};

use super::harness::{assert_number, TestSheet};

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

    assert_eq!(sheet.eval("=AND(\"x\", TRUE)"), Value::Error(ErrorKind::Value));
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
fn na_function_returns_na_error() {
    let mut sheet = TestSheet::new();
    assert_eq!(sheet.eval("=NA()"), Value::Error(ErrorKind::NA));
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
