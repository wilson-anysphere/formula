use formula_engine::{
    parse_formula, Expr, ParseOptions, PostfixExpr, PostfixOp, SerializeOptions,
};

#[test]
fn parses_spill_operator_on_cell_ref() {
    let ast = parse_formula("=A1#", ParseOptions::default()).unwrap();
    match ast.expr {
        Expr::Postfix(PostfixExpr { op: PostfixOp::SpillRange, expr }) => {
            assert!(matches!(*expr, Expr::CellRef(_)));
        }
        other => panic!("expected postfix spill expression, got {other:?}"),
    }
}

#[test]
fn parses_error_literal_starting_with_hash() {
    let ast = parse_formula("=#REF!", ParseOptions::default()).unwrap();
    assert!(matches!(ast.expr, Expr::Error(ref e) if e.eq_ignore_ascii_case("#REF!")));
}

#[test]
fn roundtrip_preserves_spill_operator() {
    let opts = ParseOptions::default();
    let ast1 = parse_formula("=A1#", opts.clone()).unwrap();
    let s = ast1.to_string(SerializeOptions::default()).unwrap();
    assert_eq!(s, "=A1#");
    let ast2 = parse_formula(&s, opts).unwrap();
    assert_eq!(ast1, ast2);
}

#[test]
fn parses_structured_ref_with_trailing_spill_operator() {
    let ast = parse_formula("=Table1[[#All],[Col]]#", ParseOptions::default()).unwrap();
    match ast.expr {
        Expr::Postfix(PostfixExpr { op: PostfixOp::SpillRange, expr }) => {
            assert!(matches!(*expr, Expr::StructuredRef(_)));
        }
        other => panic!("expected postfix spill expression, got {other:?}"),
    }
}
