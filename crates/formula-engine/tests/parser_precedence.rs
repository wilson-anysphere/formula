use formula_engine::{parse_formula, BinaryExpr, BinaryOp, Expr, ParseOptions};

fn num(s: &str) -> Expr {
    Expr::Number(s.to_string())
}

fn bin(op: BinaryOp, left: Expr, right: Expr) -> Expr {
    Expr::Binary(BinaryExpr {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

#[test]
fn parses_arithmetic_precedence() {
    let ast = parse_formula("=1+2*3^2", ParseOptions::default()).unwrap();

    let expected = bin(
        BinaryOp::Add,
        num("1"),
        bin(
            BinaryOp::Mul,
            num("2"),
            bin(BinaryOp::Pow, num("3"), num("2")),
        ),
    );

    assert_eq!(ast.expr, expected);
}

#[test]
fn parses_reference_operator_precedence() {
    let ast = parse_formula("=A1:B2 C1:D4", ParseOptions::default()).unwrap();

    let a1 = ast.expr.clone();
    // We mainly assert shape: Intersect(Range(..), Range(..))
    let Expr::Binary(BinaryExpr { op, left, right }) = a1 else {
        panic!("expected a binary expression, got {:?}", a1);
    };
    assert_eq!(op, BinaryOp::Intersect);
    assert!(matches!(
        left.as_ref(),
        Expr::Binary(BinaryExpr {
            op: BinaryOp::Range,
            ..
        })
    ));
    assert!(matches!(
        right.as_ref(),
        Expr::Binary(BinaryExpr {
            op: BinaryOp::Range,
            ..
        })
    ));
}

#[test]
fn percent_binds_tighter_than_exponent() {
    let ast = parse_formula("=2^3%", ParseOptions::default()).unwrap();

    let Expr::Binary(BinaryExpr { op, right, .. }) = ast.expr else {
        panic!("expected binary expr");
    };
    assert_eq!(op, BinaryOp::Pow);
    assert!(matches!(right.as_ref(), Expr::Postfix(_)));
}

#[test]
fn concat_binds_looser_than_addition() {
    let ast = parse_formula("=1+2&3", ParseOptions::default()).unwrap();

    let expected = bin(
        BinaryOp::Concat,
        bin(BinaryOp::Add, num("1"), num("2")),
        num("3"),
    );
    assert_eq!(ast.expr, expected);
}

#[test]
fn comparison_binds_looser_than_concat() {
    let ast = parse_formula("=1&2=12", ParseOptions::default()).unwrap();

    let expected = bin(
        BinaryOp::Eq,
        bin(BinaryOp::Concat, num("1"), num("2")),
        num("12"),
    );
    assert_eq!(ast.expr, expected);
}
