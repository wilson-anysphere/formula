use formula_engine::{lex, parse_formula, Expr, LocaleConfig, ParseOptions, TokenKind};

#[test]
fn lex_div0_error_token() {
    let tokens = lex("#DIV/0!", &LocaleConfig::en_us()).unwrap();
    assert!(matches!(
        tokens[0].kind,
        TokenKind::Error(ref s) if s == "#DIV/0!"
    ));
    assert!(matches!(tokens[1].kind, TokenKind::Eof));
}

#[test]
fn lex_na_error_token() {
    let tokens = lex("#N/A", &LocaleConfig::en_us()).unwrap();
    assert!(matches!(
        tokens[0].kind,
        TokenKind::Error(ref s) if s == "#N/A"
    ));
    assert!(matches!(tokens[1].kind, TokenKind::Eof));
}

#[test]
fn parse_error_literal_as_expression() {
    let ast = parse_formula("=#REF!", ParseOptions::default()).unwrap();
    assert!(matches!(ast.expr, Expr::Error(ref s) if s == "#REF!"));
}
