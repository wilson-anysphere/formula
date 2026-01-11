use formula_engine::{lex, LocaleConfig, ParseOptions, TokenKind};

#[test]
fn lex_en_us_decimal_and_arg_separators() {
    let locale = LocaleConfig::en_us();
    let mut opts = ParseOptions::default();
    opts.locale = locale;
    let tokens = lex("SUM(1.23,4.56)", &opts).unwrap();

    assert!(matches!(tokens[0].kind, TokenKind::Ident(ref s) if s == "SUM"));
    assert!(matches!(tokens[1].kind, TokenKind::LParen));
    assert!(matches!(tokens[2].kind, TokenKind::Number(ref n) if n == "1.23"));
    assert!(matches!(tokens[3].kind, TokenKind::ArgSep));
    assert!(matches!(tokens[4].kind, TokenKind::Number(ref n) if n == "4.56"));
    assert!(matches!(tokens[5].kind, TokenKind::RParen));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}

#[test]
fn lex_de_de_decimal_and_arg_separators() {
    let locale = LocaleConfig::de_de();
    let mut opts = ParseOptions::default();
    opts.locale = locale;
    let tokens = lex("SUM(1,23;4,56)", &opts).unwrap();

    assert!(matches!(tokens[0].kind, TokenKind::Ident(ref s) if s == "SUM"));
    assert!(matches!(tokens[1].kind, TokenKind::LParen));
    assert!(matches!(tokens[2].kind, TokenKind::Number(ref n) if n == "1,23"));
    assert!(matches!(tokens[3].kind, TokenKind::ArgSep));
    assert!(matches!(tokens[4].kind, TokenKind::Number(ref n) if n == "4,56"));
    assert!(matches!(tokens[5].kind, TokenKind::RParen));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}

#[test]
fn lex_de_de_array_separators() {
    let locale = LocaleConfig::de_de();
    let mut opts = ParseOptions::default();
    opts.locale = locale;
    let tokens = lex("{1\\2;3\\4}", &opts).unwrap();

    assert!(matches!(tokens[0].kind, TokenKind::LBrace));
    assert!(matches!(tokens[1].kind, TokenKind::Number(ref n) if n == "1"));
    assert!(matches!(tokens[2].kind, TokenKind::ArrayColSep));
    assert!(matches!(tokens[3].kind, TokenKind::Number(ref n) if n == "2"));
    assert!(matches!(tokens[4].kind, TokenKind::ArrayRowSep));
    assert!(matches!(tokens[5].kind, TokenKind::Number(ref n) if n == "3"));
    assert!(matches!(tokens[6].kind, TokenKind::ArrayColSep));
    assert!(matches!(tokens[7].kind, TokenKind::Number(ref n) if n == "4"));
    assert!(matches!(tokens[8].kind, TokenKind::RBrace));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}
