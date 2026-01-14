use formula_engine::{lex_partial, Coord, ParseOptions, ReferenceStyle, TokenKind};

#[test]
fn lex_partial_unterminated_string_literal_returns_tokens_and_error() {
    let out = lex_partial("\"hello", &ParseOptions::default());

    let err = out.error.expect("expected error for unterminated string");
    assert_eq!(err.message, "Unterminated string literal");
    assert_eq!(err.span.start, 0);
    assert_eq!(err.span.end, 6);

    assert_eq!(out.tokens.len(), 2);
    assert_eq!(out.tokens[0].kind, TokenKind::String("hello".to_string()));
    assert_eq!(out.tokens[0].span.start, 0);
    assert_eq!(out.tokens[0].span.end, 6);
    assert_eq!(out.tokens[1].kind, TokenKind::Eof);
    assert_eq!(out.tokens[1].span.start, 6);
    assert_eq!(out.tokens[1].span.end, 6);
}

#[test]
fn lex_partial_unterminated_quoted_identifier_returns_tokens_and_error() {
    let out = lex_partial("'Sheet1", &ParseOptions::default());

    let err = out
        .error
        .expect("expected error for unterminated quoted identifier");
    assert_eq!(err.message, "Unterminated quoted identifier");
    assert_eq!(err.span.start, 0);
    assert_eq!(err.span.end, 7);

    assert_eq!(out.tokens.len(), 2);
    assert_eq!(
        out.tokens[0].kind,
        TokenKind::QuotedIdent("Sheet1".to_string())
    );
    assert_eq!(out.tokens[0].span.start, 0);
    assert_eq!(out.tokens[0].span.end, 7);
    assert_eq!(out.tokens[1].kind, TokenKind::Eof);
    assert_eq!(out.tokens[1].span.start, 7);
    assert_eq!(out.tokens[1].span.end, 7);
}

#[test]
fn lex_partial_true_false_are_ident_when_followed_by_paren() {
    let out_true = lex_partial("TRUE()", &ParseOptions::default());
    assert!(out_true.error.is_none());
    assert_eq!(
        out_true.tokens[0].kind,
        TokenKind::Ident("TRUE".to_string())
    );

    let out_false = lex_partial("FALSE()", &ParseOptions::default());
    assert!(out_false.error.is_none());
    assert_eq!(
        out_false.tokens[0].kind,
        TokenKind::Ident("FALSE".to_string())
    );

    let out_literal = lex_partial("TRUE", &ParseOptions::default());
    assert!(out_literal.error.is_none());
    assert_eq!(out_literal.tokens[0].kind, TokenKind::Boolean(true));
}

#[test]
fn lex_partial_r1c1_ref_followed_by_dot_produces_field_access_tokens() {
    let opts = ParseOptions {
        reference_style: ReferenceStyle::R1C1,
        ..ParseOptions::default()
    };
    let out = lex_partial("RC[-1].Price", &opts);
    assert!(out.error.is_none());

    assert_eq!(out.tokens.len(), 4);
    assert!(
        matches!(
            &out.tokens[0].kind,
            TokenKind::R1C1Cell(cell) if cell.row == Coord::Offset(0) && cell.col == Coord::Offset(-1)
        ),
        "unexpected first token: {:?}",
        out.tokens[0].kind
    );
    assert_eq!(out.tokens[1].kind, TokenKind::Dot);
    assert_eq!(out.tokens[2].kind, TokenKind::Ident("Price".to_string()));
    assert_eq!(out.tokens[3].kind, TokenKind::Eof);
}

#[test]
fn lex_partial_r1c1_bracket_field_access_produces_tokens() {
    let opts = ParseOptions {
        reference_style: ReferenceStyle::R1C1,
        ..ParseOptions::default()
    };
    let out = lex_partial(r#"RC[-1].["Change%"]"#, &opts);
    assert!(
        out.error.is_none(),
        "expected no lex error, got {:?}",
        out.error
    );

    assert_eq!(out.tokens.len(), 6);
    assert!(
        matches!(
            &out.tokens[0].kind,
            TokenKind::R1C1Cell(cell) if cell.row == Coord::Offset(0) && cell.col == Coord::Offset(-1)
        ),
        "unexpected first token: {:?}",
        out.tokens[0].kind
    );
    assert_eq!(out.tokens[1].kind, TokenKind::Dot);
    assert_eq!(out.tokens[2].kind, TokenKind::LBracket);
    assert_eq!(
        out.tokens[3].kind,
        TokenKind::Ident("\"Change%\"".to_string())
    );
    assert_eq!(out.tokens[4].kind, TokenKind::RBracket);
    assert_eq!(out.tokens[5].kind, TokenKind::Eof);
}
