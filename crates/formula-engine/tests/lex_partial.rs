use formula_engine::{lex_partial, ParseOptions, TokenKind};

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
    assert_eq!(out.tokens[0].kind, TokenKind::QuotedIdent("Sheet1".to_string()));
    assert_eq!(out.tokens[0].span.start, 0);
    assert_eq!(out.tokens[0].span.end, 7);
    assert_eq!(out.tokens[1].kind, TokenKind::Eof);
    assert_eq!(out.tokens[1].span.start, 7);
    assert_eq!(out.tokens[1].span.end, 7);
}

