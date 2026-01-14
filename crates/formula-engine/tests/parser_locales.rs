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
fn lex_fr_fr_decimal_and_arg_separators() {
    let locale = LocaleConfig::fr_fr();
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
fn lex_fr_fr_supports_nbsp_thousands_separator_in_numbers() {
    let locale = LocaleConfig::fr_fr();
    let mut opts = ParseOptions::default();
    opts.locale = locale;
    let tokens = lex("SUM(1\u{00A0}234,56;0,5)", &opts).unwrap();

    assert!(matches!(tokens[0].kind, TokenKind::Ident(ref s) if s == "SUM"));
    assert!(matches!(tokens[1].kind, TokenKind::LParen));
    assert!(matches!(tokens[2].kind, TokenKind::Number(ref n) if n == "1234,56"));
    assert!(matches!(tokens[3].kind, TokenKind::ArgSep));
    assert!(matches!(tokens[4].kind, TokenKind::Number(ref n) if n == "0,5"));
    assert!(matches!(tokens[5].kind, TokenKind::RParen));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}

#[test]
fn lex_fr_fr_supports_narrow_nbsp_thousands_separator_in_numbers() {
    // Some French locales/spreadsheets use U+202F NARROW NO-BREAK SPACE for grouping.
    let locale = LocaleConfig::fr_fr();
    let mut opts = ParseOptions::default();
    opts.locale = locale;
    let tokens = lex("SUM(1\u{202F}234,56;0,5)", &opts).unwrap();

    assert!(matches!(tokens[0].kind, TokenKind::Ident(ref s) if s == "SUM"));
    assert!(matches!(tokens[1].kind, TokenKind::LParen));
    assert!(matches!(tokens[2].kind, TokenKind::Number(ref n) if n == "1234,56"));
    assert!(matches!(tokens[3].kind, TokenKind::ArgSep));
    assert!(matches!(tokens[4].kind, TokenKind::Number(ref n) if n == "0,5"));
    assert!(matches!(tokens[5].kind, TokenKind::RParen));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}

#[test]
fn lex_fr_fr_supports_multiple_nbsp_thousands_separators_in_numbers() {
    // Ensure we can handle more than one grouping separator in a single literal (e.g. 1 234 567).
    let locale = LocaleConfig::fr_fr();
    let mut opts = ParseOptions::default();
    opts.locale = locale;
    let tokens = lex("SUM(1\u{00A0}234\u{00A0}567,89;0,5)", &opts).unwrap();

    assert!(matches!(tokens[0].kind, TokenKind::Ident(ref s) if s == "SUM"));
    assert!(matches!(tokens[1].kind, TokenKind::LParen));
    assert!(matches!(tokens[2].kind, TokenKind::Number(ref n) if n == "1234567,89"));
    assert!(matches!(tokens[3].kind, TokenKind::ArgSep));
    assert!(matches!(tokens[4].kind, TokenKind::Number(ref n) if n == "0,5"));
    assert!(matches!(tokens[5].kind, TokenKind::RParen));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}

#[test]
fn lex_fr_fr_supports_mixed_nbsp_and_narrow_nbsp_thousands_separators_in_numbers() {
    // Spreadsheets can sometimes mix NBSP and narrow NBSP for grouping; treat them equivalently.
    let locale = LocaleConfig::fr_fr();
    let mut opts = ParseOptions::default();
    opts.locale = locale;
    let tokens = lex("SUM(1\u{00A0}234\u{202F}567,89;0,5)", &opts).unwrap();

    assert!(matches!(tokens[0].kind, TokenKind::Ident(ref s) if s == "SUM"));
    assert!(matches!(tokens[1].kind, TokenKind::LParen));
    assert!(matches!(tokens[2].kind, TokenKind::Number(ref n) if n == "1234567,89"));
    assert!(matches!(tokens[3].kind, TokenKind::ArgSep));
    assert!(matches!(tokens[4].kind, TokenKind::Number(ref n) if n == "0,5"));
    assert!(matches!(tokens[5].kind, TokenKind::RParen));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}

#[test]
fn lex_es_es_decimal_and_arg_separators() {
    let locale = LocaleConfig::es_es();
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

#[test]
fn lex_localized_error_literal_inverted_exclamation() {
    let mut opts = ParseOptions::default();
    opts.locale = LocaleConfig::en_us();

    let tokens = lex("#¡VALOR!", &opts).unwrap();
    assert!(matches!(tokens[0].kind, TokenKind::Error(ref s) if s == "#¡VALOR!"));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}

#[test]
fn lex_localized_error_literal_inverted_question() {
    let mut opts = ParseOptions::default();
    opts.locale = LocaleConfig::en_us();

    let tokens = lex("#¿NOMBRE?", &opts).unwrap();
    assert!(matches!(tokens[0].kind, TokenKind::Error(ref s) if s == "#¿NOMBRE?"));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}

#[test]
fn lex_hash_postfix_spill_range_not_error_literal() {
    let mut opts = ParseOptions::default();
    opts.locale = LocaleConfig::en_us();

    let tokens = lex("A1#", &opts).unwrap();
    assert!(matches!(tokens[0].kind, TokenKind::Cell(_)));
    assert!(matches!(tokens[1].kind, TokenKind::Hash));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}

#[test]
fn lex_localized_error_literal_with_non_ascii_letters() {
    let mut opts = ParseOptions::default();
    opts.locale = LocaleConfig::en_us();

    let tokens = lex("#ÜBERLAUF!", &opts).unwrap();
    assert!(matches!(tokens[0].kind, TokenKind::Error(ref s) if s == "#ÜBERLAUF!"));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}

#[test]
fn lex_de_de_accepts_canonical_leading_decimal() {
    let locale = LocaleConfig::de_de();
    let mut opts = ParseOptions::default();
    opts.locale = locale;
    let tokens = lex(".5", &opts).unwrap();

    assert!(matches!(tokens[0].kind, TokenKind::Number(ref n) if n == ".5"));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}

#[test]
fn lex_fr_fr_accepts_canonical_leading_decimal() {
    let locale = LocaleConfig::fr_fr();
    let mut opts = ParseOptions::default();
    opts.locale = locale;
    let tokens = lex(".5", &opts).unwrap();

    assert!(matches!(tokens[0].kind, TokenKind::Number(ref n) if n == ".5"));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}

#[test]
fn lex_es_es_accepts_canonical_leading_decimal() {
    let locale = LocaleConfig::es_es();
    let mut opts = ParseOptions::default();
    opts.locale = locale;
    let tokens = lex(".5", &opts).unwrap();

    assert!(matches!(tokens[0].kind, TokenKind::Number(ref n) if n == ".5"));
    assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
}

#[test]
fn lex_comma_decimal_locales_accept_canonical_decimal_separator_in_numbers() {
    // Users can copy/paste canonical formulas into comma-decimal locales (de-DE/fr-FR/es-ES).
    // Ensure the lexer accepts `.` as a decimal separator (as long as it isn't a thousands-group
    // pattern like `1.234` in `de-DE`).
    for locale in [
        LocaleConfig::de_de(),
        LocaleConfig::fr_fr(),
        LocaleConfig::es_es(),
    ] {
        let mut opts = ParseOptions::default();
        opts.locale = locale;
        let tokens = lex("SUM(1.23;4.56)", &opts).unwrap();

        assert!(matches!(tokens[0].kind, TokenKind::Ident(ref s) if s == "SUM"));
        assert!(matches!(tokens[1].kind, TokenKind::LParen));
        assert!(matches!(tokens[2].kind, TokenKind::Number(ref n) if n == "1.23"));
        assert!(matches!(tokens[3].kind, TokenKind::ArgSep));
        assert!(matches!(tokens[4].kind, TokenKind::Number(ref n) if n == "4.56"));
        assert!(matches!(tokens[5].kind, TokenKind::RParen));
        assert!(matches!(tokens.last().unwrap().kind, TokenKind::Eof));
    }
}
