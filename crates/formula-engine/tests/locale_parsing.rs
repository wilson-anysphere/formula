use formula_engine::{locale, parse_formula, Expr};

#[test]
fn parses_semicolon_argument_separator_in_de_de() {
    let formula = parse_formula("=SUMME(1;2;3)", &locale::DE_DE).unwrap();
    assert_eq!(
        formula.root,
        Expr::FunctionCall {
            name: "SUM".to_string(),
            args: vec![Expr::Number(1.0), Expr::Number(2.0), Expr::Number(3.0)],
        }
    );
}

#[test]
fn parses_comma_decimal_separator_in_de_de() {
    let formula = parse_formula("=SUMME(1,5;2,5)", &locale::DE_DE).unwrap();
    assert_eq!(
        formula.root,
        Expr::FunctionCall {
            name: "SUM".to_string(),
            args: vec![Expr::Number(1.5), Expr::Number(2.5)],
        }
    );

    // Round-trip: localized display and canonical persistence.
    assert_eq!(formula.to_canonical_string(), "=SUM(1.5,2.5)");
    assert_eq!(formula.to_localized_string(&locale::DE_DE), "=SUMME(1,5;2,5)");
}

#[test]
fn parses_en_us_commas_and_dots() {
    let formula = parse_formula("=SUM(1.25,2.75)", &locale::EN_US).unwrap();
    assert_eq!(
        formula.root,
        Expr::FunctionCall {
            name: "SUM".to_string(),
            args: vec![Expr::Number(1.25), Expr::Number(2.75)],
        }
    );
    assert_eq!(formula.to_localized_string(&locale::EN_US), "=SUM(1.25,2.75)");
}

