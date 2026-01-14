use formula_engine::bytecode::ast::Function;
use formula_engine::bytecode::{parse_formula, CellCoord, ErrorKind, Expr, ParseError, Value};

#[test]
fn parses_excel_error_literals() {
    let origin = CellCoord::new(0, 0);

    let cases: [(&str, ErrorKind); 14] = [
        ("#NULL!", ErrorKind::Null),
        ("#DIV/0!", ErrorKind::Div0),
        ("#VALUE!", ErrorKind::Value),
        ("#REF!", ErrorKind::Ref),
        ("#NAME?", ErrorKind::Name),
        ("#NUM!", ErrorKind::Num),
        ("#N/A", ErrorKind::NA),
        ("#GETTING_DATA", ErrorKind::GettingData),
        ("#SPILL!", ErrorKind::Spill),
        ("#CALC!", ErrorKind::Calc),
        ("#FIELD!", ErrorKind::Field),
        ("#CONNECT!", ErrorKind::Connect),
        ("#BLOCKED!", ErrorKind::Blocked),
        ("#UNKNOWN!", ErrorKind::Unknown),
    ];

    for (lit, expected) in cases {
        let formula = format!("={lit}");
        let expr = parse_formula(&formula, origin).unwrap();
        assert_eq!(
            expr,
            Expr::Literal(Value::Error(expected)),
            "failed to parse literal {lit}"
        );
    }
}

#[test]
fn parses_error_literal_in_function_call() {
    let origin = CellCoord::new(0, 0);
    let expr = parse_formula("=IF(1,#DIV/0!,0)", origin).unwrap();

    assert_eq!(
        expr,
        Expr::FuncCall {
            func: Function::If,
            args: vec![
                Expr::Literal(Value::Number(1.0)),
                Expr::Literal(Value::Error(ErrorKind::Div0)),
                Expr::Literal(Value::Number(0.0)),
            ],
        }
    );
}

#[test]
fn rejects_unknown_error_literals() {
    let origin = CellCoord::new(0, 0);
    assert_eq!(
        parse_formula("=#NOT_A_REAL_ERROR!", origin),
        Err(ParseError::UnexpectedToken(1))
    );
}

#[test]
fn parses_error_literals_case_insensitively() {
    let origin = CellCoord::new(0, 0);

    assert_eq!(
        parse_formula("=#dIv/0!", origin).unwrap(),
        Expr::Literal(Value::Error(ErrorKind::Div0))
    );
    assert_eq!(
        parse_formula("=#gEtTiNg_dAtA", origin).unwrap(),
        Expr::Literal(Value::Error(ErrorKind::GettingData))
    );
    assert_eq!(
        parse_formula("=#sPiLl!", origin).unwrap(),
        Expr::Literal(Value::Error(ErrorKind::Spill))
    );
}

#[test]
fn parses_error_literals_without_equals_and_with_whitespace() {
    let origin = CellCoord::new(0, 0);

    assert_eq!(
        parse_formula("#DIV/0!", origin).unwrap(),
        Expr::Literal(Value::Error(ErrorKind::Div0))
    );
    assert_eq!(
        parse_formula("  #REF!  ", origin).unwrap(),
        Expr::Literal(Value::Error(ErrorKind::Ref))
    );
    assert_eq!(
        parse_formula("=  #REF!  ", origin).unwrap(),
        Expr::Literal(Value::Error(ErrorKind::Ref))
    );
}

#[test]
fn parses_na_error_literal_with_bang_alias() {
    // Excel surfaces `#N/A!` in some contexts; treat it as `#N/A` for parity with the main parser.
    let origin = CellCoord::new(0, 0);
    assert_eq!(
        parse_formula("=#N/A!", origin).unwrap(),
        Expr::Literal(Value::Error(ErrorKind::NA))
    );
}
