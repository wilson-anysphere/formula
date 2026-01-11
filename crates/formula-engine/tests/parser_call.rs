use formula_engine::{parse_formula, CallExpr, Expr, LocaleConfig, ParseOptions, SerializeOptions};

fn roundtrip(formula: &str, opts: ParseOptions, ser: SerializeOptions) {
    let ast1 = parse_formula(formula, opts.clone()).unwrap();
    let s = ast1.to_string(ser).unwrap();
    let ast2 = parse_formula(&s, opts).unwrap();
    assert_eq!(ast1, ast2, "formula `{formula}` -> `{s}`");
}

#[test]
fn lambda_invocation_roundtrips() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("=LAMBDA(x,x+1)(3)", opts.clone(), ser.clone());

    // Canonical stringification should preserve the input.
    let ast = parse_formula("=LAMBDA(x,x+1)(3)", opts).unwrap();
    assert_eq!(ast.to_string(ser).unwrap(), "=LAMBDA(x,x+1)(3)");
}

#[test]
fn lambda_invocation_parses_multiple_args() {
    let opts = ParseOptions::default();
    let ast = parse_formula("=LAMBDA(x,x+1)(1,2)", opts).unwrap();

    let Expr::Call(CallExpr { callee, args }) = ast.expr else {
        panic!("expected top-level call expr, got {:?}", ast.expr);
    };
    assert_eq!(args.len(), 2);
    assert!(matches!(*callee, Expr::FunctionCall(_)));
}

#[test]
fn nested_lambda_invocations_roundtrip() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("=LAMBDA(x,LAMBDA(y,x+y))(1)(2)", opts, ser);
}

#[test]
fn invocation_uses_locale_argument_separator() {
    let mut opts = ParseOptions::default();
    opts.locale = LocaleConfig::de_de();
    let mut ser = SerializeOptions::default();
    ser.locale = LocaleConfig::de_de();

    roundtrip("=LAMBDA(x;x+1)(1;2)", opts, ser);
}
