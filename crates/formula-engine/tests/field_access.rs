use formula_engine::{parse_formula, parse_formula_partial, ParseOptions, SerializeOptions};

#[test]
fn field_access_roundtrip_ident() {
    let formula = "=A1.Price";
    let ast = parse_formula(formula, ParseOptions::default()).expect("parse");
    let out = ast
        .to_string(SerializeOptions::default())
        .expect("serialize");
    assert_eq!(out, formula);
}

#[test]
fn field_access_roundtrip_bracketed() {
    let formula = r#"=A1.["Change%"]"#;
    let ast = parse_formula(formula, ParseOptions::default()).expect("parse");
    let out = ast
        .to_string(SerializeOptions::default())
        .expect("serialize");
    assert_eq!(out, formula);
}

#[test]
fn field_access_roundtrip_nested() {
    let formula = "=A1.Address.City";
    let ast = parse_formula(formula, ParseOptions::default()).expect("parse");
    let out = ast
        .to_string(SerializeOptions::default())
        .expect("serialize");
    assert_eq!(out, formula);
}

#[test]
fn field_access_partial_parse_trailing_dot() {
    let partial = parse_formula_partial("=A1.", ParseOptions::default());
    assert!(
        partial.error.is_some(),
        "expected partial parse to capture an error"
    );
}

