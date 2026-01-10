use formula_engine::{parse_formula, LocaleConfig, ParseOptions, SerializeOptions};

fn roundtrip(formula: &str, opts: ParseOptions, ser: SerializeOptions) {
    let ast1 = parse_formula(formula, opts.clone()).unwrap();
    let s = ast1.to_string(ser).unwrap();
    let ast2 = parse_formula(&s, opts).unwrap();
    assert_eq!(ast1, ast2, "formula `{formula}` -> `{s}`");
}

#[test]
fn roundtrip_with_quoted_sheet_and_structured_refs() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("='My Sheet'!$A$1+Table1[Column]+[@Column]", opts, ser);
}

#[test]
fn roundtrip_with_external_ref_and_array_literal() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    roundtrip("=[Book.xlsx]Sheet1!A1+SUM({1,2;3,4})", opts, ser);
}

#[test]
fn roundtrip_with_xlfn_prefix_in_file_mode() {
    let opts = ParseOptions::default();
    let mut ser = SerializeOptions::default();
    ser.include_xlfn_prefix = true;
    roundtrip("=_xlfn.XLOOKUP(A1,B1,C1)", opts, ser);
}

#[test]
fn roundtrip_de_de_locale() {
    let mut opts = ParseOptions::default();
    opts.locale = LocaleConfig::de_de();
    let mut ser = SerializeOptions::default();
    ser.locale = LocaleConfig::de_de();
    roundtrip("=SUM(1,23;{1\\2;3\\4})", opts, ser);
}

#[test]
fn roundtrip_union_inside_function_arg_is_parenthesized() {
    let opts = ParseOptions::default();
    let ser = SerializeOptions::default();
    // Union operator inside function arguments requires parentheses to avoid being parsed
    // as multiple arguments.
    roundtrip("=SUM((A1,B1)+C1)", opts, ser);
}
