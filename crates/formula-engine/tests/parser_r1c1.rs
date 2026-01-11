use formula_engine::{parse_formula, CellAddr, ParseOptions, ReferenceStyle, SerializeOptions};

fn roundtrip(formula: &str, opts: ParseOptions, ser: SerializeOptions) {
    let ast1 = parse_formula(formula, opts.clone()).unwrap();
    let s = ast1.to_string(ser).unwrap();
    let ast2 = parse_formula(&s, opts).unwrap();
    assert_eq!(ast1, ast2, "formula `{formula}` -> `{s}`");
}

#[test]
fn r1c1_roundtrip_absolute_references() {
    let mut opts = ParseOptions::default();
    opts.reference_style = ReferenceStyle::R1C1;
    let mut ser = SerializeOptions::default();
    ser.reference_style = ReferenceStyle::R1C1;
    roundtrip("=R1C1+R10C5", opts, ser);
}

#[test]
fn r1c1_roundtrip_relative_and_mixed_references() {
    let mut opts = ParseOptions::default();
    opts.reference_style = ReferenceStyle::R1C1;
    let mut ser = SerializeOptions::default();
    ser.reference_style = ReferenceStyle::R1C1;
    roundtrip("=RC+R[1]C[-1]+R1C[-1]+R[2]C1", opts, ser);
}

#[test]
fn r1c1_supports_sheet_prefixes_and_quoted_names() {
    let mut opts = ParseOptions::default();
    opts.reference_style = ReferenceStyle::R1C1;
    let mut ser = SerializeOptions::default();
    ser.reference_style = ReferenceStyle::R1C1;
    roundtrip("='My Sheet'!R1C1+Sheet2!R[1]C", opts, ser);
}

#[test]
fn converts_r1c1_relative_to_a1_using_origin() {
    let mut opts = ParseOptions::default();
    opts.reference_style = ReferenceStyle::R1C1;
    let ast = parse_formula("=R[-4]C[-2]", opts).unwrap();

    let mut ser = SerializeOptions::default();
    ser.origin = Some(CellAddr::new(4, 2)); // C5
    assert_eq!(ast.to_string(ser).unwrap(), "=A1");
}

#[test]
fn converts_a1_relative_to_r1c1_using_origin() {
    let ast = parse_formula("=A1", ParseOptions::default()).unwrap();

    let mut ser = SerializeOptions::default();
    ser.reference_style = ReferenceStyle::R1C1;
    ser.origin = Some(CellAddr::new(4, 2)); // C5
    assert_eq!(ast.to_string(ser).unwrap(), "=R[-4]C[-2]");
}

