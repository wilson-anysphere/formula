use formula_engine::{
    parse_formula, CellAddr, Coord, Expr, NameRef, ParseOptions, ReferenceStyle, SerializeOptions,
};

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
fn r1c1_roundtrip_row_and_column_ranges() {
    let mut opts = ParseOptions::default();
    opts.reference_style = ReferenceStyle::R1C1;
    let mut ser = SerializeOptions::default();
    ser.reference_style = ReferenceStyle::R1C1;
    roundtrip("=R1:R3", opts.clone(), ser.clone());
    roundtrip("=C1:C3", opts, ser);
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
fn r1c1_roundtrip_with_external_workbook_prefixes() {
    let mut opts = ParseOptions::default();
    opts.reference_style = ReferenceStyle::R1C1;
    let mut ser = SerializeOptions::default();
    ser.reference_style = ReferenceStyle::R1C1;
    roundtrip("=[Book.xlsx]Sheet1!R1C1+1", opts.clone(), ser.clone());
    roundtrip("='[Book Name.xlsx]Sheet 1'!R1C1+1", opts, ser);
}

#[test]
fn r1c1_roundtrip_quotes_sheet_names_that_conflict_with_r1c1_tokens() {
    let mut opts = ParseOptions::default();
    opts.reference_style = ReferenceStyle::R1C1;
    let mut ser = SerializeOptions::default();
    ser.reference_style = ReferenceStyle::R1C1;
    roundtrip("='R1C1'!R1C1+1", opts.clone(), ser.clone());
    roundtrip("='R'!R1C1+1", opts, ser);
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

#[test]
fn r1c1_does_not_mislex_identifiers_starting_with_rc_prefix() {
    let mut opts = ParseOptions::default();
    opts.reference_style = ReferenceStyle::R1C1;
    let ast = parse_formula("=RCAR", opts).unwrap();
    assert_eq!(
        ast.expr,
        Expr::NameRef(NameRef {
            workbook: None,
            sheet: None,
            name: "RCAR".to_string()
        })
    );
}

#[test]
fn r1c1_does_not_mislex_identifiers_starting_with_r1c1_prefix() {
    let mut opts = ParseOptions::default();
    opts.reference_style = ReferenceStyle::R1C1;
    let ast = parse_formula("=R1C1FOO", opts).unwrap();
    assert_eq!(
        ast.expr,
        Expr::NameRef(NameRef {
            workbook: None,
            sheet: None,
            name: "R1C1FOO".to_string()
        })
    );
}

#[test]
fn r1c1_does_not_mislex_function_calls_starting_with_rc_prefix() {
    let mut opts = ParseOptions::default();
    opts.reference_style = ReferenceStyle::R1C1;
    // Unknown functions are legal in the syntax; they should still parse.
    parse_formula("=RCAR(1)", opts).unwrap();
}

#[test]
fn r1c1_parses_i32_min_offsets() {
    let mut opts = ParseOptions::default();
    opts.reference_style = ReferenceStyle::R1C1;
    let ast = parse_formula("=R[-2147483648]C", opts).unwrap();

    let Expr::CellRef(cell_ref) = &ast.expr else {
        panic!("expected cell ref");
    };
    assert_eq!(cell_ref.row, Coord::Offset(i32::MIN));
    assert_eq!(cell_ref.col, Coord::Offset(0));

    let mut ser = SerializeOptions::default();
    ser.reference_style = ReferenceStyle::R1C1;
    assert_eq!(ast.to_string(ser).unwrap(), "=R[-2147483648]C");
}
