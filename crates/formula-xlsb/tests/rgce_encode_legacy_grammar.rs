use formula_xlsb::rgce::{decode_rgce_with_context, encode_rgce_with_context, CellCoord};
use formula_xlsb::workbook_context::WorkbookContext;
use pretty_assertions::assert_eq;

fn normalize(formula: &str) -> String {
    let ast = formula_engine::parse_formula(formula, formula_engine::ParseOptions::default())
        .expect("parse formula");
    ast.to_string(formula_engine::SerializeOptions {
        omit_equals: true,
        ..Default::default()
    })
    .expect("serialize formula")
}

#[test]
fn legacy_encoder_roundtrips_concat_strings() {
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("=\"a\"&\"b\"", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("\"a\"&\"b\""), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_if_with_comparison_and_strings() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=IF(A1>0,\"pos\",\"neg\")", &ctx, CellCoord::new(0, 0))
        .expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("IF(A1>0,\"pos\",\"neg\")"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_le_comparison() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=A1<=B1", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("A1<=B1"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_intersection_operator() {
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("=A1:B2 C1:D4", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("A1:B2 C1:D4"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_sum_range_still_works() {
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("=SUM(A1:A3)", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("SUM(A1:A3)"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_boolean_literals() {
    let ctx = WorkbookContext::default();
    for (raw, expected) in [("=TRUE", "TRUE"), ("=FALSE", "FALSE")] {
        let encoded = encode_rgce_with_context(raw, &ctx, CellCoord::new(0, 0)).expect("encode");
        let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
        assert_eq!(normalize(expected), normalize(&decoded));
    }
}

#[test]
fn legacy_encoder_roundtrips_pow_right_associative() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=2^3^2", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("2^3^2"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_unary_minus_with_pow_precedence() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=-2^2", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("-2^2"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_implicit_intersection_on_area() {
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("=@A1:A10", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("@A1:A10"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_implicit_intersection_on_name() {
    let mut ctx = WorkbookContext::default();
    ctx.add_workbook_name("MyNamedRange", 1);

    let encoded =
        encode_rgce_with_context("=@MyNamedRange", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("@MyNamedRange"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_column_range_reference() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=SUM(A:A)", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("SUM(A:A)"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_row_range_reference() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=SUM(1:1)", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("SUM(1:1)"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_percent_operator() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=10%", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("10%"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_pow_with_percent_rhs() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=2^2%", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("2^2%"), normalize(&decoded));
}

#[test]
fn legacy_encoder_parses_scientific_notation_numbers() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=1E3+1", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("1000+1"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_union_inside_function_args() {
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("=SUM((A1,B1))", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("SUM((A1,B1))"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_ne_comparison() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=A1<>B1", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("A1<>B1"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_parenthesized_grouping() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=(1+2)*3", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("(1+2)*3"), normalize(&decoded));
}

#[test]
fn legacy_encoder_roundtrips_string_literal_quote_escaping() {
    let ctx = WorkbookContext::default();
    let encoded = encode_rgce_with_context("=\"a\"\"b\"", &ctx, CellCoord::new(0, 0)).expect("encode");
    let decoded = decode_rgce_with_context(&encoded.rgce, &ctx).expect("decode");
    assert_eq!(normalize("\"a\"\"b\""), normalize(&decoded));
}
