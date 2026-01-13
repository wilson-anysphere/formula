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
