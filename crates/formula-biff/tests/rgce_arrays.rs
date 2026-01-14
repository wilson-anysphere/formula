#![cfg(feature = "encode")]

use formula_biff::{decode_rgce_with_rgcb, encode_rgce_with_rgcb};
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
fn rgce_roundtrip_array_literal() {
    let encoded = encode_rgce_with_rgcb("={1,2;3,4}").expect("encode");
    assert!(!encoded.rgcb.is_empty(), "rgcb should be non-empty for PtgArray");
    assert!(
        encoded.rgce.iter().any(|&b| b == 0x20),
        "rgce should contain PtgArray (0x20)"
    );
    let decoded = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(normalize("={1,2;3,4}"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_sum_over_array_literal() {
    let encoded = encode_rgce_with_rgcb("=SUM({4,5})").expect("encode");
    assert!(!encoded.rgcb.is_empty(), "rgcb should be non-empty for PtgArray");
    assert!(
        encoded.rgce.iter().any(|&b| b == 0x20),
        "rgce should contain PtgArray (0x20)"
    );
    let decoded = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(normalize("SUM({4,5})"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_array_literal_mixed_types_and_blanks() {
    // 2x3 array with mixed literal types and blanks.
    let encoded =
        encode_rgce_with_rgcb("={1,,\"hi\";TRUE,#DIV/0!,FALSE}").expect("encode");
    assert!(!encoded.rgcb.is_empty(), "rgcb should be non-empty for PtgArray");
    assert!(
        encoded.rgce.iter().any(|&b| b == 0x20),
        "rgce should contain PtgArray (0x20)"
    );
    let decoded = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(normalize("{1,,\"hi\";TRUE,#DIV/0!,FALSE}"), normalize(&decoded));
}

#[test]
fn rgce_encodes_unary_plus_and_minus_in_array_literals() {
    // Unary `+` is valid syntax but is not preserved by the BIFF encoding, which stores the
    // literal numeric value. We assert semantic roundtrip by comparing against the canonical text
    // form without unary plus.
    let encoded = encode_rgce_with_rgcb("={+1,-2}").expect("encode");
    assert!(!encoded.rgcb.is_empty(), "rgcb should be non-empty for PtgArray");
    assert!(
        encoded.rgce.iter().any(|&b| b == 0x20),
        "rgce should contain PtgArray (0x20)"
    );
    let decoded = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(normalize("{1,-2}"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_array_literal_string_with_quotes() {
    let encoded = encode_rgce_with_rgcb("={\"a\"\"b\"}").expect("encode");
    assert!(!encoded.rgcb.is_empty(), "rgcb should be non-empty for PtgArray");
    assert!(
        encoded.rgce.iter().any(|&b| b == 0x20),
        "rgce should contain PtgArray (0x20)"
    );
    let decoded = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(normalize("{\"a\"\"b\"}"), normalize(&decoded));
}
