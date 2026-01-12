#![cfg(feature = "encode")]

use formula_biff::{decode_rgce, encode_rgce};
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
fn rgce_roundtrip_basic_arithmetic() {
    let rgce = encode_rgce("=B1*2").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("B1*2"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_sum_range() {
    let rgce = encode_rgce("SUM(A1:A3)").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("SUM(A1:A3)"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_if_comparison() {
    let rgce = encode_rgce("IF(A1>0,1,0)").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("IF(A1>0,1,0)"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_intersection() {
    let rgce = encode_rgce("A1:B2 C1:D4").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("A1:B2 C1:D4"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_implicit_intersection_range() {
    let rgce = encode_rgce("@A1:A3").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("@A1:A3"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_spill_range() {
    let rgce = encode_rgce("A1#").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("A1#"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_structured_ref_this_row() {
    let rgce = encode_rgce("[@Col]").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("[@Col]"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_structured_ref_table_column() {
    let rgce = encode_rgce("Table1[Col]").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("Table1[Col]"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_structured_ref_headers_column() {
    let rgce = encode_rgce("Table1[[#Headers],[Col]]").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("Table1[[#Headers],[Col]]"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_structured_ref_all() {
    let rgce = encode_rgce("Table1[#All]").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("Table1[#All]"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_structured_ref_column_range() {
    let rgce = encode_rgce("Table1[[Col1]:[Col2]]").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("Table1[[Col1]:[Col2]]"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_structured_ref_explicit_implicit_intersection() {
    let rgce = encode_rgce("@Table1[Col]").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("@Table1[Col]"), normalize(&decoded));
}
