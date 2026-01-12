#![cfg(feature = "encode")]

use formula_biff::{decode_rgce, encode_rgce, EncodeRgceError};
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
fn rgce_roundtrip_if_missing_arg() {
    // Excel encodes blank function arguments as `PtgMissArg` (0x16). Missing args can appear at
    // any position, not just as a trailing optional argument.
    let rgce = encode_rgce("IF(,1,0)").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("IF(,1,0)"), normalize(&decoded));
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
fn rgce_encode_structured_ref_is_unsupported() {
    for formula in [
        "Table1[Col]",
        "[@Col]",
        "@Table1[Col]",
        "Table1[Col]#",
        "Table1[Col]%",
        "(Table1[Col])",
        "Table1[Col]+1",
        "SUM(Table1[Col])",
        "Table1[#All]",
        "Table1[[#Headers],[Col]]",
    ] {
        match encode_rgce(formula) {
            Err(EncodeRgceError::Unsupported(msg)) => {
                assert!(msg.contains("table-id"), "unexpected message: {msg}");
            }
            other => panic!("expected Unsupported error, got: {other:?} (formula={formula})"),
        }
    }
}

#[test]
fn rgce_encode_field_access_is_unsupported() {
    match encode_rgce("A1.Price") {
        Err(EncodeRgceError::Unsupported(_)) => {}
        other => panic!("expected Unsupported error, got: {other:?}"),
    }
}

#[test]
fn rgce_roundtrip_discount_securities_and_tbill_functions() {
    // Ensure the BIFF encoder/decoder roundtrips for the discount security + T-bill functions.
    // This mainly exercises `PtgFuncVar` for optional `basis` arguments and `PtgFunc` for fixed
    // arity T-bill functions.
    for formula in [
        "DISC(DATE(2020,1,1),DATE(2021,1,1),97,100)",
        "DISC(DATE(2020,1,1),DATE(2021,1,1),97,100,1)",
        "DISC(DATE(2020,1,1),DATE(2021,1,1),97,100,)",
        "PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),0.05,100)",
        "PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),0.05,100,2)",
        "PRICEDISC(DATE(2020,1,1),DATE(2021,1,1),0.05,100,)",
        "YIELDDISC(DATE(2020,1,1),DATE(2021,1,1),97,100)",
        "YIELDDISC(DATE(2020,1,1),DATE(2021,1,1),97,100,3)",
        "YIELDDISC(DATE(2020,1,1),DATE(2021,1,1),97,100,)",
        "INTRATE(DATE(2020,1,1),DATE(2021,1,1),97,100)",
        "INTRATE(DATE(2020,1,1),DATE(2021,1,1),97,100,0)",
        "INTRATE(DATE(2020,1,1),DATE(2021,1,1),97,100,)",
        "RECEIVED(DATE(2020,1,1),DATE(2021,1,1),95,0.05)",
        "RECEIVED(DATE(2020,1,1),DATE(2021,1,1),95,0.05,0)",
        "RECEIVED(DATE(2020,1,1),DATE(2021,1,1),95,0.05,)",
        "PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04)",
        "PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,0)",
        "PRICEMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,0.04,)",
        "YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,100.76923076923077)",
        "YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,100.76923076923077,0)",
        "YIELDMAT(DATE(2020,1,1),DATE(2021,1,1),DATE(2019,1,1),0.05,100.76923076923077,)",
        "TBILLPRICE(DATE(2020,1,1),DATE(2020,7,1),0.05)",
        "TBILLYIELD(DATE(2020,1,1),DATE(2020,7,1),97.47222222222223)",
        "TBILLEQ(DATE(2020,1,1),DATE(2020,12,31),0.05)",
    ] {
        let rgce = encode_rgce(formula).expect("encode");
        let decoded = decode_rgce(&rgce).expect("decode");
        assert_eq!(normalize(formula), normalize(&decoded));
    }
}

#[test]
fn rgce_roundtrip_modern_error_literals() {
    for (code, lit) in [
        (0x2C, "#SPILL!"),
        (0x2D, "#CALC!"),
        (0x2E, "#FIELD!"),
        (0x2F, "#CONNECT!"),
        (0x30, "#BLOCKED!"),
        (0x31, "#UNKNOWN!"),
    ] {
        let rgce = encode_rgce(lit).expect("encode");
        assert_eq!(rgce, vec![0x1C, code], "encode {lit}");

        let decoded = decode_rgce(&rgce).expect("decode");
        assert_eq!(decoded, lit, "decode code={code:#04x}");
    }
}
