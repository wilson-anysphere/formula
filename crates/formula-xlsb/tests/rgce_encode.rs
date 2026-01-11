use formula_xlsb::rgce::decode_rgce;
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
fn rgce_roundtrip_b1_times_2() {
    let rgce = formula_biff::encode_rgce("=B1*2").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("B1*2"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_sum_range() {
    let rgce = formula_biff::encode_rgce("SUM(A1:A3)").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("SUM(A1:A3)"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_if_comparison() {
    let rgce = formula_biff::encode_rgce("IF(A1>0,1,0)").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("IF(A1>0,1,0)"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_intersection() {
    let rgce = formula_biff::encode_rgce("A1:B2 C1:D4").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("A1:B2 C1:D4"), normalize(&decoded));
}

