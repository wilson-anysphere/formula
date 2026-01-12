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
fn rgce_roundtrip_if_missing_arg() {
    let rgce = formula_biff::encode_rgce("IF(,1,0)").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("IF(,1,0)"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_intersection() {
    let rgce = formula_biff::encode_rgce("A1:B2 C1:D4").expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(normalize("A1:B2 C1:D4"), normalize(&decoded));
}

#[test]
fn rgce_roundtrip_discount_securities_and_tbill_functions() {
    // Keep parity with `formula-biff`'s roundtrip coverage for these functions. This exercises the
    // XLSB decoder's `PtgFunc`/`PtgFuncVar` handling for both fixed and optional arity functions.
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
        let rgce = formula_biff::encode_rgce(formula).expect("encode");
        let decoded = decode_rgce(&rgce).expect("decode");
        assert_eq!(normalize(formula), normalize(&decoded));
    }
}
