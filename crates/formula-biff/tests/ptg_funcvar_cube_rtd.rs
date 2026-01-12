#![cfg(feature = "encode")]

use formula_biff::{decode_rgce, encode_rgce, EncodeRgceError};
use pretty_assertions::assert_eq;

fn roundtrip(formula: &str) {
    let rgce = encode_rgce(formula).expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(decoded, formula);
}

#[test]
fn rgce_roundtrip_cube_and_rtd_functions() {
    // Exercise `PtgFuncVar` encoding/decoding for external data functions that have explicit
    // argument-count metadata in `function_ids.rs`.
    for formula in [
        r#"RTD("prog","server","topic")"#,
        r#"CUBEVALUE("conn","[Measures].[Sales]")"#,
        r#"CUBEMEMBER("conn","[Dim].[All]","Caption")"#,
        r#"CUBEMEMBERPROPERTY("conn","[Dim].[All]","PROP")"#,
        r#"CUBERANKEDMEMBER("conn","[SetExpr]",1,"Caption")"#,
        r#"CUBEKPIMEMBER("conn","KPI","KPIValue")"#,
        r#"CUBESET("conn","[SetExpr]","Caption",1,"[Measures].[Sales]")"#,
        r#"CUBESETCOUNT("[SetHandle]")"#,
    ] {
        roundtrip(formula);
    }
}

#[test]
fn encode_rejects_invalid_cube_and_rtd_argument_counts() {
    let err = encode_rgce(r#"RTD("a","b")"#).expect_err("RTD should reject argc < min");
    match err {
        EncodeRgceError::InvalidArgCount { name, got, min, max } => {
            assert_eq!(name, "RTD");
            assert_eq!(got, 2);
            assert_eq!(min, 3);
            assert_eq!(max, 255);
        }
        other => panic!("expected InvalidArgCount, got {other:?}"),
    }

    let err =
        encode_rgce("CUBESETCOUNT()").expect_err("CUBESETCOUNT should reject argc < min");
    match err {
        EncodeRgceError::InvalidArgCount { name, got, min, max } => {
            assert_eq!(name, "CUBESETCOUNT");
            assert_eq!(got, 0);
            assert_eq!(min, 1);
            assert_eq!(max, 1);
        }
        other => panic!("expected InvalidArgCount, got {other:?}"),
    }
}

