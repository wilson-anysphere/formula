use formula_biff::function_spec_from_id;
use pretty_assertions::assert_eq;

#[cfg(feature = "encode")]
use formula_biff::{decode_rgce, encode_rgce, EncodeRgceError};

#[test]
fn function_specs_match_expected_arg_ranges_for_common_vararg_functions() {
    let ztest = function_spec_from_id(324).expect("ZTEST spec");
    assert_eq!(ztest.name, "ZTEST");
    assert_eq!((ztest.min_args, ztest.max_args), (2, 3));

    let percentrank = function_spec_from_id(329).expect("PERCENTRANK spec");
    assert_eq!(percentrank.name, "PERCENTRANK");
    assert_eq!((percentrank.min_args, percentrank.max_args), (2, 3));

    let yearfrac = function_spec_from_id(451).expect("YEARFRAC spec");
    assert_eq!(yearfrac.name, "YEARFRAC");
    assert_eq!((yearfrac.min_args, yearfrac.max_args), (2, 3));

    let countifs = function_spec_from_id(481).expect("COUNTIFS spec");
    assert_eq!(countifs.name, "COUNTIFS");
    assert_eq!((countifs.min_args, countifs.max_args), (2, 254));
}

#[cfg(feature = "encode")]
#[test]
fn encode_accepts_optional_args_for_vararg_functions() {
    fn roundtrip(formula: &str) {
        let rgce = encode_rgce(formula).expect("encode");
        let decoded = decode_rgce(&rgce).expect("decode");
        assert_eq!(decoded, formula);
    }

    // Optional-arg functions.
    for formula in [
        "ZTEST(1,2)",
        "ZTEST(1,2,3)",
        "PERCENTRANK(1,2)",
        "PERCENTRANK(1,2,3)",
        "YEARFRAC(1,2)",
        "YEARFRAC(1,2,3)",
        "COUNTIFS(1,2)",
        "COUNTIFS(1,2,3,4)",
    ] {
        roundtrip(formula);
    }
}

#[cfg(feature = "encode")]
#[test]
fn encode_rejects_too_few_args_for_vararg_functions() {
    for (formula, name, min, max) in [
        ("ZTEST(1)", "ZTEST", 2, 3),
        ("PERCENTRANK(1)", "PERCENTRANK", 2, 3),
        ("YEARFRAC(1)", "YEARFRAC", 2, 3),
        ("COUNTIFS(1)", "COUNTIFS", 2, 254),
    ] {
        let err = encode_rgce(formula).expect_err("should reject invalid argc");
        match err {
            EncodeRgceError::InvalidArgCount {
                name: got_name,
                got,
                min: got_min,
                max: got_max,
            } => {
                assert_eq!(got_name, name);
                assert_eq!(got, 1);
                assert_eq!(got_min, min);
                assert_eq!(got_max, max);
            }
            other => panic!("expected InvalidArgCount for {formula}, got {other:?}"),
        }
    }
}

