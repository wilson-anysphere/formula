use formula_biff::decode_rgce;
use pretty_assertions::assert_eq;

#[test]
fn decodes_unknown_ptgerr_codes_as_unknown_literal() {
    // PtgErr (0x1C) followed by an unknown/extended error code should decode as `#UNKNOWN!`
    // instead of failing the whole rgce decode.
    let decoded = decode_rgce(&[0x1C, 0xFF]).expect("decode");
    assert_eq!(decoded, "#UNKNOWN!");

    // Optional: ensure the resulting string is parseable by `formula-engine`.
    formula_engine::parse_formula(&decoded, formula_engine::ParseOptions::default())
        .expect("parse");
}

