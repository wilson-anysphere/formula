use formula_biff::encode_rgce;
use formula_engine::{parse_formula, ParseOptions};
use formula_xlsb::rgce::decode_rgce;
use pretty_assertions::assert_eq;

fn ptg_str_rgce(s: &str) -> Vec<u8> {
    let utf16: Vec<u16> = s.encode_utf16().collect();
    assert!(utf16.len() <= u16::MAX as usize);

    let mut rgce = Vec::with_capacity(1 + 2 + utf16.len() * 2);
    rgce.push(0x17); // PtgStr
    rgce.extend_from_slice(&(utf16.len() as u16).to_le_bytes());
    for unit in utf16 {
        rgce.extend_from_slice(&unit.to_le_bytes());
    }
    rgce
}

#[test]
fn decode_escapes_quotes_in_string_literals() {
    // Underlying value contains a quote.
    let rgce = ptg_str_rgce("a\"b");
    let decoded = decode_rgce(&rgce).expect("decode");

    // Excel formula text escapes embedded quotes by doubling.
    assert_eq!(decoded, "\"a\"\"b\"");
}

#[test]
fn encode_decode_round_trip_string_literal_with_embedded_quotes() {
    let formula = "=\"He said \"\"hi\"\"\"";
    let rgce = encode_rgce(formula).expect("encode");
    let decoded = decode_rgce(&rgce).expect("decode");

    // Decoder returns formula text without the leading `=`.
    assert_eq!(decoded, "\"He said \"\"hi\"\"\"");
}

#[test]
fn decoded_string_literal_parses_in_formula_engine() {
    let rgce = ptg_str_rgce("He said \"hi\"");
    let decoded = decode_rgce(&rgce).expect("decode");

    parse_formula(&decoded, ParseOptions::default()).expect("formula should parse");
}
