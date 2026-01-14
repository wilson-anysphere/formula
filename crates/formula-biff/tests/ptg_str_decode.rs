use formula_biff::decode_rgce;
use pretty_assertions::assert_eq;

fn ptg_str(units: &[u16]) -> Vec<u8> {
    let mut out = Vec::with_capacity(1 + 2 + units.len() * 2);
    out.push(0x17); // PtgStr
    out.extend_from_slice(&(units.len() as u16).to_le_bytes());
    for u in units {
        out.extend_from_slice(&u.to_le_bytes());
    }
    out
}

#[test]
fn ptg_str_decodes_lossy_on_invalid_utf16() {
    // A standalone unpaired surrogate is malformed UTF-16.
    let rgce = ptg_str(&[0xD800]);
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(decoded, format!("\"{}\"", char::REPLACEMENT_CHARACTER));
}

#[test]
fn ptg_str_decodes_well_formed_utf16_identically() {
    let units: Vec<u16> = "Hello".encode_utf16().collect();
    let rgce = ptg_str(&units);
    let decoded = decode_rgce(&rgce).expect("decode");
    assert_eq!(decoded, "\"Hello\"");
}
