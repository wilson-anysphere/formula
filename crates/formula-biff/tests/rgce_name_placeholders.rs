use formula_biff::decode_rgce;
use pretty_assertions::assert_eq;

fn ptg_name(name_id: u32, ptg: u8) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(ptg);
    out.extend_from_slice(&name_id.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // reserved
    out
}

fn ptg_namex(ixti: u16, name_index: u16, ptg: u8) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(ptg);
    out.extend_from_slice(&ixti.to_le_bytes());
    out.extend_from_slice(&name_index.to_le_bytes());
    out
}

fn assert_parseable(formula: &str) {
    formula_engine::parse_formula(formula, formula_engine::ParseOptions::default())
        .expect("parse formula");
}

#[test]
fn decodes_ptg_name_to_safe_placeholder() {
    let rgce = ptg_name(123, 0x23);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Name_123");
    assert_parseable(&text);
}

#[test]
fn decodes_ptg_namex_to_safe_placeholder() {
    let rgce = ptg_namex(0, 1, 0x39);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "ExternName_IXTI0_N1");
    assert_parseable(&text);
}

#[test]
fn decodes_value_class_ptg_name_with_implicit_intersection_prefix() {
    let rgce = ptg_name(123, 0x43);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@Name_123");
    assert_parseable(&text);
}

#[test]
fn decodes_value_class_ptg_namex_with_implicit_intersection_prefix() {
    let rgce = ptg_namex(0, 1, 0x59);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@ExternName_IXTI0_N1");
    assert_parseable(&text);
}

