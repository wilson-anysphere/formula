use formula_biff::{decode_rgce, decode_rgce_with_rgcb, DecodeRgceError};
use pretty_assertions::assert_eq;

fn rgce_ptg_array() -> Vec<u8> {
    // BIFF12 PtgArray token + 7 unused bytes.
    vec![0x20, 0, 0, 0, 0, 0, 0, 0]
}

#[test]
fn decode_ptg_array_single_row_numbers() {
    let rgce = rgce_ptg_array();

    // Array constant: 1 row, 2 cols -> {4,5}
    let mut rgcb = Vec::new();
    rgcb.extend_from_slice(&1u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&4f64.to_le_bytes());
    rgcb.push(0x01);
    rgcb.extend_from_slice(&5f64.to_le_bytes());

    let decoded = decode_rgce_with_rgcb(&rgce, &rgcb).expect("decode");
    assert_eq!(decoded, "{4,5}");
}

#[test]
fn decode_ptg_array_mixed_types() {
    let rgce = rgce_ptg_array();

    // Array constant: 1 row, 3 cols -> {1,"hi",TRUE}
    let mut rgcb = Vec::new();
    rgcb.extend_from_slice(&2u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1

    // 1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&1f64.to_le_bytes());

    // "hi"
    rgcb.push(0x02);
    rgcb.extend_from_slice(&2u16.to_le_bytes()); // cch
    rgcb.extend_from_slice(&('h' as u16).to_le_bytes());
    rgcb.extend_from_slice(&('i' as u16).to_le_bytes());

    // TRUE
    rgcb.push(0x04);
    rgcb.push(1);

    let decoded = decode_rgce_with_rgcb(&rgce, &rgcb).expect("decode");
    assert_eq!(decoded, "{1,\"hi\",TRUE}");
}

#[test]
fn decode_ptg_array_unknown_error_code_is_best_effort() {
    let rgce = rgce_ptg_array();

    // Array constant: 1 row, 1 col -> {#UNKNOWN!}
    let mut rgcb = Vec::new();
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x10); // error
    rgcb.push(0xFF); // unknown/extended error code

    let decoded = decode_rgce_with_rgcb(&rgce, &rgcb).expect("decode");
    assert_eq!(decoded, "{#UNKNOWN!}");
}

#[test]
fn decode_ptg_array_without_rgcb_is_unsupported() {
    let rgce = rgce_ptg_array();
    match decode_rgce(&rgce) {
        Err(DecodeRgceError::UnsupportedToken { ptg: 0x20 }) => {}
        other => panic!("expected UnsupportedToken(0x20), got {other:?}"),
    }
}

