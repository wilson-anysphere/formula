use formula_biff::{decode_rgce, decode_rgce_with_rgcb, DecodeRgceError};
use pretty_assertions::assert_eq;

fn rgce_ptg_array() -> Vec<u8> {
    // BIFF12 PtgArray token + 7 unused bytes.
    vec![0x20, 0, 0, 0, 0, 0, 0, 0]
}

fn rgce_memfunc_with_array_subexpr() -> Vec<u8> {
    // PtgMemFunc: [ptg=0x29][cce: u16][subexpression bytes...]
    //
    // The subexpression is not printed, but it can contain PtgArray tokens that still consume
    // `rgcb` blocks. Ensure we advance the rgcb cursor through the mem payload.
    let subexpr = rgce_ptg_array();
    let cce: u16 = subexpr.len().try_into().expect("subexpression length fits u16");

    let mut rgce = vec![0x29];
    rgce.extend_from_slice(&cce.to_le_bytes());
    rgce.extend_from_slice(&subexpr);
    // Visible PtgArray follows.
    rgce.extend_from_slice(&rgce_ptg_array());
    rgce
}

fn rgce_memfunc_with_name_and_array_subexpr() -> Vec<u8> {
    // PtgMemFunc: [ptg=0x29][cce: u16][subexpression bytes...]
    //
    // Some real-world files include additional tokens (like `PtgName`) inside the non-printing
    // subexpression before `PtgArray`. Ensure we skip the full `PtgName` payload so we can still
    // find and consume the nested `PtgArray`'s `rgcb` block.
    //
    // PtgName: [ptg=0x23][nameId: u32][reserved: u16]
    let mut subexpr = vec![0x23];
    subexpr.extend_from_slice(&123u32.to_le_bytes());
    subexpr.extend_from_slice(&0u16.to_le_bytes()); // reserved
    subexpr.extend_from_slice(&rgce_ptg_array());

    let cce: u16 = subexpr.len().try_into().expect("subexpression length fits u16");

    let mut rgce = vec![0x29];
    rgce.extend_from_slice(&cce.to_le_bytes());
    rgce.extend_from_slice(&subexpr);
    // Visible PtgArray follows.
    rgce.extend_from_slice(&rgce_ptg_array());
    rgce
}

fn rgce_memfunc_with_ptgexp_and_array_subexpr() -> Vec<u8> {
    // PtgMemFunc: [ptg=0x29][cce: u16][subexpression bytes...]
    //
    // Some workbooks use `PtgExp` (shared formula placeholder) inside the non-printing
    // subexpression. Ensure we skip its payload so we can still find and consume any nested
    // `PtgArray` tokens.
    //
    // PtgExp: [ptg=0x01][row: u16][col: u16]
    let mut subexpr = vec![0x01];
    subexpr.extend_from_slice(&0u16.to_le_bytes()); // row
    subexpr.extend_from_slice(&0u16.to_le_bytes()); // col
    subexpr.extend_from_slice(&rgce_ptg_array());

    let cce: u16 = subexpr.len().try_into().expect("subexpression length fits u16");

    let mut rgce = vec![0x29];
    rgce.extend_from_slice(&cce.to_le_bytes());
    rgce.extend_from_slice(&subexpr);
    // Visible PtgArray follows.
    rgce.extend_from_slice(&rgce_ptg_array());
    rgce
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
fn decode_ptg_array_inside_memfunc_advances_rgcb_cursor() {
    // First PtgArray is inside the PtgMemFunc payload (non-printing), second is visible.
    let rgce = rgce_memfunc_with_array_subexpr();

    let mut rgcb = Vec::new();
    // First array constant: {111}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&111f64.to_le_bytes());
    // Second array constant: {222}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&222f64.to_le_bytes());

    let decoded = decode_rgce_with_rgcb(&rgce, &rgcb).expect("decode");
    assert_eq!(decoded, "{222}");
}

#[test]
fn decode_ptg_array_inside_memfunc_with_ptgname_advances_rgcb_cursor() {
    // Like `decode_ptg_array_inside_memfunc_advances_rgcb_cursor`, but with a `PtgName` token
    // before the nested `PtgArray`. This ensures our subexpression scanner skips the full 6-byte
    // PtgName payload (nameId + reserved) so we still consume the nested array constant block.
    let rgce = rgce_memfunc_with_name_and_array_subexpr();

    let mut rgcb = Vec::new();
    // First array constant: {111}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&111f64.to_le_bytes());
    // Second array constant: {222}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&222f64.to_le_bytes());

    let decoded = decode_rgce_with_rgcb(&rgce, &rgcb).expect("decode");
    assert_eq!(decoded, "{222}");
}

#[test]
fn decode_ptg_array_inside_memfunc_with_ptgexp_advances_rgcb_cursor() {
    // Like `decode_ptg_array_inside_memfunc_advances_rgcb_cursor`, but with a `PtgExp` token
    // before the nested `PtgArray`. This ensures our subexpression scanner skips the full 4-byte
    // PtgExp payload (row + col) so we still consume the nested array constant block.
    let rgce = rgce_memfunc_with_ptgexp_and_array_subexpr();

    let mut rgcb = Vec::new();
    // First array constant: {111}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&111f64.to_le_bytes());
    // Second array constant: {222}
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // cols_minus1
    rgcb.extend_from_slice(&0u16.to_le_bytes()); // rows_minus1
    rgcb.push(0x01);
    rgcb.extend_from_slice(&222f64.to_le_bytes());

    let decoded = decode_rgce_with_rgcb(&rgce, &rgcb).expect("decode");
    assert_eq!(decoded, "{222}");
}

#[test]
fn decode_ptg_array_without_rgcb_is_unsupported() {
    let rgce = rgce_ptg_array();
    match decode_rgce(&rgce) {
        Err(DecodeRgceError::UnsupportedToken { offset: 0, ptg: 0x20 }) => {}
        other => panic!("expected UnsupportedToken(0x20), got {other:?}"),
    }
}
