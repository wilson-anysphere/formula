use formula_biff::{decode_rgce, decode_rgce_with_base, DecodeRgceError};
use pretty_assertions::assert_eq;

fn ptg_int(n: u16) -> [u8; 3] {
    let [lo, hi] = n.to_le_bytes();
    [0x1E, lo, hi] // PtgInt
}

fn ptg_referr() -> [u8; 7] {
    // PtgRefErr: [row: u32][col: u16]
    [0x2A, 0, 0, 0, 0, 0, 0]
}

fn ptg_areaerr() -> [u8; 13] {
    // PtgAreaErr: [rowFirst: u32][rowLast: u32][colFirst: u16][colLast: u16]
    [0x2B, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
}

fn ptg_refn(row_off: i32, col_off: i16) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(0x2C); // PtgRefN
    out.extend_from_slice(&row_off.to_le_bytes());
    out.extend_from_slice(&col_off.to_le_bytes());
    out
}

fn ptg_arean(row1_off: i32, row2_off: i32, col1_off: i16, col2_off: i16) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(0x2D); // PtgAreaN
    out.extend_from_slice(&row1_off.to_le_bytes());
    out.extend_from_slice(&row2_off.to_le_bytes());
    out.extend_from_slice(&col1_off.to_le_bytes());
    out.extend_from_slice(&col2_off.to_le_bytes());
    out
}

#[test]
fn decodes_ptgreferr_and_consumes_payload() {
    // `#REF!+1`
    let mut rgce = Vec::new();
    rgce.extend_from_slice(&ptg_referr());
    rgce.extend_from_slice(&ptg_int(1));
    rgce.push(0x03); // PtgAdd
    assert_eq!(decode_rgce(&rgce).expect("decode"), "#REF!+1");
}

#[test]
fn decodes_ptgareaerr_and_consumes_payload() {
    // `#REF!+1`
    let mut rgce = Vec::new();
    rgce.extend_from_slice(&ptg_areaerr());
    rgce.extend_from_slice(&ptg_int(1));
    rgce.push(0x03); // PtgAdd
    assert_eq!(decode_rgce(&rgce).expect("decode"), "#REF!+1");
}

#[test]
fn decodes_ptgrefn_with_base() {
    // Base cell is C3 (row0=2, col0=2). Offsets (-2, -2) point at A1.
    let rgce = ptg_refn(-2, -2);
    assert_eq!(decode_rgce_with_base(&rgce, 2, 2).expect("decode"), "A1");
}

#[test]
fn decodes_ptgarean_with_base() {
    // Base cell is C3 (row0=2, col0=2). Offsets (-1..=1, -1..=1) -> B2:D4.
    let rgce = ptg_arean(-1, 1, -1, 1);
    assert_eq!(decode_rgce_with_base(&rgce, 2, 2).expect("decode"), "B2:D4");
}

#[test]
fn ptgrefn_out_of_bounds_emits_ref() {
    // Base cell A1 + row_off=-1 is out-of-bounds.
    let rgce = ptg_refn(-1, 0);
    assert_eq!(decode_rgce_with_base(&rgce, 0, 0).expect("decode"), "#REF!");
}

#[test]
fn decode_rgce_errors_on_ptgrefn_and_does_not_misparse_following_tokens() {
    // `A1+1` expressed using a relative ref token.
    let mut rgce = ptg_refn(-2, -2);
    rgce.extend_from_slice(&ptg_int(1));
    rgce.push(0x03); // PtgAdd

    // The base-unaware decoder should fail fast on the RefN token.
    match decode_rgce(&rgce) {
        Err(DecodeRgceError::UnsupportedToken { offset: 0, ptg }) => assert_eq!(ptg, 0x2C),
        other => panic!("expected UnsupportedToken(0x2C), got {other:?}"),
    }

    // The base-aware decoder should consume the full RefN payload and continue decoding.
    assert_eq!(
        decode_rgce_with_base(&rgce, 2, 2).expect("decode with base"),
        "A1+1"
    );
}
