use formula_biff::decode_rgce;
use pretty_assertions::assert_eq;

fn ptg_ref_err3d(ptg: u8) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(ptg);
    out.extend_from_slice(&0u16.to_le_bytes()); // ixti
    out.extend_from_slice(&0u32.to_le_bytes()); // row
    out.extend_from_slice(&0u16.to_le_bytes()); // col+flags
    out
}

fn ptg_area_err3d(ptg: u8) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(ptg);
    out.extend_from_slice(&0u16.to_le_bytes()); // ixti
    out.extend_from_slice(&0u32.to_le_bytes()); // row1
    out.extend_from_slice(&0u32.to_le_bytes()); // row2
    out.extend_from_slice(&0u16.to_le_bytes()); // col1+flags
    out.extend_from_slice(&0u16.to_le_bytes()); // col2+flags
    out
}

fn ptg_int(n: u16) -> [u8; 3] {
    let [lo, hi] = n.to_le_bytes();
    [0x1E, lo, hi] // PtgInt
}

#[test]
fn decodes_ptg_ref_err3d_variants() {
    for ptg in [0x3C, 0x5C, 0x7C] {
        let rgce = ptg_ref_err3d(ptg);
        let text = decode_rgce(&rgce).expect("decode");
        assert_eq!(text, "#REF!", "ptg=0x{ptg:02X}");
    }
}

#[test]
fn decodes_ptg_area_err3d_variants() {
    for ptg in [0x3D, 0x5D, 0x7D] {
        let rgce = ptg_area_err3d(ptg);
        let text = decode_rgce(&rgce).expect("decode");
        assert_eq!(text, "#REF!", "ptg=0x{ptg:02X}");
    }
}

#[test]
fn decodes_3d_error_refs_inside_expression() {
    // #REF!+1 (using PtgRefErr3d as the left operand).
    let mut rgce = ptg_ref_err3d(0x3C);
    rgce.extend_from_slice(&ptg_int(1));
    rgce.push(0x03); // PtgAdd
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "#REF!+1");

    // #REF!+2 (using PtgAreaErr3d as the left operand).
    let mut rgce = ptg_area_err3d(0x3D);
    rgce.extend_from_slice(&ptg_int(2));
    rgce.push(0x03); // PtgAdd
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "#REF!+2");
}

