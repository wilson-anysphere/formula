use formula_xlsb::rgce::{decode_rgce, encode_rgce_with_context, CellCoord};
use formula_xlsb::workbook_context::WorkbookContext;

#[test]
fn decodes_spill_operator_token() {
    // rgce for `A1#`:
    // - PtgRef (0x24) with row=0, col=0, both relative (flags 0xC000)
    // - PtgSpill (0x2F)
    let rgce = vec![0x24, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x2F];
    assert_eq!(decode_rgce(&rgce).expect("decode"), "A1#");
}

#[test]
fn encode_decode_roundtrip_spill_operator() {
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("=A1#", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert_eq!(decode_rgce(&encoded.rgce).expect("decode"), "A1#");
}
