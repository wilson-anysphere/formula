use formula_engine::{parse_formula, ParseOptions, SerializeOptions};
use formula_xlsb::rgce::{decode_rgce_with_context, encode_rgce_with_context, CellCoord};
use formula_xlsb::workbook_context::WorkbookContext;
use pretty_assertions::assert_eq;

#[test]
fn sheet_range_3d_ref_decodes_as_quoted_prefix_and_reencodes_with_same_ixti() {
    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet1", "Sheet3", 7);

    // PtgRef3d: [ptg][ixti: u16][row: u32][col: u16]
    let mut rgce = vec![0x3A];
    rgce.extend_from_slice(&7u16.to_le_bytes());
    rgce.extend_from_slice(&0u32.to_le_bytes()); // row = 0 (A1)
    rgce.extend_from_slice(&0xC000u16.to_le_bytes()); // col = A, relative row/col

    let decoded = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(decoded, "'Sheet1:Sheet3'!A1");

    // The unquoted form (`Sheet1:Sheet3!A1`) is not parseable by `formula-engine`. Ensure our
    // quoted prefix round-trips through the parser.
    let ast =
        parse_formula(&format!("={decoded}"), ParseOptions::default()).expect("parse formula");
    let normalized = ast
        .to_string(SerializeOptions {
            omit_equals: true,
            ..Default::default()
        })
        .expect("serialize formula");

    let encoded = encode_rgce_with_context(&format!("={normalized}"), &ctx, CellCoord::new(0, 0))
        .expect("encode");
    assert_eq!(encoded.rgce[0], 0x3A);
    assert_eq!(u16::from_le_bytes([encoded.rgce[1], encoded.rgce[2]]), 7);
}
