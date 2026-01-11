use formula_engine::parse_formula;
use formula_xlsb::rgce::{decode_rgce, decode_rgce_with_context};
use formula_xlsb::workbook_context::WorkbookContext;
use pretty_assertions::assert_eq;

fn rgce_area(ptg: u8) -> Vec<u8> {
    // A1:A10 in BIFF12 encoding:
    // - rows are 0-indexed u32
    // - cols are 0-indexed u14 in a u16 where:
    //   - bit 14 (0x4000): row relative
    //   - bit 15 (0x8000): col relative
    let mut out = vec![ptg];
    out.extend_from_slice(&0u32.to_le_bytes()); // rowFirst = 0 (A1)
    out.extend_from_slice(&9u32.to_le_bytes()); // rowLast  = 9 (A10)
    out.extend_from_slice(&0xC000u16.to_le_bytes()); // colFirst = A, relative row/col
    out.extend_from_slice(&0xC000u16.to_le_bytes()); // colLast  = A, relative row/col
    out
}

fn rgce_ref(ptg: u8) -> Vec<u8> {
    // A1 as a PtgRef* token: [ptg][row: u32][col: u16]
    let mut out = vec![ptg];
    out.extend_from_slice(&0u32.to_le_bytes()); // row = 0 (A1)
    out.extend_from_slice(&0xC000u16.to_le_bytes()); // col = A, relative row/col
    out
}

fn assert_parses_and_roundtrips(src: &str) {
    let ast = parse_formula(src, Default::default()).expect("formula should parse");
    let back = ast.to_string(Default::default()).expect("serialize");
    assert_eq!(back, src);
}

#[test]
fn decodes_ptg_areav_as_explicit_implicit_intersection() {
    // PtgAreaV (value class) should render as `@` to preserve legacy implicit intersection.
    let rgce = rgce_area(0x45);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@A1:A10");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_ptg_area_ref_class_without_at() {
    // PtgArea (ref class) should not render `@`.
    let rgce = rgce_area(0x25);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "A1:A10");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_ptg_area3dv_with_sheet_prefix_and_at() {
    // PtgArea3dV: [ptg][ixti: u16][area...]
    let mut rgce = vec![0x5B];
    rgce.extend_from_slice(&1u16.to_le_bytes()); // Sheet2 (by index in our decode context)
    rgce.extend_from_slice(&0u32.to_le_bytes());
    rgce.extend_from_slice(&9u32.to_le_bytes());
    rgce.extend_from_slice(&0xC000u16.to_le_bytes());
    rgce.extend_from_slice(&0xC000u16.to_le_bytes());

    let mut ctx = WorkbookContext::default();
    ctx.add_extern_sheet("Sheet2", "Sheet2", 1);

    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");

    assert_eq!(text, "@Sheet2!A1:A10");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn does_not_emit_at_for_single_cell_ptg_refv() {
    let rgce = rgce_ref(0x44);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "A1");
    assert_parses_and_roundtrips(&text);
}
