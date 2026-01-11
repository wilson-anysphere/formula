use formula_engine::{parse_formula, ParseOptions};
use formula_xlsb::rgce::{decode_rgce_with_rgcb, encode_rgce_with_context, CellCoord};
use formula_xlsb::workbook_context::WorkbookContext;
use pretty_assertions::assert_eq;

#[test]
fn decodes_ptgarray_with_trailing_array_data() {
    // rgce = [PtgArray][unused7]
    let rgce = vec![0x20, 0, 0, 0, 0, 0, 0, 0];

    // rgcb = [cols_minus1: u16][rows_minus1: u16] + cells (row-major)
    let mut rgcb = Vec::new();
    rgcb.extend_from_slice(&1u16.to_le_bytes()); // 2 cols
    rgcb.extend_from_slice(&1u16.to_le_bytes()); // 2 rows

    for v in [1.0f64, 2.0, 3.0, 4.0] {
        rgcb.push(0x01); // xltypeNum
        rgcb.extend_from_slice(&v.to_le_bytes());
    }

    let text = decode_rgce_with_rgcb(&rgce, &rgcb).expect("decode rgce");
    assert_eq!(text, "{1,2;3,4}");

    parse_formula(&text, ParseOptions::default()).expect("formula-engine parses decoded array");
}

#[test]
fn encode_decode_roundtrip_array_constant() {
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("={1,2;3,4}", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(!encoded.rgcb.is_empty());
    let text = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(text, "{1,2;3,4}");
}

#[test]
fn encode_decode_roundtrip_sum_over_array_constant() {
    let ctx = WorkbookContext::default();
    let encoded =
        encode_rgce_with_context("=SUM({1,2,3})", &ctx, CellCoord::new(0, 0)).expect("encode");
    assert!(!encoded.rgcb.is_empty());
    let text = decode_rgce_with_rgcb(&encoded.rgce, &encoded.rgcb).expect("decode");
    assert_eq!(text, "SUM({1,2,3})");

    parse_formula(&format!("={text}"), ParseOptions::default())
        .expect("formula-engine parses decoded formula");
}
