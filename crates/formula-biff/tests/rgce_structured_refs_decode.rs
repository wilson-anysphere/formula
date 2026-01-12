use formula_biff::decode_rgce;
use pretty_assertions::assert_eq;

/// Build a BIFF12 structured reference token (`PtgList`) encoded as `PtgExtend` + `etpg=0x19`.
///
/// Payload layout (MS-XLSB 2.5.198.51):
/// `[table_id: u32][flags: u16][col_first: u16][col_last: u16][reserved: u16]`.
fn ptg_list(table_id: u32, flags: u16, col_first: u16, col_last: u16, ptg: u8) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(ptg);
    out.push(0x19); // etpg=0x19 (PtgList)
    out.extend_from_slice(&table_id.to_le_bytes());
    out.extend_from_slice(&flags.to_le_bytes());
    out.extend_from_slice(&col_first.to_le_bytes());
    out.extend_from_slice(&col_last.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // reserved
    out
}

#[test]
fn decodes_structured_ref_table_column() {
    let rgce = ptg_list(1, 0x0000, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[Column2]");
}

#[test]
fn decodes_structured_ref_this_row() {
    // Best-effort: 0x0010 is treated as "#This Row" and rendered using the `[@Col]` shorthand.
    let rgce = ptg_list(1, 0x0010, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "[@Column2]");
}

#[test]
fn decodes_structured_ref_this_row_all_columns() {
    let rgce = ptg_list(1, 0x0010, 0, 0, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "[@]");
}

#[test]
fn decodes_structured_ref_headers_column() {
    let rgce = ptg_list(1, 0x0002, 2, 2, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[[#Headers],[Column2]]");
}

#[test]
fn decodes_structured_ref_item_only_all() {
    let rgce = ptg_list(1, 0x0001, 0, 0, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[#All]");
}

#[test]
fn decodes_structured_ref_column_range() {
    let rgce = ptg_list(1, 0x0000, 2, 4, 0x18);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "Table1[[Column2]:[Column4]]");
}

#[test]
fn decodes_structured_ref_value_class_emits_explicit_implicit_intersection() {
    let rgce = ptg_list(1, 0x0000, 2, 2, 0x38);
    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "@Table1[Column2]");
}

