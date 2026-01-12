use formula_engine::parse_formula;
use formula_xlsb::rgce::decode_rgce_with_context;
use formula_xlsb::workbook_context::WorkbookContext;
use pretty_assertions::assert_eq;

fn assert_parses_and_roundtrips(src: &str) {
    let ast = parse_formula(src, Default::default()).expect("formula should parse");
    let back = ast.to_string(Default::default()).expect("serialize");
    assert_eq!(back, src);
}

fn ctx_table1() -> WorkbookContext {
    let mut ctx = WorkbookContext::default();
    ctx.add_table(1, "Table1");
    ctx.add_table_column(1, 1, "Item");
    ctx.add_table_column(1, 2, "Qty");
    ctx.add_table_column(1, 3, "Price");
    ctx.add_table_column(1, 4, "Total");
    ctx
}

/// Build a `PtgExtend` structured reference token (`PtgList`) using the layout
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

/// Build a `PtgExtend` structured reference token (`PtgList`) using layout B:
/// `[table_id:u32][col_first_raw:u32][col_last_raw:u32]`, where:
/// - `col_first_raw` packs `[col_first:u16][flags:u16]`
/// - `col_last_raw` packs `[col_last:u16][reserved:u16]`
fn ptg_list_layout_b(
    table_id: u32,
    flags: u16,
    col_first: u16,
    col_last: u16,
    ptg: u8,
) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(ptg);
    out.push(0x19); // etpg=0x19 (PtgList)
    out.extend_from_slice(&table_id.to_le_bytes());

    let col_first_raw = (col_first as u32) | ((flags as u32) << 16);
    let col_last_raw = col_last as u32; // reserved=0
    out.extend_from_slice(&col_first_raw.to_le_bytes());
    out.extend_from_slice(&col_last_raw.to_le_bytes());
    out
}

/// Build a `PtgExtend` structured reference token (`PtgList`) using layout C:
/// `[table_id:u32][flags:u32][col_spec:u32]`, where:
/// - `col_spec` packs `[col_first:u16][col_last:u16]`
fn ptg_list_layout_c(
    table_id: u32,
    flags: u32,
    col_first: u16,
    col_last: u16,
    ptg: u8,
) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(ptg);
    out.push(0x19); // etpg=0x19 (PtgList)
    out.extend_from_slice(&table_id.to_le_bytes());
    out.extend_from_slice(&flags.to_le_bytes());

    let col_spec = (col_first as u32) | ((col_last as u32) << 16);
    out.extend_from_slice(&col_spec.to_le_bytes());
    out
}

#[test]
fn decodes_structured_ref_table_column() {
    let ctx = ctx_table1();
    let rgce = ptg_list(1, 0x0000, 2, 2, 0x18); // Table1[Qty]
    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "Table1[Qty]");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_structured_ref_this_row() {
    // Flags are best-effort; the decoder treats 0x0010 as "#This Row".
    let ctx = ctx_table1();
    let rgce = ptg_list(1, 0x0010, 2, 2, 0x18); // [@Qty]
    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "[@Qty]");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_structured_ref_this_row_range() {
    // Flags are best-effort; the decoder treats 0x0010 as "#This Row".
    let ctx = ctx_table1();
    let rgce = ptg_list(1, 0x0010, 2, 4, 0x18); // [@[Qty]:[Total]]
    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "[@[Qty]:[Total]]");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_structured_ref_headers_column() {
    // Flags are best-effort; the decoder treats 0x0002 as "#Headers".
    let ctx = ctx_table1();
    let rgce = ptg_list(1, 0x0002, 2, 2, 0x18); // Table1[[#Headers],[Qty]]
    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "Table1[[#Headers],[Qty]]");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_structured_ref_item_only_all() {
    let ctx = ctx_table1();
    let rgce = ptg_list(1, 0x0001, 0, 0, 0x18); // Table1[#All]
    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "Table1[#All]");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_structured_ref_value_class_emits_explicit_implicit_intersection() {
    let ctx = ctx_table1();
    let rgce = ptg_list(1, 0x0000, 2, 2, 0x38); // value-class PtgExtendV
    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "@Table1[Qty]");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_structured_ref_table_column_layout_b() {
    let ctx = ctx_table1();
    let rgce = ptg_list_layout_b(1, 0x0000, 2, 2, 0x18); // Table1[Qty]
    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "Table1[Qty]");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_structured_ref_headers_column_layout_b() {
    // Use an unknown high flag bit to ensure layout A produces an unmapped column id, so the
    // context-based scoring must select layout B.
    let ctx = ctx_table1();
    let rgce = ptg_list_layout_b(1, 0x8002, 2, 2, 0x18); // Table1[[#Headers],[Qty]]
    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "Table1[[#Headers],[Qty]]");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_structured_ref_table_column_layout_c() {
    // Add junk in the upper 16 bits of the u32 flags field so that layout A would interpret it as
    // a wildly out-of-range column id, forcing the best-effort decoder to prefer layout C.
    let ctx = ctx_table1();
    let rgce = ptg_list_layout_c(1, 0xF000_0000, 2, 2, 0x18); // Table1[Qty]
    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "Table1[Qty]");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_structured_ref_headers_column_layout_c() {
    let ctx = ctx_table1();
    // Use an additional unknown flag bit so that layout B would interpret the low 16 bits of the
    // u32 flags field as an out-of-range column id, ensuring the scoring prefers layout C.
    let rgce = ptg_list_layout_c(1, 0x8002, 2, 2, 0x18); // Table1[[#Headers],[Qty]]
    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "Table1[[#Headers],[Qty]]");
    assert_parses_and_roundtrips(&text);
}
