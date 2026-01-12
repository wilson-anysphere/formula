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
