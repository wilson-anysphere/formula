use formula_engine::parse_formula;
use formula_xlsb::rgce::decode_rgce_with_context;
use formula_xlsb::workbook_context::WorkbookContext;
use pretty_assertions::assert_eq;

fn assert_parses_and_roundtrips(src: &str) {
    let ast = parse_formula(src, Default::default()).expect("formula should parse");
    let back = ast.to_string(Default::default()).expect("serialize");
    assert_eq!(back, src);
}

fn ctx_table() -> WorkbookContext {
    let mut ctx = WorkbookContext::default();
    // Use a non-default display name so the test will exercise workbook-context lookups rather
    // than relying on the fallback `Table{ID}` naming.
    ctx.add_table(1, "Orders");
    ctx.add_table_column(1, 1, "Item");
    ctx.add_table_column(1, 2, "Qty");
    ctx
}

/// Build a minimal `rgce` stream containing a single `PtgExtend` token with `etpg=0x19` (PtgList).
fn rgce_ptg_list_with_payload(payload: [u8; 12]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(0x18); // PtgExtend (reference class)
    out.push(0x19); // etpg=0x19 (PtgList / structured ref)
    out.extend_from_slice(&payload);
    out
}

/// Build a minimal `rgce` stream containing a single `PtgExtend` token with `etpg=0x19` (PtgList),
/// with extra prefix bytes inserted before the canonical 12-byte payload.
fn rgce_ptg_list_with_prefixed_payload(prefix: &[u8], payload: [u8; 12]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(0x18); // PtgExtend (reference class)
    out.push(0x19); // etpg=0x19 (PtgList / structured ref)
    out.extend_from_slice(prefix);
    out.extend_from_slice(&payload);
    out
}

/// Payload layout B (observed in the wild):
/// `[table_id: u32][col_first_raw: u32][col_last_raw: u32]`
/// where `col_first_raw` packs `[col_first: u16][flags: u16]` (little endian), and `col_last_raw`
/// packs `[col_last: u16][reserved: u16]`.
fn ptg_list_payload_layout_b(table_id: u32, flags: u16, col_first: u16, col_last: u16) -> [u8; 12] {
    let col_first_raw = (u32::from(flags) << 16) | u32::from(col_first);
    let col_last_raw = u32::from(col_last);

    let mut payload = [0u8; 12];
    payload[0..4].copy_from_slice(&table_id.to_le_bytes());
    payload[4..8].copy_from_slice(&col_first_raw.to_le_bytes());
    payload[8..12].copy_from_slice(&col_last_raw.to_le_bytes());
    payload
}

/// Payload layout C (observed in the wild):
/// `[table_id: u32][flags: u32][col_spec: u32]`
/// where `col_spec` packs `[col_first: u16][col_last: u16]` (little endian).
fn ptg_list_payload_layout_c(table_id: u32, flags: u32, col_first: u16, col_last: u16) -> [u8; 12] {
    let col_spec = (u32::from(col_last) << 16) | u32::from(col_first);

    let mut payload = [0u8; 12];
    payload[0..4].copy_from_slice(&table_id.to_le_bytes());
    payload[4..8].copy_from_slice(&flags.to_le_bytes());
    payload[8..12].copy_from_slice(&col_spec.to_le_bytes());
    payload
}

#[test]
fn decodes_structured_ref_payload_layout_b_uses_context_column_names() {
    // Use FLAG_THIS_ROW so the *wrong* payload interpretation will treat it as a (non-this-row)
    // flag value and/or will decode `col_first=16`, producing the `Column16` placeholder.
    let ctx = ctx_table();
    let payload = ptg_list_payload_layout_b(1, 0x0010, 2, 2);
    let rgce = rgce_ptg_list_with_payload(payload);

    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "[@Qty]");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_structured_ref_payload_layout_c_uses_context_column_names() {
    let ctx = ctx_table();
    let payload = ptg_list_payload_layout_c(1, 0x0010, 2, 2);
    let rgce = rgce_ptg_list_with_payload(payload);

    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "[@Qty]");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_structured_ref_payload_layout_b_with_2_byte_prefix_padding() {
    let ctx = ctx_table();
    let payload = ptg_list_payload_layout_b(1, 0x0010, 2, 2);
    let rgce = rgce_ptg_list_with_prefixed_payload(&[0, 0], payload);

    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "[@Qty]");
    assert_parses_and_roundtrips(&text);
}

#[test]
fn decodes_structured_ref_payload_layout_b_with_4_byte_prefix_padding() {
    let ctx = ctx_table();
    let payload = ptg_list_payload_layout_b(1, 0x0010, 2, 2);
    let rgce = rgce_ptg_list_with_prefixed_payload(&[0, 0, 0, 0], payload);

    let text = decode_rgce_with_context(&rgce, &ctx).expect("decode");
    assert_eq!(text, "[@Qty]");
    assert_parses_and_roundtrips(&text);
}
