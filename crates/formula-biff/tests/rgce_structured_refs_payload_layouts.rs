use formula_biff::decode_rgce;
use pretty_assertions::assert_eq;

fn normalize(formula: &str) -> String {
    let ast = formula_engine::parse_formula(formula, formula_engine::ParseOptions::default())
        .expect("parse formula");
    ast.to_string(formula_engine::SerializeOptions {
        omit_equals: true,
        ..Default::default()
    })
    .expect("serialize formula")
}

/// Build a minimal `rgce` stream containing a single `PtgExtend` token with `etpg=0x19` (PtgList).
fn rgce_ptg_list_with_payload(payload: [u8; 12]) -> Vec<u8> {
    let mut out = Vec::new();
    out.push(0x18); // PtgExtend (reference class)
    out.push(0x19); // etpg=0x19 (PtgList / structured ref)
    out.extend_from_slice(&payload);
    out
}

/// Payload layout B (observed in the wild):
/// `[table_id: u32][col_first_raw: u32][col_last_raw: u32]`
/// where `col_first_raw` packs `[col_first: u16][flags: u16]` (little endian), and `col_last_raw`
/// packs `[col_last: u16][reserved: u16]`.
fn ptg_list_payload_layout_b(
    table_id: u32,
    flags: u16,
    col_first: u16,
    col_last: u16,
) -> [u8; 12] {
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
fn ptg_list_payload_layout_c(
    table_id: u32,
    flags: u32,
    col_first: u16,
    col_last: u16,
) -> [u8; 12] {
    let col_spec = (u32::from(col_last) << 16) | u32::from(col_first);

    let mut payload = [0u8; 12];
    payload[0..4].copy_from_slice(&table_id.to_le_bytes());
    payload[4..8].copy_from_slice(&flags.to_le_bytes());
    payload[8..12].copy_from_slice(&col_spec.to_le_bytes());
    payload
}

#[test]
fn decodes_structured_ref_payload_layout_b() {
    // Use FLAG_THIS_ROW so the *wrong* payload interpretation will treat it as a non-this-row
    // flag value and/or will decode a mismatched column id.
    let payload = ptg_list_payload_layout_b(1, 0x0010, 2, 2);
    let rgce = rgce_ptg_list_with_payload(payload);

    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "[@Column2]");
    assert_eq!(normalize(&text), normalize("[@Column2]"));
}

#[test]
fn decodes_structured_ref_payload_layout_c() {
    let payload = ptg_list_payload_layout_c(1, 0x0010, 2, 2);
    let rgce = rgce_ptg_list_with_payload(payload);

    let text = decode_rgce(&rgce).expect("decode");
    assert_eq!(text, "[@Column2]");
    assert_eq!(normalize(&text), normalize("[@Column2]"));
}

