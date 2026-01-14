use formula_xlsb::{biff12_varint, parse_sheet_bin_with_context, CellValue};
use formula_xlsb::workbook_context::WorkbookContext;
use pretty_assertions::assert_eq;
use std::io::Cursor;

fn push_record(out: &mut Vec<u8>, id: u32, data: &[u8]) {
    biff12_varint::write_record_id(out, id).expect("write record id");
    let len = u32::try_from(data.len()).expect("record too large");
    biff12_varint::write_record_len(out, len).expect("write record len");
    out.extend_from_slice(data);
}

fn ctx_table() -> WorkbookContext {
    let mut ctx = WorkbookContext::default();
    ctx.add_table(1, "Orders");
    ctx.add_table_column(1, 1, "Item");
    ctx.add_table_column(1, 2, "Qty");
    ctx
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

fn build_shared_structured_ref_sheet(base_rgce: &[u8]) -> Vec<u8> {
    // Record IDs follow the conventions used by `formula-xlsb`'s BIFF12 reader.
    const WORKSHEET_BEGIN: u32 = 0x0081;
    const WORKSHEET_END: u32 = 0x0082;
    const SHEETDATA_BEGIN: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const DIMENSION: u32 = 0x0094;

    const ROW: u32 = 0x0000;
    const FMLA_NUM: u32 = 0x0009;
    const SHR_FMLA: u32 = 0x0010;

    let mut sheet = Vec::new();
    push_record(&mut sheet, WORKSHEET_BEGIN, &[]);

    // BrtWsDim: cover B1:B2.
    let mut dim = Vec::new();
    dim.extend_from_slice(&0u32.to_le_bytes()); // r1
    dim.extend_from_slice(&1u32.to_le_bytes()); // r2
    dim.extend_from_slice(&1u32.to_le_bytes()); // c1
    dim.extend_from_slice(&1u32.to_le_bytes()); // c2
    push_record(&mut sheet, DIMENSION, &dim);

    push_record(&mut sheet, SHEETDATA_BEGIN, &[]);

    // Row 0
    push_record(&mut sheet, ROW, &0u32.to_le_bytes());

    // Shared formula over B1:B2: base rgce contains a structured reference token.
    let mut shr_fmla = Vec::new();
    shr_fmla.extend_from_slice(&0u32.to_le_bytes()); // range_r1
    shr_fmla.extend_from_slice(&1u32.to_le_bytes()); // range_r2
    shr_fmla.extend_from_slice(&1u32.to_le_bytes()); // range_c1
    shr_fmla.extend_from_slice(&1u32.to_le_bytes()); // range_c2
    shr_fmla.extend_from_slice(&(base_rgce.len() as u32).to_le_bytes());
    shr_fmla.extend_from_slice(base_rgce);
    push_record(&mut sheet, SHR_FMLA, &shr_fmla);

    // Base cell formula in B1 (full rgce).
    let mut b1 = Vec::new();
    b1.extend_from_slice(&1u32.to_le_bytes()); // col B
    b1.extend_from_slice(&0u32.to_le_bytes()); // style
    b1.extend_from_slice(&0.0f64.to_le_bytes()); // cached value
    b1.extend_from_slice(&0u16.to_le_bytes()); // flags
    b1.extend_from_slice(&(base_rgce.len() as u32).to_le_bytes());
    b1.extend_from_slice(base_rgce);
    push_record(&mut sheet, FMLA_NUM, &b1);

    // Row 1
    push_record(&mut sheet, ROW, &1u32.to_le_bytes());

    // Dependent cell B2 uses PtgExp to reference base cell B1 (row=0,col=1).
    let ptgexp: [u8; 5] = [0x01, 0x00, 0x00, 0x01, 0x00];
    let mut b2 = Vec::new();
    b2.extend_from_slice(&1u32.to_le_bytes()); // col B
    b2.extend_from_slice(&0u32.to_le_bytes()); // style
    b2.extend_from_slice(&0.0f64.to_le_bytes()); // cached value
    b2.extend_from_slice(&0u16.to_le_bytes()); // flags
    b2.extend_from_slice(&(ptgexp.len() as u32).to_le_bytes());
    b2.extend_from_slice(&ptgexp);
    push_record(&mut sheet, FMLA_NUM, &b2);

    push_record(&mut sheet, SHEETDATA_END, &[]);
    push_record(&mut sheet, WORKSHEET_END, &[]);
    sheet
}

#[test]
fn materializes_shared_formula_containing_structured_ref_payload_layout_b() {
    let payload = ptg_list_payload_layout_b(1, 0x0010, 2, 2); // #This Row + Qty

    // Structured references can appear as PtgExtend (0x18), PtgExtendV (0x38), or PtgExtendA
    // (0x58). Shared formula materialization should preserve the class byte and payload for all
    // of them.
    for &ptg in &[0x18u8, 0x38u8, 0x58u8] {
        let base_rgce: Vec<u8> = {
            let mut v = Vec::new();
            v.push(ptg);
            v.push(0x19); // etpg=PtgList
            v.extend_from_slice(&payload);
            v
        };

        let sheet = build_shared_structured_ref_sheet(&base_rgce);
        let ctx = ctx_table();
        let parsed =
            parse_sheet_bin_with_context(&mut Cursor::new(sheet), &[], &ctx).expect("parse sheet");

        let b2 = parsed
            .cells
            .iter()
            .find(|c| c.row == 1 && c.col == 1)
            .expect("B2 present");
        assert_eq!(b2.value, CellValue::Number(0.0));

        let formula = b2.formula.as_ref().expect("B2 formula");
        assert_eq!(formula.text.as_deref(), Some("[@Qty]"));
        // Ensure we stored a materialized rgce (not just PtgExp).
        assert_eq!(formula.rgce.first().copied(), Some(ptg));
        assert_eq!(formula.rgce, base_rgce, "ptg=0x{ptg:02X}");
    }
}

#[test]
fn materializes_shared_formula_containing_structured_ref_payload_layout_c() {
    let payload = ptg_list_payload_layout_c(1, 0x0010, 2, 2); // #This Row + Qty
    for &ptg in &[0x18u8, 0x38u8, 0x58u8] {
        let base_rgce: Vec<u8> = {
            let mut v = Vec::new();
            v.push(ptg);
            v.push(0x19); // etpg=PtgList
            v.extend_from_slice(&payload);
            v
        };

        let sheet = build_shared_structured_ref_sheet(&base_rgce);
        let ctx = ctx_table();
        let parsed =
            parse_sheet_bin_with_context(&mut Cursor::new(sheet), &[], &ctx).expect("parse sheet");

        let b2 = parsed
            .cells
            .iter()
            .find(|c| c.row == 1 && c.col == 1)
            .expect("B2 present");
        assert_eq!(b2.value, CellValue::Number(0.0));

        let formula = b2.formula.as_ref().expect("B2 formula");
        assert_eq!(formula.text.as_deref(), Some("[@Qty]"));
        // Ensure we stored a materialized rgce (not just PtgExp).
        assert_eq!(formula.rgce.first().copied(), Some(ptg));
        assert_eq!(formula.rgce, base_rgce, "ptg=0x{ptg:02X}");
    }
}
