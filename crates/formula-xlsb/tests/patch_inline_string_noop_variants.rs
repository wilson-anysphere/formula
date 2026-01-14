use formula_xlsb::biff12_varint;
use formula_xlsb::{parse_sheet_bin, patch_sheet_bin, CellEdit, CellValue};
use pretty_assertions::assert_eq;
use std::io::Cursor;

fn biff12_record(id: u32, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    biff12_varint::write_record_id(&mut out, id).expect("write record id");
    biff12_varint::write_record_len(&mut out, payload.len() as u32).expect("write record len");
    out.extend_from_slice(payload);
    out
}

fn utf16_le_bytes(s: &str) -> Vec<u8> {
    let mut out = Vec::new();
    for unit in s.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn sheet_with_single_cell_st(cell_st_payload: &[u8]) -> Vec<u8> {
    // Record ids (subset):
    // - BrtBeginSheetData 0x0091
    // - BrtEndSheetData   0x0092
    // - BrtRow            0x0000
    // - BrtCellSt         0x0006
    const SHEETDATA: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const ROW: u32 = 0x0000;
    const CELL_ST: u32 = 0x0006;

    let mut sheet_bin = Vec::new();
    sheet_bin.extend_from_slice(&biff12_record(SHEETDATA, &[]));
    sheet_bin.extend_from_slice(&biff12_record(ROW, &0u32.to_le_bytes()));
    sheet_bin.extend_from_slice(&biff12_record(CELL_ST, cell_st_payload));
    sheet_bin.extend_from_slice(&biff12_record(SHEETDATA_END, &[]));
    sheet_bin
}

#[test]
fn noop_inline_string_flagged_layout_with_missing_rich_phonetic_blocks_is_preserved() {
    // Flagged BrtCellSt layout:
    //   [col:u32][style:u32][cch:u32][flags:u8][utf16 bytes...][optional extras...]
    //
    // Some producers set the rich/phonetic bits in `flags` but omit the corresponding payload
    // blocks. For a no-op edit we should preserve the record bytes verbatim (and not error).
    let text = "Hi".to_string();
    let cch = text.encode_utf16().count() as u32;
    let utf16 = utf16_le_bytes(&text);

    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&cch.to_le_bytes());
    cell_st_payload.push(0x83); // flags: rich + phonetic + reserved bit, but no blocks follow
    cell_st_payload.extend_from_slice(&utf16);

    let sheet_bin = sheet_with_single_cell_st(&cell_st_payload);
    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text(text),
        new_style: None,
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    };

    let patched = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");
    assert_eq!(patched, sheet_bin);
}

#[test]
fn noop_inline_string_simple_layout_with_trailing_bytes_is_preserved() {
    // Simple BrtCellSt layout:
    //   [col:u32][style:u32][cch:u32][utf16 bytes...]
    //
    // Some streams contain extra trailing bytes after the UTF-16 text. A no-op edit should
    // preserve those bytes.
    let text = "Hello".to_string();
    let cch = text.encode_utf16().count() as u32;
    let utf16 = utf16_le_bytes(&text);

    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&cch.to_le_bytes());
    cell_st_payload.extend_from_slice(&utf16);
    cell_st_payload.extend_from_slice(&[0xDE, 0xAD, 0xBE, 0xEF]); // trailing junk bytes

    let sheet_bin = sheet_with_single_cell_st(&cell_st_payload);
    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text(text),
        new_style: None,
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    };

    let patched = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");
    assert_eq!(patched, sheet_bin);
}

#[test]
fn flagged_inline_string_text_change_is_not_treated_as_noop() {
    // Flagged layout (flags byte present) with a real text change should not be treated as a
    // no-op; the patcher should rewrite the cell.
    let old_text = "Old".to_string();
    let new_text = "New".to_string();

    let cch = old_text.encode_utf16().count() as u32;
    let utf16 = utf16_le_bytes(&old_text);

    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&cch.to_le_bytes());
    cell_st_payload.push(0); // flags
    cell_st_payload.extend_from_slice(&utf16);

    let sheet_bin = sheet_with_single_cell_st(&cell_st_payload);
    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text(new_text.clone()),
        new_style: None,
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    };

    let patched = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");
    assert_ne!(patched, sheet_bin, "expected patched output to differ");

    let parsed = parse_sheet_bin(&mut Cursor::new(&patched), &[]).expect("parse patched sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("find cell");
    assert_eq!(cell.value, CellValue::Text(new_text));
}

#[test]
fn malformed_flagged_inline_string_text_change_succeeds_and_is_parseable() {
    // Some producers set rich/phonetic bits in the flags byte but omit the corresponding blocks.
    // The patcher should still be able to rewrite the text and emit a parseable record by
    // inserting empty block headers.
    let old_text = "Old".to_string();
    let new_text = "New".to_string();

    let cch = old_text.encode_utf16().count() as u32;
    let utf16 = utf16_le_bytes(&old_text);

    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&cch.to_le_bytes());
    cell_st_payload.push(0x83); // flags: rich + phonetic (+ reserved), but blocks are missing
    cell_st_payload.extend_from_slice(&utf16);

    let sheet_bin = sheet_with_single_cell_st(&cell_st_payload);
    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text(new_text.clone()),
        new_style: None,
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    };

    let patched = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");

    let parsed = parse_sheet_bin(&mut Cursor::new(&patched), &[]).expect("parse patched sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("find cell");
    assert_eq!(cell.value, CellValue::Text(new_text));
}

#[test]
fn parser_reads_malformed_flagged_inline_string_without_blocks() {
    // Regression: some BrtCellSt records include a flags byte with rich/phonetic bits set but omit
    // the corresponding blocks. The reader should still decode the UTF-16 text correctly.
    let text = "Old".to_string();
    let cch = text.encode_utf16().count() as u32;
    let utf16 = utf16_le_bytes(&text);

    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&cch.to_le_bytes());
    cell_st_payload.push(0x83); // flags: rich + phonetic (+ reserved), but blocks are missing
    cell_st_payload.extend_from_slice(&utf16);

    let sheet_bin = sheet_with_single_cell_st(&cell_st_payload);
    let parsed = parse_sheet_bin(&mut Cursor::new(&sheet_bin), &[]).expect("parse sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("find cell");
    assert_eq!(cell.value, CellValue::Text(text));
}
