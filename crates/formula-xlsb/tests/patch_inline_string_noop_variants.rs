use formula_xlsb::biff12_varint;
use formula_xlsb::{parse_sheet_bin, patch_sheet_bin, patch_sheet_bin_streaming, CellEdit, CellValue};
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

fn find_cell_st_payload(sheet_bin: &[u8]) -> Vec<u8> {
    const CELL_ST: u32 = 0x0006;
    let mut cursor = Cursor::new(sheet_bin);
    loop {
        let id = biff12_varint::read_record_id(&mut cursor)
            .expect("read record id")
            .expect("record id");
        let len = biff12_varint::read_record_len(&mut cursor)
            .expect("read record len")
            .expect("record len") as usize;
        let start = cursor.position() as usize;
        let end = start + len;
        let payload = sheet_bin.get(start..end).expect("payload bytes");
        cursor.set_position(end as u64);
        if id == CELL_ST {
            return payload.to_vec();
        }
    }
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
fn style_update_inline_string_simple_layout_with_trailing_bytes_is_preserved() {
    // Simple BrtCellSt layout:
    //   [col:u32][style:u32][cch:u32][utf16 bytes...]
    //
    // Some streams contain extra trailing bytes after the UTF-16 text. When applying a *style-only*
    // update (text unchanged), the patcher should preserve those bytes instead of rewriting the
    // string payload and dropping the unknown suffix.
    let text = "Hello".to_string();
    let cch = text.encode_utf16().count() as u32;
    let utf16 = utf16_le_bytes(&text);
    let trailing = [0xDE, 0xAD, 0xBE, 0xEF];

    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&cch.to_le_bytes());
    cell_st_payload.extend_from_slice(&utf16);
    cell_st_payload.extend_from_slice(&trailing);

    let sheet_bin = sheet_with_single_cell_st(&cell_st_payload);
    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text(text.clone()),
        new_style: Some(7),
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    };

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &[edit.clone()]).expect("patch sheet");

    let mut patched_stream = Vec::new();
    let changed = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &[edit])
        .expect("patch sheet streaming");
    assert!(changed);
    assert_eq!(patched_stream, patched_in_mem);

    let parsed = parse_sheet_bin(&mut Cursor::new(&patched_in_mem), &[]).expect("parse patched");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("find cell");
    assert_eq!(cell.value, CellValue::Text(text));
    assert_eq!(cell.style, 7);

    let payload = find_cell_st_payload(&patched_in_mem);
    assert_eq!(
        u32::from_le_bytes(payload[4..8].try_into().unwrap()),
        7,
        "expected style to be updated in-place"
    );
    assert!(
        payload.ends_with(&trailing),
        "expected trailing bytes to be preserved on style-only update"
    );
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

#[test]
fn parser_reads_simple_inline_string_with_trailing_bytes() {
    // Some BrtCellSt records use the simple layout but include extra trailing bytes after the
    // UTF-16 text. The reader should still decode the text correctly.
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
    let parsed = parse_sheet_bin(&mut Cursor::new(&sheet_bin), &[]).expect("parse sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("find cell");
    assert_eq!(cell.value, CellValue::Text(text));
}

#[test]
fn parser_reads_flagged_inline_string_with_low_byte_zero() {
    // Regression: some valid flagged-layout strings can start with a UTF-16 code unit whose low
    // byte is `0x00` (e.g. U+0100). Our layout heuristic must not misclassify these as the simple
    // layout (which would decode NULs / garbage).
    let text = "Ā".to_string(); // U+0100 => UTF-16LE bytes [0x00, 0x01]
    let cch = text.encode_utf16().count() as u32;
    let utf16 = utf16_le_bytes(&text);

    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&cch.to_le_bytes());
    cell_st_payload.push(0); // flags byte (flagged layout)
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

#[test]
fn parser_reads_flagged_inline_string_with_surrogate_pairs() {
    // Regression: some producers emit the flagged layout even when there are no rich/phonetic
    // extras (`flags=0`). The reader must still decode strings that contain surrogate pairs
    // correctly (and not misclassify the record as the simple layout).
    let text = "😀".to_string();
    let cch = text.encode_utf16().count() as u32;
    assert_eq!(cch, 2, "expected emoji to use a surrogate pair");
    let utf16 = utf16_le_bytes(&text);

    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&cch.to_le_bytes());
    cell_st_payload.push(0); // flags byte (flagged layout)
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

#[test]
fn parser_reads_simple_inline_string_with_trailing_bytes_low_byte_zero() {
    // Regression: simple-layout strings can start with a UTF-16 code unit whose low byte happens
    // to look like a plausible flags byte (e.g. `0x00`). When extra trailing bytes are present, the
    // reader must still decode the simple layout correctly and ignore the suffix.
    let text = "Ā".to_string(); // U+0100 => UTF-16LE bytes [0x00, 0x01]
    let cch = text.encode_utf16().count() as u32;
    let utf16 = utf16_le_bytes(&text);

    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&cch.to_le_bytes());
    cell_st_payload.extend_from_slice(&utf16);
    cell_st_payload.extend_from_slice(&[0xDE, 0xAD, 0xBE, 0xEF]); // trailing junk bytes

    let sheet_bin = sheet_with_single_cell_st(&cell_st_payload);
    let parsed = parse_sheet_bin(&mut Cursor::new(&sheet_bin), &[]).expect("parse sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("find cell");
    assert_eq!(cell.value, CellValue::Text(text));
}

#[test]
fn parser_reads_simple_inline_string_with_trailing_bytes_non_ascii() {
    // Ensure the reader does not misclassify a simple-layout string as flagged when trailing bytes
    // are present and the UTF-16 bytes do not have the common ASCII `high=0` pattern.
    let text = "😀".to_string();
    let cch = text.encode_utf16().count() as u32;
    let utf16 = utf16_le_bytes(&text);

    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&cch.to_le_bytes());
    cell_st_payload.extend_from_slice(&utf16);
    cell_st_payload.extend_from_slice(&[0xDE, 0xAD, 0xBE, 0xEF]); // trailing junk bytes

    let sheet_bin = sheet_with_single_cell_st(&cell_st_payload);
    let parsed = parse_sheet_bin(&mut Cursor::new(&sheet_bin), &[]).expect("parse sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("find cell");
    assert_eq!(cell.value, CellValue::Text(text));
}
