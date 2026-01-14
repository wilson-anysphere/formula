use formula_xlsb::{patch_sheet_bin, patch_sheet_bin_streaming, CellEdit, CellValue};
use pretty_assertions::assert_eq;
use std::io::Cursor;

fn encode_biff12_id(id: u32) -> Vec<u8> {
    // BIFF12 record ids use the same 7-bit varint encoding as record lengths.
    let mut out = Vec::new();
    let mut value = id;
    loop {
        let mut byte = (value & 0x7F) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        out.push(byte);
        if value == 0 {
            break;
        }
    }
    out
}

fn encode_biff12_len(mut len: u32) -> Vec<u8> {
    let mut out = Vec::new();
    loop {
        let mut byte = (len & 0x7F) as u8;
        len >>= 7;
        if len != 0 {
            byte |= 0x80;
        }
        out.push(byte);
        if len == 0 {
            break;
        }
    }
    out
}

fn biff12_record(id: u32, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&encode_biff12_id(id));
    out.extend_from_slice(&encode_biff12_len(payload.len() as u32));
    out.extend_from_slice(payload);
    out
}

fn encode_xl_wide_string(
    s: &str,
    flags: u16,
    flags_width: usize,
    rich_runs: Option<&[u8]>,
    phonetic: Option<&[u8]>,
) -> Vec<u8> {
    let units: Vec<u16> = s.encode_utf16().collect();
    let mut out = Vec::new();
    out.extend_from_slice(&(units.len() as u32).to_le_bytes());
    match flags_width {
        1 => out.push(flags as u8),
        2 => out.extend_from_slice(&flags.to_le_bytes()),
        other => panic!("unexpected flags width {other}"),
    }
    for u in units {
        out.extend_from_slice(&u.to_le_bytes());
    }

    if flags & 0x0001 != 0 {
        let rich = rich_runs.expect("rich flag requires runs");
        assert_eq!(rich.len() % 8, 0, "rich run bytes must be multiple of 8");
        out.extend_from_slice(&((rich.len() / 8) as u32).to_le_bytes());
        out.extend_from_slice(rich);
    }

    if flags & 0x0002 != 0 {
        let pho = phonetic.expect("phonetic flag requires bytes");
        out.extend_from_slice(&(pho.len() as u32).to_le_bytes());
        out.extend_from_slice(pho);
    }

    out
}

fn read_varint(data: &[u8], offset: &mut usize) -> u32 {
    let mut v: u32 = 0;
    for i in 0..4 {
        let byte = *data.get(*offset).expect("varint byte");
        *offset += 1;
        v |= ((byte & 0x7F) as u32) << (7 * i);
        if byte & 0x80 == 0 {
            return v;
        }
    }
    panic!("invalid BIFF12 varint (more than 4 bytes)");
}

fn find_record_payload<'a>(stream: &'a [u8], desired_id: u32) -> &'a [u8] {
    let mut offset = 0usize;
    while offset < stream.len() {
        let id = read_varint(stream, &mut offset);
        let len = read_varint(stream, &mut offset) as usize;
        let payload_end = offset + len;
        let payload = stream.get(offset..payload_end).expect("record payload");
        offset = payload_end;
        if id == desired_id {
            return payload;
        }
    }
    panic!("record id {desired_id} not found");
}

fn build_sheet_bin_with_cell_st(cell_st_payload: &[u8]) -> Vec<u8> {
    // Record ids (subset):
    // - BrtWsDim          0x0094
    // - BrtBeginSheetData 0x0091
    // - BrtEndSheetData   0x0092
    // - BrtRow            0x0000
    // - BrtCellSt         0x0006
    const DIMENSION: u32 = 0x0094;
    const SHEETDATA: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const ROW: u32 = 0x0000;
    const CELL_ST: u32 = 0x0006;

    let mut sheet_bin = Vec::new();
    // Provide DIMENSION up-front so the streaming patcher can operate without falling back to the
    // in-memory patcher (it requires DIMENSION to appear before SHEETDATA).
    //
    // BrtWsDim payload: [r1:u32][r2:u32][c1:u32][c2:u32] (inclusive bounds in XLSB).
    let dim_payload = [
        0u32.to_le_bytes(),
        0u32.to_le_bytes(),
        0u32.to_le_bytes(),
        0u32.to_le_bytes(),
    ]
    .concat();
    sheet_bin.extend_from_slice(&biff12_record(DIMENSION, &dim_payload));
    sheet_bin.extend_from_slice(&biff12_record(SHEETDATA, &[]));
    sheet_bin.extend_from_slice(&biff12_record(ROW, &0u32.to_le_bytes()));
    sheet_bin.extend_from_slice(&biff12_record(CELL_ST, cell_st_payload));
    sheet_bin.extend_from_slice(&biff12_record(SHEETDATA_END, &[]));
    sheet_bin
}

#[test]
fn patcher_preserves_flagged_inline_string_layout_on_value_update() {
    const CELL_ST: u32 = 0x0006;

    let original_text = "Hello".to_string();
    let wide = encode_xl_wide_string(&original_text, 0, 1, None, None);
    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&wide);
    let sheet_bin = build_sheet_bin_with_cell_st(&cell_st_payload);

    let new_text = "New".to_string();
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

    // Ensure the sheet parses and the value changed.
    let parsed =
        formula_xlsb::parse_sheet_bin(&mut Cursor::new(&patched), &[]).expect("parse sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("find cell");
    assert_eq!(cell.value, CellValue::Text(new_text));

    // Ensure the BrtCellSt payload still uses the flagged wide-string layout:
    //   [col][style][cch][flags:u8][utf16...]
    let payload = find_record_payload(&patched, CELL_ST);
    let cch = u32::from_le_bytes(payload[8..12].try_into().expect("cch bytes")) as usize;
    let expected_simple_len = 12 + cch * 2;
    assert_ne!(
        payload.len(),
        expected_simple_len,
        "expected patched inline string to preserve the flagged layout (flags byte present)"
    );
    assert_eq!(payload[12], 0, "expected flags byte to be preserved");
}

#[test]
fn patcher_preserves_inline_string_flags_and_emits_empty_rich_phonetic_blocks() {
    const CELL_ST: u32 = 0x0006;

    // Make the rich/phonetic bytes distinctive to avoid accidental matches.
    let rich_runs: Vec<u8> = vec![
        0xDE, 0xAD, 0xBE, 0xEF, 0x10, 0x11, 0x12, 0x13, 0xFE, 0xED, 0xFA, 0xCE, 0x20, 0x21,
        0x22, 0x23,
    ];
    let phonetic_bytes: Vec<u8> = vec![0xA1, 0xB2, 0xC3, 0xD4, 0xE5, 0xF6, 0x07];

    let original_text = "RichPho".to_string();
    let wide = encode_xl_wide_string(
        &original_text,
        0x0003,
        1,
        Some(&rich_runs),
        Some(&phonetic_bytes),
    );
    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&wide);
    let sheet_bin = build_sheet_bin_with_cell_st(&cell_st_payload);

    let new_text = "New".to_string();
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

    // The original rich/phonetic bytes should not remain if the string was rewritten.
    assert!(
        !patched
            .windows(rich_runs.len())
            .any(|w| w == rich_runs.as_slice()),
        "expected rich run bytes to be dropped when inline string value changes"
    );
    assert!(
        !patched
            .windows(phonetic_bytes.len())
            .any(|w| w == phonetic_bytes.as_slice()),
        "expected phonetic bytes to be dropped when inline string value changes"
    );

    let parsed =
        formula_xlsb::parse_sheet_bin(&mut Cursor::new(&patched), &[]).expect("parse sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("find cell");
    assert_eq!(cell.value, CellValue::Text(new_text));

    let payload = find_record_payload(&patched, CELL_ST);
    assert_eq!(payload[12], 0x03, "expected flags byte to be preserved");

    let cch = u32::from_le_bytes(payload[8..12].try_into().expect("cch bytes")) as usize;
    let utf16_end = 13 + cch * 2;
    let c_run_offset = utf16_end;
    let cb_offset = c_run_offset + 4;

    let c_run = u32::from_le_bytes(
        payload[c_run_offset..c_run_offset + 4]
            .try_into()
            .expect("cRun bytes"),
    );
    let cb = u32::from_le_bytes(
        payload[cb_offset..cb_offset + 4]
            .try_into()
            .expect("cb bytes"),
    );
    assert_eq!(c_run, 0, "expected empty rich text run payload");
    assert_eq!(cb, 0, "expected empty phonetic payload");

    let expected_len = 13 + cch * 2 + 8;
    assert_eq!(
        payload.len(),
        expected_len,
        "expected patched inline string to include only empty rich/phonetic blocks"
    );
}

#[test]
fn patcher_preserves_inline_string_extras_on_style_update_when_text_unchanged() {
    const CELL_ST: u32 = 0x0006;

    // Make the rich/phonetic bytes distinctive to avoid accidental matches.
    let rich_runs: Vec<u8> = vec![
        0xDE, 0xAD, 0xBE, 0xEF, 0x10, 0x11, 0x12, 0x13, 0xFE, 0xED, 0xFA, 0xCE, 0x20, 0x21,
        0x22, 0x23,
    ];
    let phonetic_bytes: Vec<u8> = vec![0xA1, 0xB2, 0xC3, 0xD4, 0xE5, 0xF6, 0x07];

    let original_text = "RichPho".to_string();
    let wide = encode_xl_wide_string(
        &original_text,
        0x0003,
        1,
        Some(&rich_runs),
        Some(&phonetic_bytes),
    );
    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&wide);
    let sheet_bin = build_sheet_bin_with_cell_st(&cell_st_payload);

    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text(original_text.clone()),
        new_style: Some(7),
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    };
    let patched = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");

    // The original rich/phonetic bytes should be preserved when the string did not change.
    assert!(
        patched
            .windows(rich_runs.len())
            .any(|w| w == rich_runs.as_slice()),
        "expected rich run bytes to be preserved on style-only update"
    );
    assert!(
        patched
            .windows(phonetic_bytes.len())
            .any(|w| w == phonetic_bytes.as_slice()),
        "expected phonetic bytes to be preserved on style-only update"
    );

    let parsed =
        formula_xlsb::parse_sheet_bin(&mut Cursor::new(&patched), &[]).expect("parse sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("find cell");
    assert_eq!(cell.value, CellValue::Text(original_text));
    assert_eq!(cell.style, 7);

    // Validate the rich/phonetic headers are still present and match the original bytes.
    let payload = find_record_payload(&patched, CELL_ST);
    assert_eq!(payload[12], 0x03, "expected flags byte to be preserved");
    let cch = u32::from_le_bytes(payload[8..12].try_into().expect("cch bytes")) as usize;
    let utf16_end = 13 + cch * 2;
    let c_run = u32::from_le_bytes(
        payload[utf16_end..utf16_end + 4]
            .try_into()
            .expect("cRun bytes"),
    ) as usize;
    assert_eq!(c_run * 8, rich_runs.len());
    let rich_start = utf16_end + 4;
    let rich_end = rich_start + rich_runs.len();
    assert_eq!(&payload[rich_start..rich_end], rich_runs.as_slice());

    let cb_offset = rich_end;
    let cb = u32::from_le_bytes(
        payload[cb_offset..cb_offset + 4]
            .try_into()
            .expect("cb bytes"),
    ) as usize;
    assert_eq!(cb, phonetic_bytes.len());
    let pho_start = cb_offset + 4;
    let pho_end = pho_start + phonetic_bytes.len();
    assert_eq!(&payload[pho_start..pho_end], phonetic_bytes.as_slice());
}

#[test]
fn streaming_patcher_preserves_flagged_inline_string_layout_on_value_update() {
    const CELL_ST: u32 = 0x0006;

    let original_text = "Hello".to_string();
    let wide = encode_xl_wide_string(&original_text, 0, 1, None, None);
    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&wide);
    let sheet_bin = build_sheet_bin_with_cell_st(&cell_st_payload);

    let new_text = "New".to_string();
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

    let mut out = Vec::new();
    let changed =
        patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut out, &[edit]).expect("patch");
    assert!(changed, "expected streaming patcher to report changes");

    let parsed =
        formula_xlsb::parse_sheet_bin(&mut Cursor::new(&out), &[]).expect("parse sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("find cell");
    assert_eq!(cell.value, CellValue::Text(new_text));

    let payload = find_record_payload(&out, CELL_ST);
    let cch = u32::from_le_bytes(payload[8..12].try_into().expect("cch bytes")) as usize;
    let expected_simple_len = 12 + cch * 2;
    assert_ne!(payload.len(), expected_simple_len);
    assert_eq!(payload[12], 0);
}

#[test]
fn streaming_patcher_preserves_inline_string_extras_on_style_update_when_text_unchanged() {
    const CELL_ST: u32 = 0x0006;

    // Make the rich/phonetic bytes distinctive to avoid accidental matches.
    let rich_runs: Vec<u8> = vec![
        0xDE, 0xAD, 0xBE, 0xEF, 0x10, 0x11, 0x12, 0x13, 0xFE, 0xED, 0xFA, 0xCE, 0x20, 0x21,
        0x22, 0x23,
    ];
    let phonetic_bytes: Vec<u8> = vec![0xA1, 0xB2, 0xC3, 0xD4, 0xE5, 0xF6, 0x07];

    let original_text = "RichPho".to_string();
    let wide = encode_xl_wide_string(
        &original_text,
        0x0003,
        1,
        Some(&rich_runs),
        Some(&phonetic_bytes),
    );
    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&wide);
    let sheet_bin = build_sheet_bin_with_cell_st(&cell_st_payload);

    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text(original_text.clone()),
        new_style: Some(7),
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    };

    let mut out = Vec::new();
    let changed =
        patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut out, &[edit]).expect("patch");
    assert!(changed, "expected streaming patcher to report changes");

    // The original rich/phonetic bytes should be preserved when the string did not change.
    assert!(
        out.windows(rich_runs.len()).any(|w| w == rich_runs.as_slice()),
        "expected rich run bytes to be preserved on style-only update"
    );
    assert!(
        out.windows(phonetic_bytes.len())
            .any(|w| w == phonetic_bytes.as_slice()),
        "expected phonetic bytes to be preserved on style-only update"
    );

    let parsed = formula_xlsb::parse_sheet_bin(&mut Cursor::new(&out), &[]).expect("parse sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 0)
        .expect("find cell");
    assert_eq!(cell.value, CellValue::Text(original_text));
    assert_eq!(cell.style, 7);

    // Validate the rich/phonetic headers are still present and match the original bytes.
    let payload = find_record_payload(&out, CELL_ST);
    assert_eq!(payload[12], 0x03, "expected flags byte to be preserved");
    let cch = u32::from_le_bytes(payload[8..12].try_into().expect("cch bytes")) as usize;
    let utf16_end = 13 + cch * 2;
    let c_run = u32::from_le_bytes(
        payload[utf16_end..utf16_end + 4]
            .try_into()
            .expect("cRun bytes"),
    ) as usize;
    assert_eq!(c_run * 8, rich_runs.len());
    let rich_start = utf16_end + 4;
    let rich_end = rich_start + rich_runs.len();
    assert_eq!(&payload[rich_start..rich_end], rich_runs.as_slice());

    let cb_offset = rich_end;
    let cb = u32::from_le_bytes(
        payload[cb_offset..cb_offset + 4]
            .try_into()
            .expect("cb bytes"),
    ) as usize;
    assert_eq!(cb, phonetic_bytes.len());
    let pho_start = cb_offset + 4;
    let pho_end = pho_start + phonetic_bytes.len();
    assert_eq!(&payload[pho_start..pho_end], phonetic_bytes.as_slice());
}
