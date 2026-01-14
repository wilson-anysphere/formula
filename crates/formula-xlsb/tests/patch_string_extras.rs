use formula_xlsb::{patch_sheet_bin, CellEdit, CellValue};
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

fn rewrite_formula_string_headers_as_two_byte_varints(sheet_bin: &[u8]) -> Vec<u8> {
    const FORMULA_STRING: u32 = 0x0008;

    let mut offset = 0usize;
    let mut out = Vec::with_capacity(sheet_bin.len() + 16);

    while offset < sheet_bin.len() {
        let record_start = offset;
        let id = read_varint(sheet_bin, &mut offset);
        let _id_end = offset;
        let len = read_varint(sheet_bin, &mut offset) as usize;
        let len_end = offset;
        let payload_start = len_end;
        let payload_end = payload_start + len;
        let payload = sheet_bin
            .get(payload_start..payload_end)
            .expect("record payload");
        offset = payload_end;

        if id == FORMULA_STRING {
            assert!(
                id < 0x80 && len < 0x80,
                "test helper only supports 1-byte varints"
            );
            // Non-canonical, but valid, 2-byte LEB128 varints for values < 128.
            out.extend_from_slice(&[(id as u8) | 0x80, 0x00]);
            out.extend_from_slice(&[(len as u8) | 0x80, 0x00]);
            out.extend_from_slice(payload);
        } else {
            out.extend_from_slice(
                sheet_bin
                    .get(record_start..payload_end)
                    .expect("record bytes"),
            );
        }
    }

    out
}

fn formula_string_header_raw(sheet_bin: &[u8], desired_col: u32) -> (Vec<u8>, Vec<u8>, usize) {
    const FORMULA_STRING: u32 = 0x0008;

    let mut offset = 0usize;
    while offset < sheet_bin.len() {
        let record_start = offset;
        let id = read_varint(sheet_bin, &mut offset);
        let id_end = offset;
        let len = read_varint(sheet_bin, &mut offset) as usize;
        let len_end = offset;
        let payload_start = len_end;
        let payload_end = payload_start + len;
        let payload = sheet_bin
            .get(payload_start..payload_end)
            .expect("record payload");
        offset = payload_end;

        if id == FORMULA_STRING {
            let col = u32::from_le_bytes(payload[0..4].try_into().expect("col bytes"));
            if col == desired_col {
                return (
                    sheet_bin[record_start..id_end].to_vec(),
                    sheet_bin[id_end..len_end].to_vec(),
                    len,
                );
            }
        }
    }

    panic!("FORMULA_STRING record for col {desired_col} not found");
}

fn build_sheet_bin() -> (Vec<u8>, String, Vec<u8>, String, Vec<u8>, String, Vec<u8>) {
    // Record ids (subset):
    // - BrtBeginSheetData 0x0091
    // - BrtEndSheetData   0x0092
    // - BrtRow            0x0000
    // - BrtCellSt         0x0006
    // - BrtFmlaString     0x0008
    const SHEETDATA: u32 = 0x0091;
    const SHEETDATA_END: u32 = 0x0092;
    const ROW: u32 = 0x0000;
    const CELL_ST: u32 = 0x0006;
    const FORMULA_STRING: u32 = 0x0008;

    let mut sheet_bin = Vec::new();
    sheet_bin.extend_from_slice(&biff12_record(SHEETDATA, &[]));
    sheet_bin.extend_from_slice(&biff12_record(ROW, &0u32.to_le_bytes()));

    // A1: inline string with flags byte present (flags=0).
    let inline_text = r#"He said "Hi""#.to_string();
    let inline_wide = encode_xl_wide_string(&inline_text, 0, 1, None, None);
    let mut cell_st_payload = Vec::new();
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // col
    cell_st_payload.extend_from_slice(&0u32.to_le_bytes()); // style
    cell_st_payload.extend_from_slice(&inline_wide);
    sheet_bin.extend_from_slice(&biff12_record(CELL_ST, &cell_st_payload));

    // B1: formula cached string with rich-text runs.
    let rich_runs: Vec<u8> = vec![0x10, 0x11, 0x12, 0x13, 0x20, 0x21, 0x22, 0x23];
    let rich_text = "Rich".to_string();
    let rich_wide = encode_xl_wide_string(&rich_text, 0x0001, 2, Some(&rich_runs), None);
    let mut fmla_rich = Vec::new();
    fmla_rich.extend_from_slice(&1u32.to_le_bytes()); // col
    fmla_rich.extend_from_slice(&0u32.to_le_bytes()); // style
    fmla_rich.extend_from_slice(&rich_wide);
    fmla_rich.extend_from_slice(&2u32.to_le_bytes()); // cce
    fmla_rich.extend_from_slice(&[0x1D, 0x01]); // rgce: PtgBool TRUE
    sheet_bin.extend_from_slice(&biff12_record(FORMULA_STRING, &fmla_rich));

    // C1: formula cached string with phonetic/extended bytes.
    let phonetic_bytes: Vec<u8> = vec![0x90, 0x91, 0x92, 0x93, 0x94];
    let pho_text = "Pho".to_string();
    let pho_wide = encode_xl_wide_string(&pho_text, 0x0002, 2, None, Some(&phonetic_bytes));
    let mut fmla_pho = Vec::new();
    fmla_pho.extend_from_slice(&2u32.to_le_bytes()); // col
    fmla_pho.extend_from_slice(&0u32.to_le_bytes()); // style
    fmla_pho.extend_from_slice(&pho_wide);
    fmla_pho.extend_from_slice(&2u32.to_le_bytes()); // cce
    fmla_pho.extend_from_slice(&[0x1D, 0x01]); // rgce: PtgBool TRUE
    sheet_bin.extend_from_slice(&biff12_record(FORMULA_STRING, &fmla_pho));

    sheet_bin.extend_from_slice(&biff12_record(SHEETDATA_END, &[]));

    (
        sheet_bin,
        inline_text,
        rich_runs,
        rich_text,
        phonetic_bytes,
        pho_text,
        vec![0x1D, 0x00], // rgce: PtgBool FALSE
    )
}

#[test]
fn patcher_preserves_flagged_inline_string_record_for_noop_edits() {
    let (sheet_bin, inline_text, _rich_runs, _rich_text, _pho_bytes, _pho_text, _new_rgce) =
        build_sheet_bin();

    let edit = CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text(inline_text),
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    };

    let patched = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");
    assert_eq!(patched, sheet_bin);
}

#[test]
fn patcher_updates_formula_rgce_without_losing_rich_or_phonetic_cached_bytes() {
    let (sheet_bin, _inline_text, rich_runs, rich_text, phonetic_bytes, pho_text, new_rgce) =
        build_sheet_bin();
    let tweaked = rewrite_formula_string_headers_as_two_byte_varints(&sheet_bin);

    let (rich_id_raw, rich_len_raw, rich_len) = formula_string_header_raw(&tweaked, 1);
    assert_eq!(
        rich_id_raw,
        vec![0x88, 0x00],
        "expected non-canonical id varint for BrtFmlaString"
    );
    assert_eq!(
        rich_len_raw,
        vec![(rich_len as u8) | 0x80, 0x00],
        "expected non-canonical len varint for rich formula payload"
    );

    let (pho_id_raw, pho_len_raw, pho_len) = formula_string_header_raw(&tweaked, 2);
    assert_eq!(
        pho_id_raw,
        vec![0x88, 0x00],
        "expected non-canonical id varint for BrtFmlaString"
    );
    assert_eq!(
        pho_len_raw,
        vec![(pho_len as u8) | 0x80, 0x00],
        "expected non-canonical len varint for phonetic formula payload"
    );

    let edit_rich = CellEdit {
        row: 0,
        col: 1,
        new_value: CellValue::Text(rich_text.clone()),
        clear_formula: false,
        new_formula: Some(new_rgce.clone()),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    };
    let patched_rich = patch_sheet_bin(&tweaked, &[edit_rich]).expect("patch rich formula");

    let (patched_id_raw, patched_len_raw, patched_len) =
        formula_string_header_raw(&patched_rich, 1);
    assert_eq!(patched_id_raw, rich_id_raw);
    assert_eq!(patched_len_raw, rich_len_raw);
    assert_eq!(patched_len, rich_len);

    // Ensure the cached rich-run bytes were preserved verbatim.
    assert!(
        patched_rich
            .windows(rich_runs.len())
            .any(|w| w == rich_runs.as_slice()),
        "expected patched output to still contain rich run bytes"
    );

    // Ensure the sheet parses and the formula token stream was updated.
    let parsed =
        formula_xlsb::parse_sheet_bin(&mut Cursor::new(&patched_rich), &[]).expect("parse sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("find patched rich cell");
    assert_eq!(cell.value, CellValue::Text(rich_text));
    assert_eq!(cell.formula.as_ref().unwrap().rgce, new_rgce);

    let edit_pho = CellEdit {
        row: 0,
        col: 2,
        new_value: CellValue::Text(pho_text.clone()),
        clear_formula: false,
        new_formula: Some(new_rgce.clone()),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    };
    let patched_pho = patch_sheet_bin(&tweaked, &[edit_pho]).expect("patch phonetic formula");

    let (patched_id_raw, patched_len_raw, patched_len) = formula_string_header_raw(&patched_pho, 2);
    assert_eq!(patched_id_raw, pho_id_raw);
    assert_eq!(patched_len_raw, pho_len_raw);
    assert_eq!(patched_len, pho_len);

    assert!(
        patched_pho
            .windows(phonetic_bytes.len())
            .any(|w| w == phonetic_bytes.as_slice()),
        "expected patched output to still contain phonetic bytes"
    );

    let parsed =
        formula_xlsb::parse_sheet_bin(&mut Cursor::new(&patched_pho), &[]).expect("parse sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 2)
        .expect("find patched phonetic cell");
    assert_eq!(cell.value, CellValue::Text(pho_text));
    assert_eq!(cell.formula.as_ref().unwrap().rgce, new_rgce);
}

#[test]
fn patcher_clears_rich_phonetic_flags_when_rewriting_cached_string_value() {
    let (sheet_bin, _inline_text, rich_runs, rich_text, _pho_bytes, _pho_text, _new_rgce) =
        build_sheet_bin();

    let edit = CellEdit {
        row: 0,
        col: 1,
        new_value: CellValue::Text("New".to_string()),
        clear_formula: false,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    };
    let patched = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch cached string");

    // Cached formatting bytes should not remain if we rewrote the string value.
    assert!(
        !patched
            .windows(rich_runs.len())
            .any(|w| w == rich_runs.as_slice()),
        "expected rich run bytes to be dropped when cached value changes"
    );

    let parsed =
        formula_xlsb::parse_sheet_bin(&mut Cursor::new(&patched), &[]).expect("parse sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("find patched cell");
    assert_eq!(cell.value, CellValue::Text("New".to_string()));
    // Formula bytes remain unchanged, but we should still have a formula payload.
    assert_eq!(cell.formula.as_ref().unwrap().rgce, vec![0x1D, 0x01]);
    // sanity: original cached text was different.
    assert_ne!(rich_text, "New");
}

#[test]
fn patcher_can_convert_formula_string_cell_to_plain_text_cell() {
    let (sheet_bin, _inline_text, _rich_runs, rich_text, _pho_bytes, _pho_text, _new_rgce) =
        build_sheet_bin();

    let edit = CellEdit {
        row: 0,
        col: 1,
        new_value: CellValue::Text(rich_text.clone()),
        clear_formula: true,
        new_formula: None,
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
        new_style: None,
    };
    let patched = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet");

    let parsed =
        formula_xlsb::parse_sheet_bin(&mut Cursor::new(&patched), &[]).expect("parse sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| c.row == 0 && c.col == 1)
        .expect("find patched cell");
    assert_eq!(cell.value, CellValue::Text(rich_text));
    assert!(
        cell.formula.is_none(),
        "expected formula metadata to be removed when clear_formula=true"
    );
}
