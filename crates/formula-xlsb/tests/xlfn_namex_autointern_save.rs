#![cfg(feature = "write")]

use std::io::{Cursor, Read};
use std::path::PathBuf;

use formula_xlsb::biff12_varint;
use formula_xlsb::rgce::CellCoord;
use formula_xlsb::{CellValue, FormulaTextCellEdit, XlsbWorkbook};
use tempfile::tempdir;

fn read_xl_wide_string(payload: &[u8], offset: &mut usize) -> Option<String> {
    if *offset + 4 > payload.len() {
        return None;
    }
    let cch = u32::from_le_bytes(payload[*offset..*offset + 4].try_into().ok()?) as usize;
    *offset += 4;
    let byte_len = cch.checked_mul(2)?;
    if *offset + byte_len > payload.len() {
        return None;
    }
    let bytes = &payload[*offset..*offset + byte_len];
    *offset += byte_len;
    let mut units = Vec::with_capacity(cch);
    for chunk in bytes.chunks_exact(2) {
        units.push(u16::from_le_bytes([chunk[0], chunk[1]]));
    }
    String::from_utf16(&units).ok()
}

#[test]
fn save_with_cell_formula_text_edits_auto_interns_missing_xlfn_namex_function() {
    let fixture_path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/simple.xlsb");

    let tmpdir = tempdir().expect("tempdir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::copy(&fixture_path, &input_path).expect("copy fixture");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");

    // Pick a cell that does not already exist in the fixture so the patcher inserts a new formula
    // record rather than attempting to convert an existing value cell into a formula cell.
    let row = 10;
    let col = 10;

    wb.save_with_cell_formula_text_edits(
        &output_path,
        0,
        &[FormulaTextCellEdit {
            row,
            col,
            new_value: CellValue::Number(0.0),
            formula: "=_xlfn.SOME_FUTURE_FUNCTION(1,2)".to_string(),
        }],
    )
    .expect("save_with_cell_formula_text_edits");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open saved workbook");

    // WorkbookContext should resolve the newly-interned NameX function so we can re-encode.
    formula_xlsb::rgce::encode_rgce_with_context_ast(
        "=_xlfn.SOME_FUTURE_FUNCTION(1,2)",
        wb2.workbook_context(),
        CellCoord::new(row, col),
    )
    .expect("re-encode future function using updated workbook context");

    let sheet = wb2.read_sheet(0).expect("read sheet");
    let cell = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (row, col))
        .expect("edited cell exists");

    let formula = cell.formula.as_ref().expect("A1 formula");
    assert_eq!(
        formula.text.as_deref(),
        Some("_xlfn.SOME_FUTURE_FUNCTION(1,2)")
    );

    // Verify workbook.bin gained an AddIn SupBook + ExternName entry for the function.
    let file = std::fs::File::open(&output_path).expect("open output");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut wb_entry = zip.by_name("xl/workbook.bin").expect("workbook.bin");
    let mut workbook_bin = Vec::with_capacity(wb_entry.size() as usize);
    wb_entry
        .read_to_end(&mut workbook_bin)
        .expect("read workbook.bin");

    let mut cursor = Cursor::new(workbook_bin.as_slice());
    let mut saw_addin_supbook = false;
    let mut saw_future_extern_name = false;

    loop {
        let Some(id) = biff12_varint::read_record_id(&mut cursor)
            .expect("read record id")
        else {
            break;
        };
        let Some(len) = biff12_varint::read_record_len(&mut cursor)
            .expect("read record len")
        else {
            break;
        };
        let mut payload = vec![0u8; len as usize];
        cursor.read_exact(&mut payload).expect("read payload");

        match id {
            0x00AE => {
                // SupBook: [ctab:u16][raw_name: xlWideString]
                if payload.len() < 2 {
                    continue;
                }
                let mut off = 2usize;
                if let Some(raw_name) = read_xl_wide_string(&payload, &mut off) {
                    if raw_name == "\u{0001}" {
                        saw_addin_supbook = true;
                    }
                }
            }
            0x0023 | 0x0168 => {
                // ExternName layout A: [flags:u16][scope:u16][name: xlWideString]
                if payload.len() < 4 {
                    continue;
                }
                let mut off = 4usize;
                if let Some(name) = read_xl_wide_string(&payload, &mut off) {
                    if name == "_xlfn.SOME_FUTURE_FUNCTION" {
                        saw_future_extern_name = true;
                    }
                }
            }
            _ => {}
        }
    }

    assert!(saw_addin_supbook, "expected AddIn SupBook in workbook.bin");
    assert!(
        saw_future_extern_name,
        "expected ExternName for _xlfn.SOME_FUTURE_FUNCTION in workbook.bin"
    );
}
