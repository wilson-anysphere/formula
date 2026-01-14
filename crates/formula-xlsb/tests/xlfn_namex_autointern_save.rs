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
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut workbook_bin = Vec::new();
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

#[test]
fn save_with_cell_formula_text_edits_interns_multiple_xlfn_functions_without_duplicates() {
    let fixture_path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/simple.xlsb");

    let tmpdir = tempdir().expect("tempdir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::copy(&fixture_path, &input_path).expect("copy fixture");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");

    // Use multiple edits that reference the same future function (to exercise deduping) and a
    // second distinct future function (to ensure we intern multiple NameX entries in one pass).
    let row1 = 10;
    let col1 = 10;
    let row2 = 10;
    let col2 = 11;

    let edits = [
        FormulaTextCellEdit {
            row: row1,
            col: col1,
            new_value: CellValue::Number(0.0),
            formula: "=_xlfn.SOME_FUTURE_FUNCTION(1)".to_string(),
        },
        FormulaTextCellEdit {
            row: row2,
            col: col2,
            new_value: CellValue::Number(0.0),
            formula: "=_xlfn.ANOTHER_FUTURE_FUNCTION(2)+_xlfn.SOME_FUTURE_FUNCTION(3)".to_string(),
        },
    ];

    wb.save_with_cell_formula_text_edits(&output_path, 0, &edits)
        .expect("save_with_cell_formula_text_edits");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open saved workbook");

    // WorkbookContext should resolve the newly-interned NameX functions so we can re-encode.
    for edit in &edits {
        formula_xlsb::rgce::encode_rgce_with_context_ast(
            &edit.formula,
            wb2.workbook_context(),
            CellCoord::new(edit.row, edit.col),
        )
        .expect("re-encode future function using updated workbook context");
    }

    // Verify both formulas round-trip through the reader (decoded formula text is stored without
    // the leading `=`).
    let sheet = wb2.read_sheet(0).expect("read sheet");
    let cell1 = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (row1, col1))
        .expect("edited cell1 exists");
    let cell2 = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (row2, col2))
        .expect("edited cell2 exists");

    let f1 = cell1.formula.as_ref().expect("cell1 formula");
    let f2 = cell2.formula.as_ref().expect("cell2 formula");
    let text1 = f1.text.as_deref().unwrap_or_default().to_ascii_uppercase();
    let text2 = f2.text.as_deref().unwrap_or_default().to_ascii_uppercase();
    assert!(
        text1.contains("SOME_FUTURE_FUNCTION"),
        "expected cell1 decoded formula text to reference SOME_FUTURE_FUNCTION, got {:?}",
        f1.text
    );
    assert!(
        text2.contains("SOME_FUTURE_FUNCTION") && text2.contains("ANOTHER_FUTURE_FUNCTION"),
        "expected cell2 decoded formula text to reference both functions, got {:?}",
        f2.text
    );

    // Verify workbook.bin gained exactly one ExternName entry per function.
    let file = std::fs::File::open(&output_path).expect("open output");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut wb_entry = zip.by_name("xl/workbook.bin").expect("workbook.bin");
    let mut workbook_bin = Vec::with_capacity(wb_entry.size() as usize);
    wb_entry
        .read_to_end(&mut workbook_bin)
        .expect("read workbook.bin");

    let mut cursor = Cursor::new(workbook_bin.as_slice());
    let mut addin_supbook_count = 0usize;
    let mut some_future_count = 0usize;
    let mut another_future_count = 0usize;

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
                        addin_supbook_count += 1;
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
                    if name.eq_ignore_ascii_case("_xlfn.SOME_FUTURE_FUNCTION") {
                        some_future_count += 1;
                    }
                    if name.eq_ignore_ascii_case("_xlfn.ANOTHER_FUTURE_FUNCTION") {
                        another_future_count += 1;
                    }
                }
            }
            _ => {}
        }
    }

    assert!(
        addin_supbook_count >= 1,
        "expected AddIn SupBook in workbook.bin"
    );
    assert_eq!(
        some_future_count, 1,
        "expected exactly one ExternName for _xlfn.SOME_FUTURE_FUNCTION"
    );
    assert_eq!(
        another_future_count, 1,
        "expected exactly one ExternName for _xlfn.ANOTHER_FUTURE_FUNCTION"
    );
}
