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

fn first_ptg_namex_ref(rgce: &[u8]) -> Option<(u16, u16)> {
    // PtgNameX token layout: [ptgNameX{R,V,A}][ixti:u16][name_index:u16]
    //
    // We intentionally ignore the token *class* (R/V/A) and focus on the `(ixti, name_index)`
    // payload, since this regression test is about stable NameX indices, not expression typing.
    //
    // We scan token-by-token (instead of searching raw byte windows) so we don't accidentally
    // match `0x39` inside unrelated token payloads (e.g. floating point constants).
    let mut i = 0usize;
    while i < rgce.len() {
        let ptg = *rgce.get(i)?;
        i += 1;

        match ptg {
            // PtgNameX: [ixti:u16][name_index:u16]
            0x39 | 0x59 | 0x79 => {
                let end = i.checked_add(4)?;
                let payload = rgce.get(i..end)?;
                let ixti = u16::from_le_bytes([payload[0], payload[1]]);
                let name_index = u16::from_le_bytes([payload[2], payload[3]]);
                return Some((ixti, name_index));
            }

            // PtgInt: [u16]
            0x1E | 0x3E | 0x5E => {
                i = i.checked_add(2)?;
                if i > rgce.len() {
                    return None;
                }
            }
            // PtgNum: [f64]
            0x1F | 0x3F | 0x5F => {
                i = i.checked_add(8)?;
                if i > rgce.len() {
                    return None;
                }
            }
            // PtgFuncVar: [argc:u8][iftab:u16]
            0x22 | 0x42 | 0x62 => {
                i = i.checked_add(3)?;
                if i > rgce.len() {
                    return None;
                }
            }

            // Unexpected token in this test formula. Bail out instead of risking desync.
            _ => return None,
        }
    }

    None
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
    // Avoid pre-allocating based on attacker-controlled ZIP metadata.
    let mut workbook_bin = Vec::new();
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

#[test]
fn save_with_cell_formula_text_edits_auto_interns_xlfn_xlws_namespaced_function() {
    let fixture_path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/simple.xlsb");

    let tmpdir = tempdir().expect("tempdir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::copy(&fixture_path, &input_path).expect("copy fixture");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");

    // Use a synthetic `_xlfn._xlws.*` function name to ensure we exercise the code path that
    // treats `_xlws.` names as BIFF UDF/namex calls (iftab=255).
    let row = 10;
    let col = 12;
    let formula_text = "=_xlfn._xlws.SOME_FUTURE_WEBSERVICE(1)";

    wb.save_with_cell_formula_text_edits(
        &output_path,
        0,
        &[FormulaTextCellEdit {
            row,
            col,
            new_value: CellValue::Number(0.0),
            formula: formula_text.to_string(),
        }],
    )
    .expect("save_with_cell_formula_text_edits");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open saved workbook");

    // WorkbookContext should resolve the newly-interned NameX function so we can re-encode.
    formula_xlsb::rgce::encode_rgce_with_context_ast(
        formula_text,
        wb2.workbook_context(),
        CellCoord::new(row, col),
    )
    .expect("re-encode xlws function using updated workbook context");

    let sheet = wb2.read_sheet(0).expect("read sheet");
    let cell = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (row, col))
        .expect("edited cell exists");

    let formula = cell.formula.as_ref().expect("formula exists");
    assert_eq!(
        formula.text.as_deref(),
        Some("_xlfn._xlws.SOME_FUTURE_WEBSERVICE(1)")
    );

    // Verify workbook.bin gained an AddIn SupBook + ExternName entry for the function.
    let file = std::fs::File::open(&output_path).expect("open output");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut wb_entry = zip.by_name("xl/workbook.bin").expect("workbook.bin");
    // Avoid pre-allocating based on attacker-controlled ZIP metadata.
    let mut workbook_bin = Vec::new();
    wb_entry
        .read_to_end(&mut workbook_bin)
        .expect("read workbook.bin");

    let mut cursor = Cursor::new(workbook_bin.as_slice());
    let mut saw_addin_supbook = false;
    let mut saw_xlws_extern_name = false;

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
                    if name == "_xlfn._xlws.SOME_FUTURE_WEBSERVICE" {
                        saw_xlws_extern_name = true;
                    }
                }
            }
            _ => {}
        }
    }

    assert!(saw_addin_supbook, "expected AddIn SupBook in workbook.bin");
    assert!(
        saw_xlws_extern_name,
        "expected ExternName for _xlfn._xlws.SOME_FUTURE_WEBSERVICE in workbook.bin"
    );
}

#[test]
fn save_with_cell_formula_text_edits_auto_interns_xlfn_xludf_namespaced_function() {
    let fixture_path =
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/simple.xlsb");

    let tmpdir = tempdir().expect("tempdir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::copy(&fixture_path, &input_path).expect("copy fixture");

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");

    // `_xludf.` names are also encoded as BIFF UDF calls (iftab=255) and therefore require a
    // NameX extern-function entry to be encodable in the workbook formula token stream.
    let row = 10;
    let col = 13;
    let formula_text = "=_xlfn._xludf.SOME_FUTURE_UDF(1)";

    wb.save_with_cell_formula_text_edits(
        &output_path,
        0,
        &[FormulaTextCellEdit {
            row,
            col,
            new_value: CellValue::Number(0.0),
            formula: formula_text.to_string(),
        }],
    )
    .expect("save_with_cell_formula_text_edits");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open saved workbook");

    // WorkbookContext should resolve the newly-interned NameX function so we can re-encode.
    formula_xlsb::rgce::encode_rgce_with_context_ast(
        formula_text,
        wb2.workbook_context(),
        CellCoord::new(row, col),
    )
    .expect("re-encode xludf function using updated workbook context");

    let sheet = wb2.read_sheet(0).expect("read sheet");
    let cell = sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (row, col))
        .expect("edited cell exists");

    let formula = cell.formula.as_ref().expect("formula exists");
    assert_eq!(
        formula.text.as_deref(),
        Some("_xlfn._xludf.SOME_FUTURE_UDF(1)")
    );

    // Verify workbook.bin gained an AddIn SupBook + ExternName entry for the function.
    let file = std::fs::File::open(&output_path).expect("open output");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut wb_entry = zip.by_name("xl/workbook.bin").expect("workbook.bin");
    // Avoid pre-allocating based on attacker-controlled ZIP metadata.
    let mut workbook_bin = Vec::new();
    wb_entry
        .read_to_end(&mut workbook_bin)
        .expect("read workbook.bin");

    let mut cursor = Cursor::new(workbook_bin.as_slice());
    let mut saw_addin_supbook = false;
    let mut saw_xludf_extern_name = false;

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
                    if name == "_xlfn._xludf.SOME_FUTURE_UDF" {
                        saw_xludf_extern_name = true;
                    }
                }
            }
            _ => {}
        }
    }

    assert!(saw_addin_supbook, "expected AddIn SupBook in workbook.bin");
    assert!(
        saw_xludf_extern_name,
        "expected ExternName for _xlfn._xludf.SOME_FUTURE_UDF in workbook.bin"
    );
}

#[test]
fn save_with_cell_formula_text_edits_reuses_existing_addin_supbook() {
    let fixture_path = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/udf.xlsb");

    let tmpdir = tempdir().expect("tempdir");
    let input_path = tmpdir.path().join("input.xlsb");
    let output_path = tmpdir.path().join("output.xlsb");
    std::fs::copy(&fixture_path, &input_path).expect("copy fixture");

    // Count how many AddIn SupBooks the fixture already has so we can ensure the patcher inserts
    // into an existing one rather than synthesizing a new entry.
    let file = std::fs::File::open(&input_path).expect("open input");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut wb_entry = zip.by_name("xl/workbook.bin").expect("workbook.bin");
    let mut workbook_bin = Vec::new();
    wb_entry
        .read_to_end(&mut workbook_bin)
        .expect("read workbook.bin");

    let mut cursor = Cursor::new(workbook_bin.as_slice());
    let mut addin_supbook_count_before = 0usize;
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

        if id == 0x00AE {
            if payload.len() < 2 {
                continue;
            }
            let mut off = 2usize;
            if let Some(raw_name) = read_xl_wide_string(&payload, &mut off) {
                if raw_name == "\u{0001}" {
                    addin_supbook_count_before += 1;
                }
            }
        }
    }
    assert!(
        addin_supbook_count_before >= 1,
        "expected udf.xlsb fixture to include an AddIn SupBook"
    );

    let wb = XlsbWorkbook::open(&input_path).expect("open workbook");

    let encoded_before = formula_xlsb::rgce::encode_rgce_with_context_ast(
        "=MyAddinFunc(1,2)",
        wb.workbook_context(),
        CellCoord::new(0, 0),
    )
    .expect("encode MyAddinFunc before patch");
    let namex_ref_before = first_ptg_namex_ref(&encoded_before.rgce)
        .expect("expected MyAddinFunc to encode using a PtgNameX token");

    // Insert a new `_xlfn.*` future function; udf.xlsb already has an AddIn SupBook + NameX table,
    // so the patcher should *reuse* it.
    let row = 10;
    let col = 10;
    let formula_text = "=_xlfn.FUTURE_FUNC_IN_EXISTING_SUPBOOK(1)";

    wb.save_with_cell_formula_text_edits(
        &output_path,
        0,
        &[FormulaTextCellEdit {
            row,
            col,
            new_value: CellValue::Number(0.0),
            formula: formula_text.to_string(),
        }],
    )
    .expect("save_with_cell_formula_text_edits");

    let wb2 = XlsbWorkbook::open(&output_path).expect("open saved workbook");

    // Existing NameX functions should still encode with the same NameX token payload (stable
    // indices), regardless of how constants are tokenized.
    let encoded_after = formula_xlsb::rgce::encode_rgce_with_context_ast(
        "=MyAddinFunc(1,2)",
        wb2.workbook_context(),
        CellCoord::new(0, 0),
    )
    .expect("encode MyAddinFunc after patch");
    let namex_ref_after = first_ptg_namex_ref(&encoded_after.rgce)
        .expect("expected MyAddinFunc to encode using a PtgNameX token");
    assert_eq!(
        namex_ref_after, namex_ref_before,
        "expected MyAddinFunc NameX reference to remain stable after patch"
    );

    // New future function should be encodable using the updated workbook context.
    formula_xlsb::rgce::encode_rgce_with_context_ast(
        formula_text,
        wb2.workbook_context(),
        CellCoord::new(row, col),
    )
    .expect("encode future function using updated workbook context");

    // Verify workbook.bin still has the same number of AddIn SupBooks, and includes the new
    // ExternName entry.
    let file = std::fs::File::open(&output_path).expect("open output");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut wb_entry = zip.by_name("xl/workbook.bin").expect("workbook.bin");
    let mut workbook_bin = Vec::new();
    wb_entry
        .read_to_end(&mut workbook_bin)
        .expect("read workbook.bin");

    let mut cursor = Cursor::new(workbook_bin.as_slice());
    let mut addin_supbook_count_after = 0usize;
    let mut saw_my_addin_func = false;
    let mut saw_new_future_func = false;

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
                if payload.len() < 2 {
                    continue;
                }
                let mut off = 2usize;
                if let Some(raw_name) = read_xl_wide_string(&payload, &mut off) {
                    if raw_name == "\u{0001}" {
                        addin_supbook_count_after += 1;
                    }
                }
            }
            0x0023 | 0x0168 => {
                if payload.len() < 4 {
                    continue;
                }
                let mut off = 4usize;
                if let Some(name) = read_xl_wide_string(&payload, &mut off) {
                    if name.eq_ignore_ascii_case("MyAddinFunc") {
                        saw_my_addin_func = true;
                    }
                    if name.eq_ignore_ascii_case("_xlfn.FUTURE_FUNC_IN_EXISTING_SUPBOOK") {
                        saw_new_future_func = true;
                    }
                }
            }
            _ => {}
        }
    }

    assert_eq!(
        addin_supbook_count_after, addin_supbook_count_before,
        "expected save_with_cell_formula_text_edits to reuse existing AddIn SupBook"
    );
    assert!(
        saw_my_addin_func,
        "expected workbook.bin to retain existing MyAddinFunc extern name"
    );
    assert!(
        saw_new_future_func,
        "expected workbook.bin to include ExternName for _xlfn.FUTURE_FUNC_IN_EXISTING_SUPBOOK"
    );
}
