use std::io;
use std::io::{Cursor, Read};

use formula_xlsb::{
    patch_sheet_bin, patch_sheet_bin_streaming, rgce_references_rgcb, CellEdit, CellValue, Error,
};

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

fn rgce_memfunc_with_array() -> Vec<u8> {
    // PtgMemFunc: [ptg=0x29][cce: u16][subexpression bytes...]
    //
    // Use a nested subexpression that contains `PtgArray` so `rgce_references_rgcb` must scan
    // through the mem payload (instead of bailing on an unknown token).
    let subexpr = fixture_builder::rgce::array_placeholder();

    let mut rgce = vec![0x29];
    rgce.extend_from_slice(
        &u16::try_from(subexpr.len())
            .expect("subexpression length fits in u16")
            .to_le_bytes(),
    );
    rgce.extend_from_slice(&subexpr);
    rgce
}

fn rgce_memfunc_with_unknown_subexpr_and_outer_array() -> Vec<u8> {
    // Regression guard: scanning nested `PtgMem*` subexpressions should not prevent us from
    // detecting `PtgArray` that appears *outside* the mem token.
    //
    // This uses `PtgExp` (`0x01`) as an intentionally-unsupported-but-real ptg inside the mem
    // subexpression.
    let subexpr = vec![0x01, 0x00, 0x00, 0x00, 0x00]; // PtgExp + dummy coords

    let mut rgce = vec![0x29];
    rgce.extend_from_slice(
        &u16::try_from(subexpr.len())
            .expect("subexpression length fits in u16")
            .to_le_bytes(),
    );
    rgce.extend_from_slice(&subexpr);
    rgce.extend_from_slice(&fixture_builder::rgce::array_placeholder());
    rgce
}

fn rgce_extend_list_then_array() -> Vec<u8> {
    // PtgExtend: [ptg=0x18][etpg=0x19][payload: 12 bytes]
    //
    // This is the structured-reference (PtgList) extend token. The patcher should be able to
    // consume it and still detect later PtgArray tokens.
    let mut rgce = vec![0x18, 0x19];
    rgce.extend_from_slice(&[0u8; 12]);
    rgce.extend_from_slice(&fixture_builder::rgce::array_placeholder());
    rgce
}

fn rgce_extend_list_with_2_byte_prefix_then_array() -> Vec<u8> {
    // Some producers appear to insert prefix bytes before the canonical 12-byte PtgList payload.
    // Ensure `rgce_references_rgcb` can still skip the token and detect later PtgArray tokens.
    //
    // Payload bytes are chosen such that naive 12-byte consumption would desync the stream
    // (leaving trailing payload bytes before the `PtgArray` opcode).
    let mut rgce = vec![0x18, 0x19];
    rgce.extend_from_slice(&[0u8; 2]); // prefix padding
    rgce.extend_from_slice(&[
        0x01, 0x00, 0x00, 0x00, // table_id = 1
        0x00, 0x00, 0x10, 0x00, // col_first_raw = flags<<16 (flags=0x0010, col_first=0)
        0x00, 0x00, 0x00, 0x00, // col_last_raw = 0
    ]);
    rgce.extend_from_slice(&fixture_builder::rgce::array_placeholder());
    rgce
}

fn rgce_referr_then_array() -> Vec<u8> {
    // PtgRefErr: [ptg=0x2A][row:u32][col:u16]
    let mut rgce = vec![0x2A];
    rgce.extend_from_slice(&[0u8; 6]);
    rgce.extend_from_slice(&fixture_builder::rgce::array_placeholder());
    rgce
}

fn rgce_attr_choose_with_ptgarray_bytes_in_jump_table() -> Vec<u8> {
    // PtgAttrChoose:
    //   [ptg=0x19]
    //   [grbit=0x04 (tAttrChoose)]
    //   [wAttr: u16] (number of u16 offsets in the jump table)
    //   [jump table: wAttr * u16]
    //
    // Ensure the jump table contains `0x20` so a naive raw-byte scan would incorrectly
    // interpret it as `PtgArray`.
    let mut rgce = vec![0x19, 0x04];
    rgce.extend_from_slice(&1u16.to_le_bytes());
    rgce.extend_from_slice(&[0x20, 0x00]);
    rgce
}

fn rgce_str_literal_with_ptgarray_byte_sequence() -> Vec<u8> {
    // PtgStr: [ptg=0x17][cch: u16][utf16 chars...]
    //
    // Include a single UTF-16 space (0x0020 -> bytes 0x20 0x00) so a naive scan would
    // false-positive on `PtgArray`.
    vec![0x17, 0x01, 0x00, 0x20, 0x00]
}

fn rgce_ptgref_with_ptgarray_byte_sequence_in_row() -> Vec<u8> {
    // PtgRef: [ptg=0x24][row:u32][col:u16]
    //
    // Set the row to 0x20 so the payload contains a `0x20` byte. A naive raw-byte scan would
    // incorrectly interpret this as `PtgArray`.
    vec![0x24, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00]
}

fn ptg_str_literal(s: &str) -> Vec<u8> {
    // PtgStr: [ptg=0x17][cch: u16][utf16 chars...]
    let mut out = vec![0x17];
    let units: Vec<u16> = s.encode_utf16().collect();
    out.extend_from_slice(&(units.len() as u16).to_le_bytes());
    for unit in units {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out
}

fn ptg_bool_literal(v: bool) -> Vec<u8> {
    // PtgBool: [ptg=0x1D][b: u8]
    vec![0x1D, if v { 1 } else { 0 }]
}

fn ptg_err_literal(code: u8) -> Vec<u8> {
    // PtgErr: [ptg=0x1C][code: u8]
    vec![0x1C, code]
}

fn assert_invalid_input_contains(err: Error, needle: &str) {
    match err {
        Error::Io(io_err) => {
            assert_eq!(io_err.kind(), io::ErrorKind::InvalidInput);
            assert!(
                io_err.to_string().contains(needle),
                "expected error message to contain {needle:?}, got: {io_err}"
            );
        }
        other => panic!("expected Error::Io, got {other:?}"),
    }
}

fn read_sheet1_bin_from_fixture(bytes: &[u8]) -> Vec<u8> {
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).expect("open xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("read sheet1.bin");
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut out = Vec::new();
    entry.read_to_end(&mut out).expect("read sheet bytes");
    out
}

fn assert_patch_requires_rgcb(new_rgce: Vec<u8>) {
    // Create a minimal worksheet with a single numeric formula that has no trailing `rgcb` bytes.
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_formula_num(0, 0, 1.0, vec![0x1E, 0x01, 0x00], vec![]);

    let xlsb_bytes = builder.build_bytes();
    let sheet_bin = read_sheet1_bin_from_fixture(&xlsb_bytes);

    let err = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(1.0),
            new_style: None,
            clear_formula: false,
            new_formula: Some(new_rgce),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect_err("expected patch to reject missing rgcb bytes for PtgArray");

    let Error::Io(io_err) = err else {
        panic!("expected Error::Io, got {err:?}");
    };
    assert_eq!(io_err.kind(), io::ErrorKind::InvalidInput);
    assert!(
        io_err
            .to_string()
            .contains("requires rgcb bytes (PtgArray present); set CellEdit.new_rgcb"),
        "unexpected error message: {io_err}"
    );
}

fn assert_patch_does_not_require_rgcb(new_rgce: Vec<u8>) {
    // Create a minimal worksheet with a single numeric formula that has no trailing `rgcb` bytes.
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_formula_num(0, 0, 1.0, vec![0x1E, 0x01, 0x00], vec![]);

    let xlsb_bytes = builder.build_bytes();
    let sheet_bin = read_sheet1_bin_from_fixture(&xlsb_bytes);

    patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(1.0),
            new_style: None,
            clear_formula: false,
            new_formula: Some(new_rgce),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect("expected patch to succeed without rgcb bytes");
}

fn assert_streaming_patch_requires_rgcb(new_rgce: Vec<u8>) {
    // Create a minimal worksheet with a single numeric formula that has no trailing `rgcb` bytes.
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_formula_num(0, 0, 1.0, vec![0x1E, 0x01, 0x00], vec![]);

    let xlsb_bytes = builder.build_bytes();
    let sheet_bin = read_sheet1_bin_from_fixture(&xlsb_bytes);

    let mut out = Vec::new();
    let err = patch_sheet_bin_streaming(
        Cursor::new(&sheet_bin),
        &mut out,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Number(1.0),
            new_style: None,
            clear_formula: false,
            new_formula: Some(new_rgce),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect_err("expected streaming patch to reject missing rgcb bytes for PtgArray");

    let Error::Io(io_err) = err else {
        panic!("expected Error::Io, got {err:?}");
    };
    assert_eq!(io_err.kind(), io::ErrorKind::InvalidInput);
    assert!(
        io_err
            .to_string()
            .contains("requires rgcb bytes (PtgArray present); set CellEdit.new_rgcb"),
        "unexpected error message: {io_err}"
    );
}

fn assert_streaming_patch_does_not_require_rgcb(new_rgce: Vec<u8>) {
    // Create a minimal worksheet with a single numeric formula that has no trailing `rgcb` bytes.
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_formula_num(0, 0, 1.0, vec![0x1E, 0x01, 0x00], vec![]);

    let xlsb_bytes = builder.build_bytes();
    let sheet_bin = read_sheet1_bin_from_fixture(&xlsb_bytes);

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(1.0),
        new_style: None,
        clear_formula: false,
        new_formula: Some(new_rgce),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let patched_in_mem = patch_sheet_bin(&sheet_bin, &edits).expect("patch_sheet_bin");

    let mut patched_stream = Vec::new();
    patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut patched_stream, &edits)
        .expect("expected streaming patch to succeed without rgcb bytes");

    assert_eq!(patched_stream, patched_in_mem);
}

#[test]
fn rgce_references_rgcb_detects_ptgarray_inside_memfunc() {
    let rgce = rgce_memfunc_with_array();
    assert!(rgce_references_rgcb(&rgce));
}

#[test]
fn rgce_references_rgcb_detects_ptgarray_outside_memfunc_even_if_mem_subexpr_has_unknown_ptg() {
    let rgce = rgce_memfunc_with_unknown_subexpr_and_outer_array();
    assert!(rgce_references_rgcb(&rgce));
}

#[test]
fn rgce_references_rgcb_detects_ptgarray_after_ptgextend_list() {
    let rgce = rgce_extend_list_then_array();
    assert!(rgce_references_rgcb(&rgce));
}

#[test]
fn rgce_references_rgcb_detects_ptgarray_after_prefixed_ptgextend_list() {
    let rgce = rgce_extend_list_with_2_byte_prefix_then_array();
    assert!(rgce_references_rgcb(&rgce));
}

#[test]
fn rgce_references_rgcb_detects_ptgarray_after_ptgreferr() {
    let rgce = rgce_referr_then_array();
    assert!(rgce_references_rgcb(&rgce));
}

#[test]
fn patch_sheet_bin_errors_when_formula_requires_rgcb_but_new_rgcb_is_none() {
    assert_patch_requires_rgcb(rgce_memfunc_with_array());
    assert_patch_requires_rgcb(rgce_extend_list_then_array());
    assert_patch_requires_rgcb(rgce_extend_list_with_2_byte_prefix_then_array());
    assert_patch_requires_rgcb(rgce_referr_then_array());
}

#[test]
fn patch_sheet_bin_streaming_errors_when_formula_requires_rgcb_but_new_rgcb_is_none() {
    assert_streaming_patch_requires_rgcb(rgce_memfunc_with_array());
    assert_streaming_patch_requires_rgcb(rgce_extend_list_then_array());
    assert_streaming_patch_requires_rgcb(rgce_extend_list_with_2_byte_prefix_then_array());
    assert_streaming_patch_requires_rgcb(rgce_referr_then_array());
}

#[test]
fn patch_sheet_bin_rejects_converting_value_cell_to_formula_that_requires_rgcb_without_new_rgcb() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(1.0),
        new_style: None,
        clear_formula: false,
        new_formula: Some(fixture_builder::rgce::array_placeholder()),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let err = patch_sheet_bin(&sheet_bin, &edits).expect_err(
        "expected InvalidInput when converting value cell to formula without rgcb bytes",
    );
    match err {
        Error::Io(io_err) => {
            assert_eq!(io_err.kind(), io::ErrorKind::InvalidInput);
            assert!(
                io_err.to_string().contains("set CellEdit.new_rgcb"),
                "unexpected error message: {io_err}"
            );
        }
        other => panic!("expected Error::Io, got {other:?}"),
    }
}

#[test]
fn patch_sheet_bin_streaming_rejects_converting_value_cell_to_formula_that_requires_rgcb_without_new_rgcb(
) {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_number(0, 0, 1.0);
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Number(1.0),
        new_style: None,
        clear_formula: false,
        new_formula: Some(fixture_builder::rgce::array_placeholder()),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let mut out = Vec::new();
    let err = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut out, &edits).expect_err(
        "expected InvalidInput when streaming convert value cell to formula without rgcb bytes",
    );
    match err {
        Error::Io(io_err) => {
            assert_eq!(io_err.kind(), io::ErrorKind::InvalidInput);
            assert!(
                io_err.to_string().contains("set CellEdit.new_rgcb"),
                "unexpected error message: {io_err}"
            );
        }
        other => panic!("expected Error::Io, got {other:?}"),
    }
}

#[test]
fn patch_sheet_bin_streaming_allows_missing_rgcb_when_ptgarray_bytes_only_in_attr_choose_payload() {
    assert_streaming_patch_does_not_require_rgcb(rgce_attr_choose_with_ptgarray_bytes_in_jump_table());
    assert_streaming_patch_does_not_require_rgcb(rgce_str_literal_with_ptgarray_byte_sequence());
    assert_streaming_patch_does_not_require_rgcb(rgce_ptgref_with_ptgarray_byte_sequence_in_row());
}

#[test]
fn rgce_references_rgcb_does_not_false_positive_on_attr_choose_jump_table() {
    let rgce = rgce_attr_choose_with_ptgarray_bytes_in_jump_table();
    assert!(!rgce_references_rgcb(&rgce));
}

#[test]
fn rgce_references_rgcb_does_not_false_positive_on_ptgstr_payload() {
    let rgce = rgce_str_literal_with_ptgarray_byte_sequence();
    assert!(!rgce_references_rgcb(&rgce));
}

#[test]
fn rgce_references_rgcb_does_not_false_positive_on_ptgref_payload() {
    let rgce = rgce_ptgref_with_ptgarray_byte_sequence_in_row();
    assert!(!rgce_references_rgcb(&rgce));
}

#[test]
fn patch_sheet_bin_allows_missing_rgcb_when_ptgarray_bytes_only_in_attr_choose_payload() {
    assert_patch_does_not_require_rgcb(rgce_attr_choose_with_ptgarray_bytes_in_jump_table());
    assert_patch_does_not_require_rgcb(rgce_str_literal_with_ptgarray_byte_sequence());
    assert_patch_does_not_require_rgcb(rgce_ptgref_with_ptgarray_byte_sequence_in_row());
}

#[test]
fn patch_sheet_bin_rejects_brt_fmla_string_update_that_requires_rgcb_without_new_rgcb() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_formula_str(0, 0, "Hello", ptg_str_literal("Hello"));
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text("Hello".to_string()),
        new_style: None,
        clear_formula: false,
        new_formula: Some(fixture_builder::rgce::array_placeholder()),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let err = patch_sheet_bin(&sheet_bin, &edits)
        .expect_err("expected InvalidInput when updating formula string without rgcb");
    match err {
        Error::Io(io_err) => {
            assert_eq!(io_err.kind(), io::ErrorKind::InvalidInput);
            assert!(
                io_err.to_string().contains("set CellEdit.new_rgcb"),
                "unexpected error message: {io_err}"
            );
        }
        other => panic!("expected Error::Io, got {other:?}"),
    }
}

#[test]
fn patch_sheet_bin_streaming_rejects_brt_fmla_string_update_that_requires_rgcb_without_new_rgcb() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_formula_str(0, 0, "Hello", ptg_str_literal("Hello"));
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Text("Hello".to_string()),
        new_style: None,
        clear_formula: false,
        new_formula: Some(fixture_builder::rgce::array_placeholder()),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let mut out = Vec::new();
    let err = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut out, &edits)
        .expect_err("expected InvalidInput when streaming update formula string without rgcb");
    match err {
        Error::Io(io_err) => {
            assert_eq!(io_err.kind(), io::ErrorKind::InvalidInput);
            assert!(
                io_err.to_string().contains("set CellEdit.new_rgcb"),
                "unexpected error message: {io_err}"
            );
        }
        other => panic!("expected Error::Io, got {other:?}"),
    }
}

#[test]
fn patch_sheet_bin_rejects_brt_fmla_bool_update_that_requires_rgcb_without_new_rgcb() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_formula_bool(0, 0, true, ptg_bool_literal(true));
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Bool(true),
        new_style: None,
        clear_formula: false,
        new_formula: Some(fixture_builder::rgce::array_placeholder()),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let err = patch_sheet_bin(&sheet_bin, &edits)
        .expect_err("expected InvalidInput when updating formula bool without rgcb");
    match err {
        Error::Io(io_err) => {
            assert_eq!(io_err.kind(), io::ErrorKind::InvalidInput);
            assert!(
                io_err.to_string().contains("set CellEdit.new_rgcb"),
                "unexpected error message: {io_err}"
            );
        }
        other => panic!("expected Error::Io, got {other:?}"),
    }
}

#[test]
fn patch_sheet_bin_streaming_rejects_brt_fmla_bool_update_that_requires_rgcb_without_new_rgcb() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_formula_bool(0, 0, true, ptg_bool_literal(true));
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Bool(true),
        new_style: None,
        clear_formula: false,
        new_formula: Some(fixture_builder::rgce::array_placeholder()),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let mut out = Vec::new();
    let err = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut out, &edits)
        .expect_err("expected InvalidInput when streaming update formula bool without rgcb");
    match err {
        Error::Io(io_err) => {
            assert_eq!(io_err.kind(), io::ErrorKind::InvalidInput);
            assert!(
                io_err.to_string().contains("set CellEdit.new_rgcb"),
                "unexpected error message: {io_err}"
            );
        }
        other => panic!("expected Error::Io, got {other:?}"),
    }
}

#[test]
fn patch_sheet_bin_rejects_brt_fmla_error_update_that_requires_rgcb_without_new_rgcb() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_formula_err(0, 0, 0x2A, ptg_err_literal(0x2A)); // #N/A
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Error(0x2A),
        new_style: None,
        clear_formula: false,
        new_formula: Some(fixture_builder::rgce::array_placeholder()),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let err = patch_sheet_bin(&sheet_bin, &edits)
        .expect_err("expected InvalidInput when updating formula error without rgcb");
    match err {
        Error::Io(io_err) => {
            assert_eq!(io_err.kind(), io::ErrorKind::InvalidInput);
            assert!(
                io_err.to_string().contains("set CellEdit.new_rgcb"),
                "unexpected error message: {io_err}"
            );
        }
        other => panic!("expected Error::Io, got {other:?}"),
    }
}

#[test]
fn patch_sheet_bin_streaming_rejects_brt_fmla_error_update_that_requires_rgcb_without_new_rgcb() {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_formula_err(0, 0, 0x2A, ptg_err_literal(0x2A)); // #N/A
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let edits = [CellEdit {
        row: 0,
        col: 0,
        new_value: CellValue::Error(0x2A),
        new_style: None,
        clear_formula: false,
        new_formula: Some(fixture_builder::rgce::array_placeholder()),
        new_rgcb: None,
        new_formula_flags: None,
        shared_string_index: None,
    }];

    let mut out = Vec::new();
    let err = patch_sheet_bin_streaming(Cursor::new(&sheet_bin), &mut out, &edits)
        .expect_err("expected InvalidInput when streaming update formula error without rgcb");
    match err {
        Error::Io(io_err) => {
            assert_eq!(io_err.kind(), io::ErrorKind::InvalidInput);
            assert!(
                io_err.to_string().contains("set CellEdit.new_rgcb"),
                "unexpected error message: {io_err}"
            );
        }
        other => panic!("expected Error::Io, got {other:?}"),
    }
}

#[test]
fn patch_sheet_bin_rejects_brt_fmla_string_rgce_replacement_when_existing_rgcb_present_without_new_rgcb(
) {
    // Seed a value cell, convert it to a formula string cell that carries non-empty rgcb bytes,
    // then attempt to replace the rgce without explicitly providing new_rgcb.
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_inline_string(0, 0, "Hello");
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let sheet_with_rgcb = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("Hello".to_string()),
            new_style: None,
            clear_formula: false,
            new_formula: Some(fixture_builder::rgce::array_placeholder()),
            new_rgcb: Some(vec![0xAA]),
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect("convert value cell to formula string with rgcb");

    let err = patch_sheet_bin(
        &sheet_with_rgcb,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("Hello".to_string()),
            new_style: None,
            clear_formula: false,
            new_formula: Some(ptg_str_literal("Hello")),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect_err("expected InvalidInput when replacing rgce without new_rgcb");
    assert_invalid_input_contains(err, "provide CellEdit.new_rgcb");
}

#[test]
fn patch_sheet_bin_streaming_rejects_brt_fmla_string_rgce_replacement_when_existing_rgcb_present_without_new_rgcb(
) {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_inline_string(0, 0, "Hello");
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let sheet_with_rgcb = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("Hello".to_string()),
            new_style: None,
            clear_formula: false,
            new_formula: Some(fixture_builder::rgce::array_placeholder()),
            new_rgcb: Some(vec![0xAA]),
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect("convert value cell to formula string with rgcb");

    let mut out = Vec::new();
    let err = patch_sheet_bin_streaming(
        Cursor::new(&sheet_with_rgcb),
        &mut out,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Text("Hello".to_string()),
            new_style: None,
            clear_formula: false,
            new_formula: Some(ptg_str_literal("Hello")),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect_err("expected InvalidInput when streaming replace rgce without new_rgcb");
    assert_invalid_input_contains(err, "provide CellEdit.new_rgcb");
}

#[test]
fn patch_sheet_bin_rejects_brt_fmla_bool_rgce_replacement_when_existing_rgcb_present_without_new_rgcb(
) {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_bool(0, 0, true);
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let sheet_with_rgcb = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Bool(true),
            new_style: None,
            clear_formula: false,
            new_formula: Some(fixture_builder::rgce::array_placeholder()),
            new_rgcb: Some(vec![0xAA]),
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect("convert value cell to formula bool with rgcb");

    let err = patch_sheet_bin(
        &sheet_with_rgcb,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Bool(true),
            new_style: None,
            clear_formula: false,
            new_formula: Some(ptg_bool_literal(true)),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect_err("expected InvalidInput when replacing rgce without new_rgcb");
    assert_invalid_input_contains(err, "provide CellEdit.new_rgcb");
}

#[test]
fn patch_sheet_bin_streaming_rejects_brt_fmla_bool_rgce_replacement_when_existing_rgcb_present_without_new_rgcb(
) {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_bool(0, 0, true);
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let sheet_with_rgcb = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Bool(true),
            new_style: None,
            clear_formula: false,
            new_formula: Some(fixture_builder::rgce::array_placeholder()),
            new_rgcb: Some(vec![0xAA]),
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect("convert value cell to formula bool with rgcb");

    let mut out = Vec::new();
    let err = patch_sheet_bin_streaming(
        Cursor::new(&sheet_with_rgcb),
        &mut out,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Bool(true),
            new_style: None,
            clear_formula: false,
            new_formula: Some(ptg_bool_literal(true)),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect_err("expected InvalidInput when streaming replace rgce without new_rgcb");
    assert_invalid_input_contains(err, "provide CellEdit.new_rgcb");
}

#[test]
fn patch_sheet_bin_rejects_brt_fmla_error_rgce_replacement_when_existing_rgcb_present_without_new_rgcb(
) {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_error(0, 0, 0x2A); // #N/A
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let sheet_with_rgcb = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Error(0x2A),
            new_style: None,
            clear_formula: false,
            new_formula: Some(fixture_builder::rgce::array_placeholder()),
            new_rgcb: Some(vec![0xAA]),
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect("convert value cell to formula error with rgcb");

    let err = patch_sheet_bin(
        &sheet_with_rgcb,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Error(0x2A),
            new_style: None,
            clear_formula: false,
            new_formula: Some(ptg_err_literal(0x2A)),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect_err("expected InvalidInput when replacing rgce without new_rgcb");
    assert_invalid_input_contains(err, "provide CellEdit.new_rgcb");
}

#[test]
fn patch_sheet_bin_streaming_rejects_brt_fmla_error_rgce_replacement_when_existing_rgcb_present_without_new_rgcb(
) {
    let mut builder = XlsbFixtureBuilder::new();
    builder.set_cell_error(0, 0, 0x2A); // #N/A
    let sheet_bin = read_sheet1_bin_from_fixture(&builder.build_bytes());

    let sheet_with_rgcb = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Error(0x2A),
            new_style: None,
            clear_formula: false,
            new_formula: Some(fixture_builder::rgce::array_placeholder()),
            new_rgcb: Some(vec![0xAA]),
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect("convert value cell to formula error with rgcb");

    let mut out = Vec::new();
    let err = patch_sheet_bin_streaming(
        Cursor::new(&sheet_with_rgcb),
        &mut out,
        &[CellEdit {
            row: 0,
            col: 0,
            new_value: CellValue::Error(0x2A),
            new_style: None,
            clear_formula: false,
            new_formula: Some(ptg_err_literal(0x2A)),
            new_rgcb: None,
            new_formula_flags: None,
            shared_string_index: None,
        }],
    )
    .expect_err("expected InvalidInput when streaming replace rgce without new_rgcb");
    assert_invalid_input_contains(err, "provide CellEdit.new_rgcb");
}
