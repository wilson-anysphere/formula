use std::io;
use std::io::{Cursor, Read};

use formula_xlsb::{patch_sheet_bin, rgce_references_rgcb, CellEdit, CellValue, Error};

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

fn rgce_referr_then_array() -> Vec<u8> {
    // PtgRefErr: [ptg=0x2A][row:u32][col:u16]
    let mut rgce = vec![0x2A];
    rgce.extend_from_slice(&[0u8; 6]);
    rgce.extend_from_slice(&fixture_builder::rgce::array_placeholder());
    rgce
}

fn read_sheet1_bin_from_fixture(bytes: &[u8]) -> Vec<u8> {
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).expect("open xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("read sheet1.bin");
    let mut out = Vec::with_capacity(entry.size() as usize);
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
fn rgce_references_rgcb_detects_ptgarray_after_ptgreferr() {
    let rgce = rgce_referr_then_array();
    assert!(rgce_references_rgcb(&rgce));
}

#[test]
fn patch_sheet_bin_errors_when_formula_requires_rgcb_but_new_rgcb_is_none() {
    assert_patch_requires_rgcb(rgce_memfunc_with_array());
    assert_patch_requires_rgcb(rgce_extend_list_then_array());
    assert_patch_requires_rgcb(rgce_referr_then_array());
}
