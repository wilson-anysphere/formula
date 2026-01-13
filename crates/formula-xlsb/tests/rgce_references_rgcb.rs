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

fn read_sheet1_bin_from_fixture(bytes: &[u8]) -> Vec<u8> {
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).expect("open xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("read sheet1.bin");
    let mut out = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut out).expect("read sheet bytes");
    out
}

#[test]
fn rgce_references_rgcb_detects_ptgarray_inside_memfunc() {
    let rgce = rgce_memfunc_with_array();
    assert!(rgce_references_rgcb(&rgce));
}

#[test]
fn patch_sheet_bin_errors_when_formula_requires_rgcb_but_new_rgcb_is_none() {
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
            new_formula: Some(rgce_memfunc_with_array()),
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
