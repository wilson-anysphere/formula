use std::io::Cursor;

use formula_model::{CellRef, CellValue};
use formula_xlsx::{
    load_from_bytes, patch_xlsx_streaming_workbook_cell_patches, CellPatch, WorkbookCellPatches,
};
use rust_xlsxwriter::Workbook;

#[test]
fn streaming_patches_match_sheet_names_nfkc_case_insensitively() {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name("Kelvin")
        .expect("set worksheet name to Kelvin sign");
    let input_bytes = workbook.save_to_buffer().expect("write xlsx bytes");

    let cell = CellRef::from_a1("A1").unwrap();

    let mut patches = WorkbookCellPatches::default();
    patches.set_cell(
        // U+212A KELVIN SIGN (K) is NFKC-equivalent to ASCII 'K'. Excel matches sheet names under
        // compatibility normalization + case-insensitive comparison.
        "Kelvin",
        cell,
        CellPatch::set_value(CellValue::String("patched".to_string())),
    );

    let mut out = Cursor::new(Vec::new());
    patch_xlsx_streaming_workbook_cell_patches(Cursor::new(input_bytes), &mut out, &patches)
        .expect("streaming patch should resolve sheet selector");
    let out_bytes = out.into_inner();

    let doc = load_from_bytes(&out_bytes).expect("read patched workbook");
    let sheet = doc
        .workbook
        .sheet_by_name("Kelvin")
        .expect("sheet should exist");
    assert_eq!(sheet.value(cell), CellValue::String("patched".to_string()));
}
