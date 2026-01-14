#![cfg(feature = "write")]

use std::io::{Cursor, Read};

use formula_xlsb::workbook_context::WorkbookContext;
use formula_xlsb::{parse_sheet_bin_with_context, patch_sheet_bin, CellEdit, CellValue};
use pretty_assertions::assert_eq;

mod fixture_builder;
use fixture_builder::XlsbFixtureBuilder;

#[test]
fn cell_edit_with_formula_text_with_context_in_sheet_encodes_tableless_structured_ref_in_multi_table_workbook(
) {
    let mut ctx = WorkbookContext::default();
    ctx.add_table(1, "Table1");
    ctx.add_table_column(1, 1, "Item");
    ctx.add_table_column(1, 2, "Qty");
    ctx.add_table_range(1, "Sheet1".to_string(), 0, 0, 9, 1); // A1:B10

    ctx.add_table(2, "Table2");
    ctx.add_table_column(2, 1, "Item");
    ctx.add_table_column(2, 2, "Qty");
    ctx.add_table_range(2, "Sheet1".to_string(), 0, 3, 9, 4); // D1:E10

    let builder = XlsbFixtureBuilder::new();
    let xlsb_bytes = builder.build_bytes();
    let mut zip = zip::ZipArchive::new(Cursor::new(xlsb_bytes)).expect("open xlsb zip");
    let mut entry = zip
        .by_name("xl/worksheets/sheet1.bin")
        .expect("find sheet1.bin");
    // Do not trust `ZipFile::size()` for allocation; ZIP metadata is untrusted and can
    // advertise enormous uncompressed sizes (zip-bomb style OOM).
    let mut sheet_bin = Vec::new();
    entry.read_to_end(&mut sheet_bin).expect("read sheet bytes");

    // D2 (inside Table2).
    let row = 1;
    let col = 3;

    let edit = CellEdit::with_formula_text_with_context_in_sheet(
        row,
        col,
        CellValue::Number(0.0),
        "=[@Qty]",
        "Sheet1",
        &ctx,
    )
    .expect("encode formula with context + sheet");

    let patched = patch_sheet_bin(&sheet_bin, &[edit]).expect("patch sheet bin");
    let parsed = parse_sheet_bin_with_context(&mut Cursor::new(&patched), &[], &ctx)
        .expect("parse patched sheet");
    let cell = parsed
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (row, col))
        .expect("patched cell exists");
    assert_eq!(cell.value, CellValue::Number(0.0));

    let formula = cell.formula.as_ref().expect("formula metadata preserved");
    assert_eq!(formula.text.as_deref(), Some("[@Qty]"));
    assert_eq!(
        formula.rgce,
        vec![
            0x18, 0x19, // PtgExtend + etpg=PtgList
            2, 0, 0, 0, // inferred Table2 id
            0x10, 0x00, // flags (#This Row)
            2, 0, // col_first (Qty)
            2, 0, // col_last (Qty)
            0, 0, // reserved
        ]
    );
}
