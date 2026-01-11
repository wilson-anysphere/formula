use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::time::{SystemTime, UNIX_EPOCH};

use formula_xlsb::{patch_sheet_bin, CellEdit, CellValue, XlsbWorkbook};
use pretty_assertions::assert_eq;

fn read_zip_part(path: &str, part_path: &str) -> Vec<u8> {
    let file = File::open(path).expect("open xlsb fixture");
    let mut zip = zip::ZipArchive::new(file).expect("open zip");
    let mut entry = zip.by_name(part_path).expect("find part");
    let mut bytes = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut bytes).expect("read part bytes");
    bytes
}

#[test]
fn patch_sheet_bin_round_trips_numeric_cell_preserving_formula() {
    let fixture = concat!(env!("CARGO_MANIFEST_DIR"), "/tests/fixtures/simple.xlsb");
    let wb = XlsbWorkbook::open(fixture).expect("open xlsb");
    let sheet_part = wb.sheet_metas()[0].part_path.clone();

    let original_sheet = wb.read_sheet(0).expect("read original sheet");
    let original_formula_rgce = original_sheet
        .cells
        .iter()
        .find(|c| (c.row, c.col) == (0, 2))
        .and_then(|c| c.formula.as_ref())
        .expect("original formula present")
        .rgce
        .clone();

    let sheet_bin = read_zip_part(fixture, &sheet_part);
    let patched_sheet_bin = patch_sheet_bin(
        &sheet_bin,
        &[CellEdit {
            row: 0,
            col: 1,
            new_value: CellValue::Number(100.0),
            new_formula: None,
        }],
    )
    .expect("patch sheet bin");

    let out_path = std::env::temp_dir().join(format!(
        "formula-xlsb-patch-{}.xlsb",
        SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("time")
            .as_nanos()
    ));

    wb.save_with_part_overrides(
        &out_path,
        &HashMap::from([(sheet_part.clone(), patched_sheet_bin)]),
    )
    .expect("write patched workbook");

    let patched = XlsbWorkbook::open(&out_path).expect("open patched workbook");
    let patched_sheet = patched.read_sheet(0).expect("read patched sheet");

    let mut cells = patched_sheet
        .cells
        .iter()
        .map(|c| ((c.row, c.col), c))
        .collect::<HashMap<_, _>>();

    assert_eq!(
        cells.remove(&(0, 1)).expect("patched B1").value,
        CellValue::Number(100.0)
    );

    let formula_cell = cells.remove(&(0, 2)).expect("formula cell still present");
    assert_eq!(formula_cell.value, CellValue::Number(85.0));
    let formula = formula_cell.formula.as_ref().expect("formula metadata preserved");
    assert_eq!(formula.text.as_deref(), Some("B1*2"));
    assert_eq!(formula.rgce.as_slice(), original_formula_rgce.as_slice());

    let _ = std::fs::remove_file(&out_path);
}
