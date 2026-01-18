#![cfg(not(target_arch = "wasm32"))]

use formula_engine::Value as EngineValue;
use formula_model::Workbook;
use formula_wasm::{WasmWorkbook, DEFAULT_SHEET};

#[test]
fn from_xlsx_bytes_imports_sheet_default_col_width_for_cell_width() {
    let mut workbook = Workbook::new();
    let sheet_id = workbook.add_sheet(DEFAULT_SHEET).unwrap();
    let sheet = workbook.sheet_mut(sheet_id).unwrap();
    sheet.default_col_width = Some(20.0);
    sheet.set_formula_a1("B1", Some(r#"CELL("width",A1)"#.to_string()))
        .unwrap();

    let bytes = formula_xlsx::XlsxDocument::new(workbook)
        .save_to_vec()
        .unwrap();
    let mut wb = WasmWorkbook::from_xlsx_bytes(&bytes).unwrap();
    wb.debug_recalculate().unwrap();

    // Excel returns the floored width with a `0` flag (no +0.1) when the column uses the sheet
    // default width.
    assert_eq!(
        wb.debug_get_engine_value(DEFAULT_SHEET, "B1"),
        EngineValue::Number(20.0)
    );
}

