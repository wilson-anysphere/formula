#![cfg(not(target_arch = "wasm32"))]

use std::io::Cursor;

use formula_engine::Value as EngineValue;
use formula_model::{CellValue, Style, Workbook};
use formula_wasm::{WasmWorkbook, DEFAULT_SHEET};

fn workbook_bytes(full_precision: bool) -> Vec<u8> {
    let mut workbook = Workbook::new();
    workbook.calc_settings.full_precision = full_precision;

    let style_id = workbook.styles.intern(Style {
        number_format: Some("0.00".to_string()),
        ..Style::default()
    });

    let sheet_id = workbook.add_sheet(DEFAULT_SHEET).unwrap();
    let sheet = workbook.sheet_mut(sheet_id).unwrap();
    sheet
        .set_value_a1("A1", CellValue::Number(1.239))
        .unwrap();
    sheet.set_style_id_a1("A1", style_id).unwrap();
    sheet
        .set_formula_a1("B1", Some("A1".to_string()))
        .unwrap();

    let mut cursor = Cursor::new(Vec::new());
    formula_xlsx::write_workbook_to_writer(&workbook, &mut cursor).unwrap();
    cursor.into_inner()
}

#[test]
fn from_xlsx_bytes_rounds_cached_numbers_when_precision_as_displayed_enabled() {
    let bytes = workbook_bytes(false);
    let mut wb = WasmWorkbook::from_xlsx_bytes(&bytes).unwrap();

    assert_eq!(
        wb.debug_get_engine_value(DEFAULT_SHEET, "A1"),
        EngineValue::Number(1.24)
    );

    wb.debug_recalculate().unwrap();
    assert_eq!(
        wb.debug_get_engine_value(DEFAULT_SHEET, "B1"),
        EngineValue::Number(1.24)
    );
}

#[test]
fn from_xlsx_bytes_preserves_cached_numbers_when_full_precision_enabled() {
    let bytes = workbook_bytes(true);
    let mut wb = WasmWorkbook::from_xlsx_bytes(&bytes).unwrap();

    assert_eq!(
        wb.debug_get_engine_value(DEFAULT_SHEET, "A1"),
        EngineValue::Number(1.239)
    );

    wb.debug_recalculate().unwrap();
    assert_eq!(
        wb.debug_get_engine_value(DEFAULT_SHEET, "B1"),
        EngineValue::Number(1.239)
    );
}

