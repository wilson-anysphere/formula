#![cfg(target_arch = "wasm32")]

use formula_core::CellChange;
use serde_json::json;
use wasm_bindgen::JsValue;
use wasm_bindgen_test::wasm_bindgen_test;

use formula_wasm::WasmWorkbook;

#[wasm_bindgen_test]
fn recalculate_reports_changed_cells() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(
        changes,
        vec![CellChange {
            sheet: formula_core::DEFAULT_SHEET.to_string(),
            address: "A2".to_string(),
            value: json!(2.0),
        }]
    );

    let cell_js = wb.get_cell("A2".to_string(), None).unwrap();
    let cell: formula_core::CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.value, json!(2.0));
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_imports_formulas_and_recalculates() {
    let mut wb = WasmWorkbook::from_xlsx_bytes(include_bytes!(
        "../../../fixtures/xlsx/formulas/formulas.xlsx"
    ))
    .unwrap();

    wb.recalculate(None).unwrap();

    let cell_js = wb.get_cell("C1".to_string(), None).unwrap();
    let cell: formula_core::CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, json!("=A1+B1"));
    assert_eq!(cell.value, json!(3.0));
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_imports_literal_cells() {
    let wb = WasmWorkbook::from_xlsx_bytes(include_bytes!(
        "../../../fixtures/xlsx/basic/basic.xlsx"
    ))
    .unwrap();

    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: formula_core::CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_eq!(a1.input, json!(1.0));
    assert_eq!(a1.value, json!(1.0));

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: formula_core::CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.input, json!("Hello"));
    assert_eq!(b1.value, json!("Hello"));
}
