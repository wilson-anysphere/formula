#![cfg(target_arch = "wasm32")]

use serde_json::Value as JsonValue;
use serde_json::json;
use wasm_bindgen::JsValue;
use wasm_bindgen_test::wasm_bindgen_test;

use formula_wasm::WasmWorkbook;

#[derive(Debug, serde::Deserialize)]
struct CellData {
    value: JsonValue,
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_imports_styles_for_cells_rows_and_cols() {
    let bytes = include_bytes!("fixtures/import_styles_cols.xlsx");
    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();

    // A1 is a style-only cell (no value/formula) with a non-default number format (0.00).
    wb.set_cell(
        "D1".to_string(),
        JsValue::from_str("=CELL(\"format\",A1)"),
        None,
    )
    .unwrap();

    // Column C has a default style (index 1), so C1 inherits the same number format.
    wb.set_cell(
        "D2".to_string(),
        JsValue::from_str("=CELL(\"format\",C1)"),
        None,
    )
    .unwrap();

    // Row 3 has a default style (index 1), so A3 inherits the same number format.
    wb.set_cell(
        "D3".to_string(),
        JsValue::from_str("=CELL(\"format\",A3)"),
        None,
    )
    .unwrap();

    // Column B is hidden, so its width should be reported as 0.
    wb.set_cell(
        "D4".to_string(),
        JsValue::from_str("=CELL(\"width\",B1)"),
        None,
    )
    .unwrap();

    // Column D has a custom width override, so `CELL("width")` should return the floored width
    // with a `.1` flag.
    wb.set_cell(
        "D5".to_string(),
        JsValue::from_str("=CELL(\"width\",D1)"),
        None,
    )
    .unwrap();

    wb.recalculate(None).unwrap();

    let d1: CellData = serde_wasm_bindgen::from_value(wb.get_cell("D1".to_string(), None).unwrap())
        .unwrap();
    let d2: CellData = serde_wasm_bindgen::from_value(wb.get_cell("D2".to_string(), None).unwrap())
        .unwrap();
    let d3: CellData = serde_wasm_bindgen::from_value(wb.get_cell("D3".to_string(), None).unwrap())
        .unwrap();
    let d4: CellData = serde_wasm_bindgen::from_value(wb.get_cell("D4".to_string(), None).unwrap())
        .unwrap();
    let d5: CellData = serde_wasm_bindgen::from_value(wb.get_cell("D5".to_string(), None).unwrap())
        .unwrap();

    assert_eq!(d1.value, JsonValue::String("F2".to_string()));
    assert_eq!(d2.value, JsonValue::String("F2".to_string()));
    assert_eq!(d3.value, JsonValue::String("F2".to_string()));
    assert_eq!(d4.value, json!(0.0));
    assert_eq!(d5.value, json!(12.1));
}
