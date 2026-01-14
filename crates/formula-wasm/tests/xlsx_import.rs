#![cfg(target_arch = "wasm32")]

use serde_json::Value as JsonValue;
use wasm_bindgen::JsValue;
use wasm_bindgen_test::wasm_bindgen_test;

use formula_wasm::WasmWorkbook;

#[derive(Debug, serde::Deserialize)]
struct CellData {
    value: JsonValue,
}

fn assert_json_number(value: &JsonValue, expected: f64) {
    let got = value
        .as_f64()
        .unwrap_or_else(|| panic!("expected JSON number, got {value:?}"));
    assert!(
        (got - expected).abs() < 1e-9,
        "expected {expected}, got {got} ({value:?})"
    );
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

    // The fixture encodes a persisted view origin via `sheetView/pane/@topLeftCell`.
    wb.set_cell(
        "D6".to_string(),
        JsValue::from_str("=INFO(\"origin\")"),
        None,
    )
    .unwrap();

    // The imported style includes horizontal alignment (right) and protection (unlocked).
    wb.set_cell(
        "D7".to_string(),
        JsValue::from_str("=CELL(\"prefix\",A1)"),
        None,
    )
    .unwrap();
    wb.set_cell(
        "D8".to_string(),
        JsValue::from_str("=CELL(\"protect\",A1)"),
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
    let d6: CellData = serde_wasm_bindgen::from_value(wb.get_cell("D6".to_string(), None).unwrap())
        .unwrap();
    let d7: CellData = serde_wasm_bindgen::from_value(wb.get_cell("D7".to_string(), None).unwrap())
        .unwrap();
    let d8: CellData = serde_wasm_bindgen::from_value(wb.get_cell("D8".to_string(), None).unwrap())
        .unwrap();

    assert_eq!(d1.value, JsonValue::String("F2".to_string()));
    assert_eq!(d2.value, JsonValue::String("F2".to_string()));
    assert_eq!(d3.value, JsonValue::String("F2".to_string()));
    assert_json_number(&d4.value, 0.0);
    assert_json_number(&d5.value, 12.1);
    assert_eq!(d6.value, JsonValue::String("$C$5".to_string()));
    assert_eq!(d7.value, JsonValue::String("\"".to_string()));
    assert_json_number(&d8.value, 0.0);
}
