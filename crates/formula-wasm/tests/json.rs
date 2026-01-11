use serde_json::json;

use formula_wasm::WasmWorkbook;

#[test]
fn to_json_preserves_engine_workbook_schema() {
    // Avoid `JsValue` constructors in a native test binary (they panic on non-wasm32 targets).
    let input = json!({
        "sheets": {
            "Sheet1": {
                "cells": {
                    "A1": 1.0,
                    "A2": "=A1*2"
                }
            }
        }
    })
    .to_string();

    let wb = WasmWorkbook::from_json(&input).unwrap();
    let json_str = wb.to_json().unwrap();
    let parsed: serde_json::Value = serde_json::from_str(&json_str).unwrap();

    // Ensure we preserve the `{ sheets: { [name]: { cells: { [A1]: scalarOrFormulaString }}}}`
    // schema consumed by `packages/engine`.
    assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A1"], json!(1.0));
    assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A2"], json!("=A1*2"));

    // Roundtrip through `fromJson` should succeed and keep the same shape.
    let wb2 = WasmWorkbook::from_json(&json_str).unwrap();
    let json_str2 = wb2.to_json().unwrap();
    let parsed2: serde_json::Value = serde_json::from_str(&json_str2).unwrap();
    assert_eq!(parsed2["sheets"]["Sheet1"]["cells"]["A2"], json!("=A1*2"));
}
