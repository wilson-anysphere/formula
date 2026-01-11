#![cfg(target_arch = "wasm32")]

use formula_core::CellChange;
use serde_json::json;
use serde_json::Value as JsonValue;
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
            value: json!(2),
        }]
    );

    let cell_js = wb.get_cell("A2".to_string(), None).unwrap();
    let cell: formula_core::CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.value, json!(2));
}

#[wasm_bindgen_test]
fn recalculate_reports_dynamic_array_spills() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell(
        "A1".to_string(),
        JsValue::from_str("=SEQUENCE(1,2)"),
        None,
    )
    .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(
        changes,
        vec![
            CellChange {
                sheet: formula_core::DEFAULT_SHEET.to_string(),
                address: "A1".to_string(),
                value: json!(1),
            },
            CellChange {
                sheet: formula_core::DEFAULT_SHEET.to_string(),
                address: "B1".to_string(),
                value: json!(2),
            },
        ]
    );
}

#[wasm_bindgen_test]
fn recalculate_reports_spill_resize_clears_trailing_cells() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell(
        "A1".to_string(),
        JsValue::from_str("=SEQUENCE(1,3)"),
        None,
    )
    .unwrap();

    wb.recalculate(None).unwrap();

    // Shrink the spill width from 3 -> 2; `C1` should be cleared and surfaced as a recalc delta.
    wb.set_cell(
        "A1".to_string(),
        JsValue::from_str("=SEQUENCE(1,2)"),
        None,
    )
    .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(
        changes,
        vec![
            CellChange {
                sheet: formula_core::DEFAULT_SHEET.to_string(),
                address: "A1".to_string(),
                value: json!(1),
            },
            CellChange {
                sheet: formula_core::DEFAULT_SHEET.to_string(),
                address: "B1".to_string(),
                value: json!(2),
            },
            CellChange {
                sheet: formula_core::DEFAULT_SHEET.to_string(),
                address: "C1".to_string(),
                value: JsonValue::Null,
            },
        ]
    );
}

#[wasm_bindgen_test]
fn recalculate_filters_changes_by_sheet_name() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    wb.set_cell(
        "A1".to_string(),
        JsValue::from_f64(3.0),
        Some("Sheet2".to_string()),
    )
    .unwrap();
    wb.set_cell(
        "A2".to_string(),
        JsValue::from_str("=A1*2"),
        Some("Sheet2".to_string()),
    )
    .unwrap();

    // Establish a baseline so subsequent sheet-scoped recalcs only report new changes.
    wb.recalculate(None).unwrap();

    wb.set_cell(
        "A1".to_string(),
        JsValue::from_f64(4.0),
        Some("Sheet2".to_string()),
    )
    .unwrap();

    let changes_js = wb.recalculate(Some("sHeEt2".to_string())).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(
        changes,
        vec![CellChange {
            sheet: "Sheet2".to_string(),
            address: "A2".to_string(),
            value: json!(8),
        }]
    );
}

#[wasm_bindgen_test]
fn recalculate_reports_cleared_spill_outputs_after_edit() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell(
        "A1".to_string(),
        JsValue::from_str("=SEQUENCE(1,3)"),
        None,
    )
    .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(
        changes,
        vec![
            CellChange {
                sheet: formula_core::DEFAULT_SHEET.to_string(),
                address: "A1".to_string(),
                value: json!(1),
            },
            CellChange {
                sheet: formula_core::DEFAULT_SHEET.to_string(),
                address: "B1".to_string(),
                value: json!(2),
            },
            CellChange {
                sheet: formula_core::DEFAULT_SHEET.to_string(),
                address: "C1".to_string(),
                value: json!(3),
            },
        ]
    );

    // Overwrite a spill output cell with a literal value. This clears the spill footprint before
    // the next recalc, so `recalculate()` must still report the remaining spill outputs as blank.
    wb.set_cell("B1".to_string(), JsValue::from_f64(99.0), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(
        changes,
        vec![
            CellChange {
                sheet: formula_core::DEFAULT_SHEET.to_string(),
                address: "A1".to_string(),
                value: json!("#SPILL!"),
            },
            CellChange {
                sheet: formula_core::DEFAULT_SHEET.to_string(),
                address: "C1".to_string(),
                value: JsonValue::Null,
            },
        ]
    );
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_imports_formulas_and_recalculates() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/formulas/formulas.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();
    wb.recalculate(None).unwrap();

    let cell_js = wb.get_cell("C1".to_string(), None).unwrap();
    let cell: formula_core::CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, json!("=A1+B1"));
    assert_eq!(cell.value, json!(3));
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_loads_basic_fixture() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/basic.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();

    // Should not error even though the fixture contains no formulas.
    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert!(changes.is_empty());

    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: formula_core::CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_eq!(a1.input, json!(1));
    assert_eq!(a1.value, json!(1));

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: formula_core::CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.input, json!("Hello"));
    assert_eq!(b1.value, json!("Hello"));
}

#[wasm_bindgen_test]
fn null_inputs_clear_cells_and_recalculate_dependents() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    // Clear A1 by setting it to `null` (empty cell in the JS protocol).
    wb.set_cell("A1".to_string(), JsValue::NULL, None).unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(
        changes,
        vec![CellChange {
            sheet: formula_core::DEFAULT_SHEET.to_string(),
            address: "A2".to_string(),
            value: json!(0),
        }]
    );

    let cell_js = wb.get_cell("A1".to_string(), None).unwrap();
    let cell: formula_core::CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, JsonValue::Null);
    assert_eq!(cell.value, JsonValue::Null);

    let exported = wb.to_json().unwrap();
    let parsed: JsonValue = serde_json::from_str(&exported).unwrap();
    let cells = parsed["sheets"][formula_core::DEFAULT_SHEET]["cells"]
        .as_object()
        .unwrap();
    assert!(!cells.contains_key("A1"));
}

#[wasm_bindgen_test]
fn from_json_treats_null_cells_as_absent() {
    let json_str = r#"{
        "sheets": {
            "Sheet1": {
                "cells": {
                    "A1": null,
                    "A2": "=A1*2"
                }
            }
        }
    }"#;

    let mut wb = WasmWorkbook::from_json(json_str).unwrap();
    wb.recalculate(None).unwrap();

    let exported = wb.to_json().unwrap();
    let parsed: JsonValue = serde_json::from_str(&exported).unwrap();
    let cells = parsed["sheets"][formula_core::DEFAULT_SHEET]["cells"]
        .as_object()
        .unwrap();

    // JSON import should not store explicit `null` cells.
    assert!(!cells.contains_key("A1"));
    assert!(cells.contains_key("A2"));
}

#[wasm_bindgen_test]
fn set_range_clears_null_entries() {
    let mut wb = WasmWorkbook::new();

    let values: Vec<Vec<JsonValue>> = vec![vec![json!(1), json!(2)]];
    wb.set_range(
        "A1:B1".to_string(),
        serde_wasm_bindgen::to_value(&values).unwrap(),
        None,
    )
    .unwrap();

    let cleared: Vec<Vec<JsonValue>> = vec![vec![JsonValue::Null, json!(2)]];
    wb.set_range(
        "A1:B1".to_string(),
        serde_wasm_bindgen::to_value(&cleared).unwrap(),
        None,
    )
    .unwrap();

    let cell_js = wb.get_cell("A1".to_string(), None).unwrap();
    let cell: formula_core::CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, JsonValue::Null);
    assert_eq!(cell.value, JsonValue::Null);

    let exported = wb.to_json().unwrap();
    let parsed: JsonValue = serde_json::from_str(&exported).unwrap();
    let cells = parsed["sheets"][formula_core::DEFAULT_SHEET]["cells"]
        .as_object()
        .unwrap();
    assert!(!cells.contains_key("A1"));
}
