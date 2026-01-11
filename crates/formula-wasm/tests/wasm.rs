#![cfg(target_arch = "wasm32")]

use serde_json::json;
use serde_json::Value as JsonValue;
use wasm_bindgen::JsValue;
use wasm_bindgen_test::wasm_bindgen_test;

use formula_wasm::{WasmWorkbook, DEFAULT_SHEET};

fn assert_json_number(value: &JsonValue, expected: f64) {
    let actual = value
        .as_f64()
        .unwrap_or_else(|| panic!("expected JSON number, got {value:?}"));
    assert_eq!(actual, expected);
}

#[derive(Debug, serde::Deserialize, PartialEq)]
struct CellChange {
    sheet: String,
    address: String,
    value: JsonValue,
}

#[derive(Debug, serde::Deserialize, PartialEq)]
struct CellData {
    sheet: String,
    address: String,
    input: JsonValue,
    value: JsonValue,
}

#[wasm_bindgen_test]
fn debug_function_registry_contains_builtins() {
    // Ensure the wasm module invoked Rust global constructors before touching the
    // function registry (otherwise it can be cached as empty).
    let _ = WasmWorkbook::new();
    assert!(formula_engine::functions::lookup_function("SUM").is_some());
    assert!(formula_engine::functions::lookup_function("SEQUENCE").is_some());
}

#[wasm_bindgen_test]
fn recalculate_reports_changed_cells() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 1);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "A2");
    assert_json_number(&changes[0].value, 2.0);

    let cell_js = wb.get_cell("A2".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_json_number(&cell.value, 2.0);
}

#[wasm_bindgen_test]
fn recalculate_returns_empty_when_no_cells_changed() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 1);

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert!(changes.is_empty());
}

#[wasm_bindgen_test]
fn recalculate_reports_lambda_values_as_placeholder_text() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell(
        "A1".to_string(),
        JsValue::from_str("=LAMBDA(x,x)"),
        None,
    )
    .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 1);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "A1");
    assert_eq!(
        changes[0].value,
        JsonValue::String("<LAMBDA>".to_string())
    );

    let cell_js = wb.get_cell("A1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.value, JsonValue::String("<LAMBDA>".to_string()));
}

#[wasm_bindgen_test]
fn recalculate_reports_dynamic_array_spills() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_str("=SEQUENCE(1,2)"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 2);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "A1");
    assert_json_number(&changes[0].value, 1.0);
    assert_eq!(changes[1].sheet, DEFAULT_SHEET);
    assert_eq!(changes[1].address, "B1");
    assert_json_number(&changes[1].value, 2.0);
}

#[wasm_bindgen_test]
fn recalculate_reports_spill_resize_clears_trailing_cells() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_str("=SEQUENCE(1,3)"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    // Shrink the spill width from 3 -> 2; `C1` should be cleared and surfaced as a recalc delta.
    wb.set_cell("A1".to_string(), JsValue::from_str("=SEQUENCE(1,2)"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 3);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "A1");
    assert_json_number(&changes[0].value, 1.0);
    assert_eq!(changes[1].sheet, DEFAULT_SHEET);
    assert_eq!(changes[1].address, "B1");
    assert_json_number(&changes[1].value, 2.0);
    assert_eq!(changes[2].sheet, DEFAULT_SHEET);
    assert_eq!(changes[2].address, "C1");
    assert!(changes[2].value.is_null());
}

#[wasm_bindgen_test]
fn recalculate_orders_changes_by_sheet_row_col() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    wb.set_cell("A1".to_string(), JsValue::from_f64(10.0), Some("Sheet2".to_string()))
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), Some("Sheet2".to_string()))
        .unwrap();

    // Establish initial formula values.
    wb.recalculate(None).unwrap();

    // Dirty both sheets before a single recalc tick so ordering is deterministic.
    wb.set_cell("A1".to_string(), JsValue::from_f64(2.0), None)
        .unwrap();
    wb.set_cell("A1".to_string(), JsValue::from_f64(11.0), Some("Sheet2".to_string()))
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 2);
    assert_eq!(changes[0].sheet, "Sheet1");
    assert_eq!(changes[0].address, "A2");
    assert_json_number(&changes[0].value, 4.0);
    assert_eq!(changes[1].sheet, "Sheet2");
    assert_eq!(changes[1].address, "A2");
    assert_json_number(&changes[1].value, 22.0);
}

#[wasm_bindgen_test]
fn recalculate_orders_changes_by_row_col_within_sheet() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("B1".to_string(), JsValue::from_str("=A1+1"), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    wb.set_cell("A1".to_string(), JsValue::from_f64(2.0), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 2);
    // Row-major: B1 (row 0, col 1) comes before A2 (row 1, col 0).
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "B1");
    assert_json_number(&changes[0].value, 3.0);
    assert_eq!(changes[1].sheet, DEFAULT_SHEET);
    assert_eq!(changes[1].address, "A2");
    assert_json_number(&changes[1].value, 4.0);
}

#[wasm_bindgen_test]
fn recalculate_filters_changes_by_sheet_argument_case_insensitive() {
    let mut wb = WasmWorkbook::new();

    wb.set_cell("A1".to_string(), JsValue::from_f64(1.0), None)
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), None)
        .unwrap();

    wb.set_cell("A1".to_string(), JsValue::from_f64(10.0), Some("Sheet2".to_string()))
        .unwrap();
    wb.set_cell("A2".to_string(), JsValue::from_str("=A1*2"), Some("Sheet2".to_string()))
        .unwrap();

    wb.recalculate(None).unwrap();

    // Dirty both sheets, then request a sheet-scoped recalc. The returned changes should still
    // be filtered to the requested sheet name (case-insensitive).
    wb.set_cell("A1".to_string(), JsValue::from_f64(2.0), None)
        .unwrap();
    wb.set_cell("A1".to_string(), JsValue::from_f64(11.0), Some("Sheet2".to_string()))
        .unwrap();

    let changes_js = wb.recalculate(Some("sHeEt1".to_string())).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 1);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "A2");
    assert_json_number(&changes[0].value, 4.0);
}

#[wasm_bindgen_test]
fn recalculate_errors_on_missing_sheet_argument() {
    let mut wb = WasmWorkbook::new();
    let err = wb.recalculate(Some("MissingSheet".to_string())).unwrap_err();
    let msg = err.as_string().unwrap_or_default();
    assert!(msg.contains("missing sheet"), "unexpected error: {msg}");
}

#[wasm_bindgen_test]
fn recalculate_reports_formula_edit_to_blank_value() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_str("=1"), None)
        .unwrap();
    wb.recalculate(None).unwrap();

    wb.set_cell("A1".to_string(), JsValue::from_str("=A2"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(
        changes,
        vec![CellChange {
            sheet: DEFAULT_SHEET.to_string(),
            address: "A1".to_string(),
            value: JsonValue::Null,
        }]
    );
}

#[wasm_bindgen_test]
fn recalculate_reports_cleared_spill_outputs_after_edit() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell("A1".to_string(), JsValue::from_str("=SEQUENCE(1,3)"), None)
        .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 3);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "A1");
    assert_json_number(&changes[0].value, 1.0);
    assert_eq!(changes[1].sheet, DEFAULT_SHEET);
    assert_eq!(changes[1].address, "B1");
    assert_json_number(&changes[1].value, 2.0);
    assert_eq!(changes[2].sheet, DEFAULT_SHEET);
    assert_eq!(changes[2].address, "C1");
    assert_json_number(&changes[2].value, 3.0);

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
                sheet: DEFAULT_SHEET.to_string(),
                address: "A1".to_string(),
                value: json!("#SPILL!"),
            },
            CellChange {
                sheet: DEFAULT_SHEET.to_string(),
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
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, json!("=A1+B1"));
    assert_json_number(&cell.value, 3.0);
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
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_json_number(&a1.input, 1.0);
    assert_json_number(&a1.value, 1.0);

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.input, json!("Hello"));
    assert_eq!(b1.value, json!("Hello"));
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_loads_shared_strings_fixture() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/shared-strings.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();

    // Should not error even though the fixture contains only shared strings.
    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert!(changes.is_empty());

    let a1_js = wb.get_cell("A1".to_string(), None).unwrap();
    let a1: CellData = serde_wasm_bindgen::from_value(a1_js).unwrap();
    assert_eq!(a1.input, json!("Hello"));
    assert_eq!(a1.value, json!("Hello"));

    let b1_js = wb.get_cell("B1".to_string(), None).unwrap();
    let b1: CellData = serde_wasm_bindgen::from_value(b1_js).unwrap();
    assert_eq!(b1.input, json!("World"));
    assert_eq!(b1.value, json!("World"));
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_loads_shared_formula_fixture() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/formulas/shared-formula.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();
    wb.recalculate(None).unwrap();

    let a2_js = wb.get_cell("A2".to_string(), None).unwrap();
    let a2: CellData = serde_wasm_bindgen::from_value(a2_js).unwrap();
    assert_eq!(a2.input, json!("=B2*2"));
    assert_json_number(&a2.value, 4.0);

    // Ensure shared-formula cells behave like real formulas (not frozen literal cached values).
    wb.set_cell("B2".to_string(), JsValue::from_f64(10.0), None)
        .unwrap();
    wb.recalculate(None).unwrap();

    let a2_js = wb.get_cell("A2".to_string(), None).unwrap();
    let a2: CellData = serde_wasm_bindgen::from_value(a2_js).unwrap();
    assert_eq!(a2.input, json!("=B2*2"));
    assert_json_number(&a2.value, 20.0);
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_loads_multi_sheet_fixture() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/basic/multi-sheet.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();
    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert!(changes.is_empty());

    let sheet2_a1_js = wb
        .get_cell("A1".to_string(), Some("Sheet2".to_string()))
        .unwrap();
    let sheet2_a1: CellData = serde_wasm_bindgen::from_value(sheet2_a1_js).unwrap();
    assert_json_number(&sheet2_a1.value, 2.0);
}

#[wasm_bindgen_test]
fn from_xlsx_bytes_imports_defined_names() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../fixtures/xlsx/metadata/defined-names.xlsx"
    ));

    let mut wb = WasmWorkbook::from_xlsx_bytes(bytes).unwrap();
    wb.set_cell("C1".to_string(), JsValue::from_str("=ZedName"), None)
        .unwrap();

    wb.recalculate(None).unwrap();

    let cell_js = wb.get_cell("C1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, json!("=ZedName"));
    assert_eq!(cell.value, json!("Hello"));
}

#[wasm_bindgen_test]
fn cross_sheet_formulas_recalculate() {
    let mut wb = WasmWorkbook::new();
    wb.set_cell(
        "A1".to_string(),
        JsValue::from_f64(1.0),
        Some("Sheet1".to_string()),
    )
    .unwrap();
    wb.set_cell(
        "A1".to_string(),
        JsValue::from_str("=Sheet1!A1*2"),
        Some("Sheet2".to_string()),
    )
    .unwrap();

    let changes_js = wb.recalculate(None).unwrap();
    let changes: Vec<CellChange> = serde_wasm_bindgen::from_value(changes_js).unwrap();
    assert_eq!(changes.len(), 1);
    assert_eq!(changes[0].sheet, "Sheet2");
    assert_eq!(changes[0].address, "A1");
    assert_json_number(&changes[0].value, 2.0);

    let cell_js = wb
        .get_cell("A1".to_string(), Some("Sheet2".to_string()))
        .unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_json_number(&cell.value, 2.0);
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
    assert_eq!(changes.len(), 1);
    assert_eq!(changes[0].sheet, DEFAULT_SHEET);
    assert_eq!(changes[0].address, "A2");
    assert_json_number(&changes[0].value, 0.0);

    let cell_js = wb.get_cell("A1".to_string(), None).unwrap();
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, JsonValue::Null);
    assert_eq!(cell.value, JsonValue::Null);

    let exported = wb.to_json().unwrap();
    let parsed: JsonValue = serde_json::from_str(&exported).unwrap();
    let cells = parsed["sheets"][DEFAULT_SHEET]["cells"].as_object().unwrap();
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
    let cells = parsed["sheets"][DEFAULT_SHEET]["cells"].as_object().unwrap();

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
    let cell: CellData = serde_wasm_bindgen::from_value(cell_js).unwrap();
    assert_eq!(cell.input, JsonValue::Null);
    assert_eq!(cell.value, JsonValue::Null);

    let exported = wb.to_json().unwrap();
    let parsed: JsonValue = serde_json::from_str(&exported).unwrap();
    let cells = parsed["sheets"][DEFAULT_SHEET]["cells"].as_object().unwrap();
    assert!(!cells.contains_key("A1"));
}
