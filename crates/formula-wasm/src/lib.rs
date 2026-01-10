use formula_core::{CellChange, CellData, Workbook, WorkbookError, DEFAULT_SHEET};
use serde_json::Value as JsonValue;
use wasm_bindgen::prelude::*;

fn to_js_error(err: WorkbookError) -> JsValue {
    JsValue::from_str(&err.to_string())
}

#[wasm_bindgen]
pub struct WasmWorkbook {
    inner: Workbook,
}

#[wasm_bindgen]
impl WasmWorkbook {
    #[wasm_bindgen(constructor)]
    pub fn new() -> WasmWorkbook {
        WasmWorkbook {
            inner: Workbook::new(),
        }
    }

    #[wasm_bindgen(js_name = "fromJson")]
    pub fn from_json(json: &str) -> Result<WasmWorkbook, JsValue> {
        Ok(WasmWorkbook {
            inner: Workbook::from_json_str(json).map_err(to_js_error)?,
        })
    }

    #[wasm_bindgen(js_name = "fromXlsxBytes")]
    pub fn from_xlsx_bytes(_bytes: &[u8]) -> Result<WasmWorkbook, JsValue> {
        Err(JsValue::from_str(
            "loading xlsx bytes is not supported in WASM yet",
        ))
    }

    #[wasm_bindgen(js_name = "toJson")]
    pub fn to_json(&self) -> Result<String, JsValue> {
        self.inner.to_json_str().map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = "getCell")]
    pub fn get_cell(&self, address: String, sheet: Option<String>) -> Result<JsValue, JsValue> {
        let cell = self
            .inner
            .get_cell(&address, sheet.as_deref())
            .map_err(to_js_error)?;
        serde_wasm_bindgen::to_value(&cell).map_err(|err| JsValue::from_str(&err.to_string()))
    }

    #[wasm_bindgen(js_name = "setCell")]
    pub fn set_cell(
        &mut self,
        address: String,
        input: JsValue,
        sheet: Option<String>,
    ) -> Result<(), JsValue> {
        let input: JsonValue = serde_wasm_bindgen::from_value(input)
            .map_err(|err| JsValue::from_str(&err.to_string()))?;
        self.inner
            .set_cell(&address, input, sheet.as_deref())
            .map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = "getRange")]
    pub fn get_range(&self, range: String, sheet: Option<String>) -> Result<JsValue, JsValue> {
        let range = self
            .inner
            .get_range(&range, sheet.as_deref())
            .map_err(to_js_error)?;
        serde_wasm_bindgen::to_value(&range).map_err(|err| JsValue::from_str(&err.to_string()))
    }

    #[wasm_bindgen(js_name = "setRange")]
    pub fn set_range(
        &mut self,
        range: String,
        values: JsValue,
        sheet: Option<String>,
    ) -> Result<(), JsValue> {
        let values: Vec<Vec<JsonValue>> = serde_wasm_bindgen::from_value(values)
            .map_err(|err| JsValue::from_str(&err.to_string()))?;
        self.inner
            .set_range(&range, values, sheet.as_deref())
            .map_err(to_js_error)
    }

    #[wasm_bindgen(js_name = "recalculate")]
    pub fn recalculate(&mut self, sheet: Option<String>) -> Result<JsValue, JsValue> {
        let changes = self
            .inner
            .recalculate(sheet.as_deref())
            .map_err(to_js_error)?;
        serde_wasm_bindgen::to_value(&changes).map_err(|err| JsValue::from_str(&err.to_string()))
    }

    #[wasm_bindgen(js_name = "defaultSheetName")]
    pub fn default_sheet_name() -> String {
        DEFAULT_SHEET.to_string()
    }
}

// Re-export the DTO types for consumers (tests, TS generator tooling, etc).
pub use formula_core::{CellChange as CoreCellChange, CellData as CoreCellData};

#[allow(dead_code)]
fn _assert_dto_serializable() {
    fn assert_serde<T: serde::Serialize + for<'de> serde::Deserialize<'de>>() {}
    assert_serde::<CellData>();
    assert_serde::<CellChange>();
}
