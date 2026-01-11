use formula_core::{CellChange, CellData, Workbook, WorkbookError, DEFAULT_SHEET};
use roxmltree::Document;
use serde_json::Value as JsonValue;
use std::io::{Cursor, Read};
use wasm_bindgen::prelude::*;
use zip::ZipArchive;

fn to_js_error(err: WorkbookError) -> JsValue {
    JsValue::from_str(&err.to_string())
}

fn to_js_display_error(err: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&err.to_string())
}

fn read_zip_entry<R: std::io::Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> Result<String, JsValue> {
    let mut file = archive.by_name(name).map_err(to_js_display_error)?;
    let mut buf = String::new();
    file.read_to_string(&mut buf).map_err(to_js_display_error)?;
    Ok(buf)
}

fn parse_shared_strings(xml: &str) -> Result<Vec<String>, JsValue> {
    let doc = Document::parse(xml).map_err(to_js_display_error)?;
    let mut out = Vec::new();

    for si in doc
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "si")
    {
        let text = si
            .descendants()
            .filter(|node| node.is_element() && node.tag_name().name() == "t")
            .filter_map(|node| node.text())
            .collect::<String>();
        out.push(text);
    }

    Ok(out)
}

fn parse_inline_string(cell: roxmltree::Node<'_, '_>) -> String {
    cell.descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "t")
        .filter_map(|node| node.text())
        .collect::<String>()
}

fn parse_cell_value(
    cell: roxmltree::Node<'_, '_>,
    shared_strings: &[String],
) -> Result<Option<JsonValue>, JsValue> {
    // Formula cells are represented by `<f>` and should be translated to Formula
    // inputs (`=...`) regardless of cached `<v>` values.
    if let Some(formula_node) = cell
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "f")
    {
        let formula = formula_node.text().unwrap_or_default();
        return Ok(Some(JsonValue::String(format!("={formula}"))));
    }

    let cell_type = cell.attribute("t");

    match cell_type {
        Some("inlineStr") => Ok(Some(JsonValue::String(parse_inline_string(cell)))),
        Some("s") => {
            let idx_text = cell
                .children()
                .find(|node| node.is_element() && node.tag_name().name() == "v")
                .and_then(|node| node.text())
                .unwrap_or_default();
            let idx: usize = idx_text.parse().map_err(to_js_display_error)?;
            let value = shared_strings.get(idx).cloned().unwrap_or_default();
            Ok(Some(JsonValue::String(value)))
        }
        Some("b") => {
            let raw = cell
                .children()
                .find(|node| node.is_element() && node.tag_name().name() == "v")
                .and_then(|node| node.text())
                .unwrap_or_default();
            Ok(Some(JsonValue::Bool(raw == "1")))
        }
        Some("str") => {
            let raw = cell
                .children()
                .find(|node| node.is_element() && node.tag_name().name() == "v")
                .and_then(|node| node.text())
                .unwrap_or_default();
            Ok(Some(JsonValue::String(raw.to_string())))
        }
        _ => {
            let raw = match cell
                .children()
                .find(|node| node.is_element() && node.tag_name().name() == "v")
                .and_then(|node| node.text())
            {
                Some(raw) => raw.trim(),
                None => return Ok(None),
            };

            let num: f64 = raw.parse().map_err(to_js_display_error)?;
            let json_num =
                serde_json::Number::from_f64(num).ok_or_else(|| JsValue::from_str("invalid number"))?;
            Ok(Some(JsonValue::Number(json_num)))
        }
    }
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
    pub fn from_xlsx_bytes(bytes: &[u8]) -> Result<WasmWorkbook, JsValue> {
        let cursor = Cursor::new(bytes);
        let mut archive = ZipArchive::new(cursor).map_err(to_js_display_error)?;

        let shared_strings = match read_zip_entry(&mut archive, "xl/sharedStrings.xml") {
            Ok(xml) => parse_shared_strings(&xml)?,
            Err(_) => Vec::new(),
        };

        let sheet_xml = read_zip_entry(&mut archive, "xl/worksheets/sheet1.xml")?;
        let sheet_doc = Document::parse(&sheet_xml).map_err(to_js_display_error)?;

        let mut wb = Workbook::new();

        for cell in sheet_doc
            .descendants()
            .filter(|node| node.is_element() && node.tag_name().name() == "c")
        {
            let Some(address) = cell.attribute("r") else {
                continue;
            };

            let Some(value) = parse_cell_value(cell, &shared_strings)? else {
                continue;
            };

            wb.set_cell(address, value, None).map_err(to_js_error)?;
        }

        Ok(WasmWorkbook { inner: wb })
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
