use std::collections::{BTreeMap, BTreeSet, HashMap};

use formula_core::{CellChange, CellData, DEFAULT_SHEET};
use formula_engine::{Engine, ErrorKind, NameDefinition, NameScope, Value as EngineValue};
use formula_model::{display_formula_text, CellRef, CellValue, DateSystem, DefinedNameScope, Range};
use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;
use wasm_bindgen::prelude::*;

fn js_err(message: impl ToString) -> JsValue {
    JsValue::from_str(&message.to_string())
}

#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
struct FormulaCellKey {
    sheet: String,
    row: u32,
    col: u32,
}

impl FormulaCellKey {
    fn new(sheet: String, cell: CellRef) -> Self {
        Self {
            sheet,
            row: cell.row,
            col: cell.col,
        }
    }

    fn address(&self) -> String {
        CellRef::new(self.row, self.col).to_a1()
    }
}

fn is_scalar_json(value: &JsonValue) -> bool {
    matches!(
        value,
        JsonValue::Null | JsonValue::Bool(_) | JsonValue::Number(_) | JsonValue::String(_)
    )
}

fn is_formula_input(value: &JsonValue) -> bool {
    value
        .as_str()
        .is_some_and(|s| s.trim_start().starts_with('='))
}

fn normalize_sheet_key(name: &str) -> String {
    name.to_ascii_uppercase()
}

fn json_to_engine_value(value: &JsonValue) -> EngineValue {
    match value {
        JsonValue::Null => EngineValue::Blank,
        JsonValue::Bool(b) => EngineValue::Bool(*b),
        JsonValue::Number(n) => EngineValue::Number(n.as_f64().unwrap_or(0.0)),
        JsonValue::String(s) => EngineValue::Text(s.clone()),
        JsonValue::Array(_) | JsonValue::Object(_) => {
            // Should be unreachable due to `is_scalar_json` validation, but keep a fallback.
            EngineValue::Blank
        }
    }
}

fn engine_value_to_json(value: EngineValue) -> JsonValue {
    match value {
        EngineValue::Blank => JsonValue::Null,
        EngineValue::Bool(b) => JsonValue::Bool(b),
        EngineValue::Text(s) => JsonValue::String(s),
        EngineValue::Number(n) => serde_json::Number::from_f64(n)
            .map(JsonValue::Number)
            .unwrap_or_else(|| JsonValue::String(ErrorKind::Num.as_code().to_string())),
        EngineValue::Error(kind) => JsonValue::String(kind.as_code().to_string()),
        // The JS protocol only supports scalar-ish values. Spill markers should not leak because
        // `Engine::get_cell_value` resolves spill cells to their concrete values. Keep a defensive
        // fallback anyway.
        EngineValue::Array(_) | EngineValue::Spill { .. } => {
            JsonValue::String(ErrorKind::Spill.as_code().to_string())
        }
    }
}

fn cell_value_to_engine(value: &CellValue) -> EngineValue {
    match value {
        CellValue::Empty => EngineValue::Blank,
        CellValue::Number(n) => EngineValue::Number(*n),
        CellValue::String(s) => EngineValue::Text(s.clone()),
        CellValue::Boolean(b) => EngineValue::Bool(*b),
        CellValue::Error(err) => match err {
            formula_model::ErrorValue::Null => EngineValue::Error(ErrorKind::Null),
            formula_model::ErrorValue::Div0 => EngineValue::Error(ErrorKind::Div0),
            formula_model::ErrorValue::Value => EngineValue::Error(ErrorKind::Value),
            formula_model::ErrorValue::Ref => EngineValue::Error(ErrorKind::Ref),
            formula_model::ErrorValue::Name => EngineValue::Error(ErrorKind::Name),
            formula_model::ErrorValue::Num => EngineValue::Error(ErrorKind::Num),
            formula_model::ErrorValue::NA => EngineValue::Error(ErrorKind::NA),
            formula_model::ErrorValue::Spill => EngineValue::Error(ErrorKind::Spill),
            formula_model::ErrorValue::Calc => EngineValue::Error(ErrorKind::Calc),
            other => EngineValue::Text(other.as_str().to_string()),
        },
        CellValue::RichText(rt) => EngineValue::Text(rt.plain_text().to_string()),
        // The workbook model can store cached array/spill results, but the WASM worker API only
        // supports scalar values today. Treat these as spill errors so downstream formulas see an
        // error rather than silently treating an array as a string.
        CellValue::Array(_) | CellValue::Spill(_) => EngineValue::Error(ErrorKind::Spill),
    }
}

fn cell_value_to_json(value: &CellValue) -> JsonValue {
    engine_value_to_json(cell_value_to_engine(value))
}

#[derive(Default)]
struct WorkbookState {
    engine: Engine,
    /// Workbook input state for `toJson`/`getCell.input`.
    ///
    /// Mirrors the simple JSON workbook schema consumed by `packages/engine`.
    sheets: BTreeMap<String, BTreeMap<String, JsonValue>>,
    /// Case-insensitive mapping (Excel semantics) from sheet key -> display name.
    sheet_lookup: HashMap<String, String>,
    /// All formula-bearing cells (sheet, A1) in deterministic order.
    formula_cells: BTreeSet<FormulaCellKey>,
    /// Last reported value for each formula cell.
    ///
    /// Used to compute deterministic `CellChange[]` results from `recalculate()`.
    last_formula_values: BTreeMap<FormulaCellKey, JsonValue>,
}

impl WorkbookState {
    fn new_empty() -> Self {
        Self {
            engine: Engine::new(),
            sheets: BTreeMap::new(),
            sheet_lookup: HashMap::new(),
            formula_cells: BTreeSet::new(),
            last_formula_values: BTreeMap::new(),
        }
    }

    fn new_with_default_sheet() -> Self {
        let mut wb = Self::new_empty();
        wb.ensure_sheet(DEFAULT_SHEET);
        wb
    }

    fn ensure_sheet(&mut self, name: &str) -> String {
        let key = normalize_sheet_key(name);
        if let Some(existing) = self.sheet_lookup.get(&key) {
            return existing.clone();
        }

        let display = name.to_string();
        self.sheet_lookup.insert(key, display.clone());
        self.sheets.entry(display.clone()).or_default();
        self.engine.ensure_sheet(&display);
        display
    }

    fn resolve_sheet(&self, name: &str) -> Option<&str> {
        let key = normalize_sheet_key(name);
        self.sheet_lookup.get(&key).map(String::as_str)
    }

    fn require_sheet(&self, name: &str) -> Result<&str, JsValue> {
        self.resolve_sheet(name)
            .ok_or_else(|| js_err(format!("missing sheet: {name}")))
    }

    fn parse_address(address: &str) -> Result<CellRef, JsValue> {
        CellRef::from_a1(address).map_err(|_| js_err(format!("invalid cell address: {address}")))
    }

    fn parse_range(range: &str) -> Result<Range, JsValue> {
        Range::from_a1(range).map_err(|_| js_err(format!("invalid range: {range}")))
    }

    fn set_cell_internal(
        &mut self,
        sheet: &str,
        address: &str,
        input: JsonValue,
    ) -> Result<(), JsValue> {
        if !is_scalar_json(&input) {
            return Err(js_err(format!("invalid cell value: {address}")));
        }

        let sheet = self.ensure_sheet(sheet);
        let cell_ref = Self::parse_address(address)?;
        let address = cell_ref.to_a1();

        let sheet_cells = self
            .sheets
            .get_mut(&sheet)
            .expect("sheet just ensured must exist");

        if is_formula_input(&input) {
            let raw = input.as_str().expect("formula input must be string");
            let canonical = raw.trim_start().to_string();

            // Reset the stored value to blank so `getCell` returns null until the next recalc,
            // matching the existing `formula-core` semantics.
            self.engine
                .set_cell_value(&sheet, &address, EngineValue::Blank)
                .map_err(|err| js_err(err.to_string()))?;
            self.engine
                .set_cell_formula(&sheet, &address, &canonical)
                .map_err(|err| js_err(err.to_string()))?;

            sheet_cells.insert(address.clone(), JsonValue::String(canonical));

            let key = FormulaCellKey::new(sheet.clone(), cell_ref);
            self.formula_cells.insert(key.clone());
            self.last_formula_values.insert(key, JsonValue::Null);
            return Ok(());
        }

        // Non-formula scalar value.
        self.engine
            .set_cell_value(&sheet, &address, json_to_engine_value(&input))
            .map_err(|err| js_err(err.to_string()))?;

        sheet_cells.insert(address.clone(), input);

        let key = FormulaCellKey::new(sheet.clone(), cell_ref);
        self.formula_cells.remove(&key);
        self.last_formula_values.remove(&key);
        Ok(())
    }

    fn get_cell_data(&self, sheet: &str, address: &str) -> Result<CellData, JsValue> {
        let sheet = self.require_sheet(sheet)?.to_string();
        let address = Self::parse_address(address)?.to_a1();

        let input = self
            .sheets
            .get(&sheet)
            .and_then(|cells| cells.get(&address))
            .cloned()
            .unwrap_or(JsonValue::Null);

        let value = engine_value_to_json(self.engine.get_cell_value(&sheet, &address));

        Ok(CellData {
            sheet,
            address,
            input,
            value,
        })
    }

    fn recalculate_internal(&mut self, sheet: Option<&str>) -> Result<Vec<CellChange>, JsValue> {
        if let Some(sheet) = sheet {
            self.require_sheet(sheet)?;
        }

        self.engine.recalculate_single_threaded();

        let sheet_filter = sheet.and_then(|s| self.resolve_sheet(s)).map(str::to_string);

        let mut changes = Vec::new();
        for key in &self.formula_cells {
            let sheet_name = &key.sheet;
            if let Some(filter) = &sheet_filter {
                if sheet_name != filter {
                    continue;
                }
            }

            let address = key.address();
            let new_value = engine_value_to_json(self.engine.get_cell_value(sheet_name, &address));
            let old_value = self
                .last_formula_values
                .get(key)
                .cloned()
                .unwrap_or(JsonValue::Null);

            if old_value != new_value {
                changes.push(CellChange {
                    sheet: sheet_name.clone(),
                    address: address.clone(),
                    value: new_value.clone(),
                });
                self.last_formula_values.insert(key.clone(), new_value);
            }
        }

        Ok(changes)
    }
}

#[wasm_bindgen]
pub struct WasmWorkbook {
    inner: WorkbookState,
}

#[wasm_bindgen]
impl WasmWorkbook {
    #[wasm_bindgen(constructor)]
    pub fn new() -> WasmWorkbook {
        WasmWorkbook {
            inner: WorkbookState::new_with_default_sheet(),
        }
    }

    #[wasm_bindgen(js_name = "fromJson")]
    pub fn from_json(json: &str) -> Result<WasmWorkbook, JsValue> {
        #[derive(Debug, Deserialize)]
        struct WorkbookJson {
            sheets: BTreeMap<String, SheetJson>,
        }

        #[derive(Debug, Deserialize)]
        struct SheetJson {
            cells: BTreeMap<String, JsonValue>,
        }

        let parsed: WorkbookJson = serde_json::from_str(json)
            .map_err(|err| js_err(format!("invalid workbook json: {}", err)))?;

        let mut wb = WorkbookState::new_empty();

        // Create all sheets up-front so cross-sheet formula references resolve correctly.
        for sheet_name in parsed.sheets.keys() {
            wb.ensure_sheet(sheet_name);
        }

        for (sheet_name, sheet) in parsed.sheets {
            for (address, input) in sheet.cells {
                if !is_scalar_json(&input) {
                    return Err(js_err(format!("invalid cell value: {address}")));
                }
                wb.set_cell_internal(&sheet_name, &address, input)?;
            }
        }

        if wb.sheets.is_empty() {
            wb.ensure_sheet(DEFAULT_SHEET);
        }

        Ok(WasmWorkbook { inner: wb })
    }

    #[wasm_bindgen(js_name = "fromXlsxBytes")]
    pub fn from_xlsx_bytes(bytes: &[u8]) -> Result<WasmWorkbook, JsValue> {
        let model = formula_xlsx::read_workbook_model_from_bytes(bytes)
            .map_err(|err| js_err(err.to_string()))?;

        let mut wb = WorkbookState::new_empty();

        // Date system influences date serials for NOW/TODAY/DATE, etc.
        wb.engine.set_date_system(match model.date_system {
            DateSystem::Excel1900 => formula_engine::date::ExcelDateSystem::EXCEL_1900,
            DateSystem::Excel1904 => formula_engine::date::ExcelDateSystem::Excel1904,
        });

        // Create all sheets up-front so formulas can resolve cross-sheet references.
        for sheet in &model.sheets {
            wb.ensure_sheet(&sheet.name);
        }

        // Best-effort defined names.
        let mut sheet_names_by_id: HashMap<u32, String> = HashMap::new();
        for sheet in &model.sheets {
            sheet_names_by_id.insert(sheet.id, sheet.name.clone());
        }

        for name in &model.defined_names {
            let scope = match name.scope {
                DefinedNameScope::Workbook => NameScope::Workbook,
                DefinedNameScope::Sheet(sheet_id) => {
                    let Some(sheet_name) = sheet_names_by_id.get(&sheet_id) else {
                        continue;
                    };
                    NameScope::Sheet(sheet_name)
                }
            };

            let refers_to = name.refers_to.trim();
            if refers_to.is_empty() {
                continue;
            }

            // Best-effort heuristic:
            // - numeric/bool constants are imported as constants
            // - everything else is imported as a reference-like expression
            let definition = if refers_to.eq_ignore_ascii_case("TRUE") {
                NameDefinition::Constant(EngineValue::Bool(true))
            } else if refers_to.eq_ignore_ascii_case("FALSE") {
                NameDefinition::Constant(EngineValue::Bool(false))
            } else if let Ok(n) = refers_to.parse::<f64>() {
                NameDefinition::Constant(EngineValue::Number(n))
            } else {
                NameDefinition::Reference(refers_to.to_string())
            };

            let _ = wb.engine.define_name(&name.name, scope, definition);
        }

        for sheet in &model.sheets {
            let sheet_name = wb
                .resolve_sheet(&sheet.name)
                .expect("sheet just ensured must resolve")
                .to_string();

            for (cell_ref, cell) in sheet.iter_cells() {
                let address = cell_ref.to_a1();

                // Skip style-only cells (not representable in this WASM DTO surface).
                let has_formula = cell.formula.is_some();
                let has_value = !cell.value.is_empty();
                if !has_formula && !has_value {
                    continue;
                }

                // Seed cached values first (including cached formula results).
                wb.engine
                    .set_cell_value(&sheet_name, &address, cell_value_to_engine(&cell.value))
                    .map_err(|err| js_err(err.to_string()))?;

                if let Some(formula) = cell.formula.as_deref() {
                    // `formula-model` stores formulas without a leading '='.
                    let display = display_formula_text(formula);
                    if !display.is_empty() {
                        // Best-effort: if the formula fails to parse (unsupported syntax), leave the
                        // cached value and still store the display formula in the input map.
                        let _ = wb.engine.set_cell_formula(&sheet_name, &address, &display);

                        let sheet_cells = wb
                            .sheets
                            .get_mut(&sheet_name)
                            .expect("sheet just ensured must exist");
                        sheet_cells.insert(address.clone(), JsonValue::String(display));

                        let key = FormulaCellKey::new(sheet_name.clone(), cell_ref);
                        wb.formula_cells.insert(key.clone());
                        wb.last_formula_values.insert(key, cell_value_to_json(&cell.value));
                        continue;
                    }
                }

                // Non-formula cell; store scalar value as input.
                let sheet_cells = wb
                    .sheets
                    .get_mut(&sheet_name)
                    .expect("sheet just ensured must exist");
                sheet_cells.insert(address, cell_value_to_json(&cell.value));
            }
        }

        if wb.sheets.is_empty() {
            wb.ensure_sheet(DEFAULT_SHEET);
        }

        Ok(WasmWorkbook { inner: wb })
    }

    #[wasm_bindgen(js_name = "toJson")]
    pub fn to_json(&self) -> Result<String, JsValue> {
        #[derive(Serialize)]
        struct WorkbookJson<'a> {
            sheets: &'a BTreeMap<String, BTreeMap<String, JsonValue>>,
        }

        serde_json::to_string(&WorkbookJson {
            sheets: &self.inner.sheets,
        })
        .map_err(|err| js_err(format!("invalid workbook json: {}", err)))
    }

    #[wasm_bindgen(js_name = "getCell")]
    pub fn get_cell(&self, address: String, sheet: Option<String>) -> Result<JsValue, JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        let cell = self.inner.get_cell_data(sheet, &address)?;
        serde_wasm_bindgen::to_value(&cell).map_err(|err| js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "setCell")]
    pub fn set_cell(
        &mut self,
        address: String,
        input: JsValue,
        sheet: Option<String>,
    ) -> Result<(), JsValue> {
        let input: JsonValue =
            serde_wasm_bindgen::from_value(input).map_err(|err| js_err(err.to_string()))?;
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        self.inner.set_cell_internal(sheet, &address, input)
    }

    #[wasm_bindgen(js_name = "getRange")]
    pub fn get_range(&self, range: String, sheet: Option<String>) -> Result<JsValue, JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        let sheet = self.inner.require_sheet(sheet)?.to_string();
        let range = WorkbookState::parse_range(&range)?;

        let mut rows: Vec<Vec<CellData>> = Vec::new();
        for row in range.start.row..=range.end.row {
            let mut cols = Vec::new();
            for col in range.start.col..=range.end.col {
                let addr = CellRef::new(row, col).to_a1();
                cols.push(self.inner.get_cell_data(&sheet, &addr)?);
            }
            rows.push(cols);
        }

        serde_wasm_bindgen::to_value(&rows).map_err(|err| js_err(err.to_string()))
    }

    #[wasm_bindgen(js_name = "setRange")]
    pub fn set_range(
        &mut self,
        range: String,
        values: JsValue,
        sheet: Option<String>,
    ) -> Result<(), JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        let range_parsed = WorkbookState::parse_range(&range)?;

        let values: Vec<Vec<JsonValue>> =
            serde_wasm_bindgen::from_value(values).map_err(|err| js_err(err.to_string()))?;

        let expected_rows = range_parsed.height() as usize;
        let expected_cols = range_parsed.width() as usize;
        if values.len() != expected_rows || values.iter().any(|row| row.len() != expected_cols) {
            return Err(js_err(format!(
                "invalid range: range {range} expects {expected_rows}x{expected_cols} values"
            )));
        }

        for (r_idx, row_values) in values.into_iter().enumerate() {
            for (c_idx, input) in row_values.into_iter().enumerate() {
                let row = range_parsed.start.row + r_idx as u32;
                let col = range_parsed.start.col + c_idx as u32;
                let addr = CellRef::new(row, col).to_a1();
                self.inner.set_cell_internal(sheet, &addr, input)?;
            }
        }

        Ok(())
    }

    #[wasm_bindgen(js_name = "recalculate")]
    pub fn recalculate(&mut self, sheet: Option<String>) -> Result<JsValue, JsValue> {
        let changes = self.inner.recalculate_internal(sheet.as_deref())?;
        serde_wasm_bindgen::to_value(&changes).map_err(|err| js_err(err.to_string()))
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
