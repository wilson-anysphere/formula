use std::collections::{BTreeMap, BTreeSet, HashMap};

use formula_core::{CellChange, CellData, DEFAULT_SHEET};
use formula_engine::{Engine, ErrorKind, NameDefinition, NameScope, Value as EngineValue};
use formula_model::{
    display_formula_text, CellRef, CellValue, DateSystem, DefinedNameScope, Range,
};
use js_sys::{Array, Object, Reflect};
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
        // LAMBDA values are valid Excel scalars but cannot be represented in the current
        // worker JSON protocol. Use a descriptive placeholder so the UI does not crash
        // when a formula returns a lambda.
        EngineValue::Lambda(_) => JsonValue::String("<LAMBDA>".to_string()),
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
    /// Spill cells that were cleared by edits since the last recalc.
    ///
    /// `Engine::recalculate_with_value_changes` can only diff values across a recalc tick; when a
    /// spill is cleared as part of `setCell`/`setRange` we stash the affected cells so the next
    /// `recalculate()` call can return `CellChange[]` entries that blank out any now-stale spill
    /// outputs in the JS cache.
    pending_spill_clears: BTreeSet<FormulaCellKey>,
    /// Formula cells that were edited since the last recalc, keyed by their previous visible value.
    ///
    /// The JS frontend applies `directChange` updates for literal edits but not for formulas; the
    /// WASM bridge resets formula cells to blank until the next `recalculate()` so `getCell` matches
    /// the existing `formula-core` semantics. This can hide "value cleared" edits when the new
    /// formula result is also blank, so we keep the previous value here and explicitly diff it
    /// against the post-recalc value.
    pending_formula_baselines: BTreeMap<FormulaCellKey, JsonValue>,
}

impl WorkbookState {
    fn new_empty() -> Self {
        Self {
            engine: Engine::new(),
            sheets: BTreeMap::new(),
            sheet_lookup: HashMap::new(),
            pending_spill_clears: BTreeSet::new(),
            pending_formula_baselines: BTreeMap::new(),
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

        if let Some((origin, end)) = self.engine.spill_range(&sheet, &address) {
            let edited_row = cell_ref.row;
            let edited_col = cell_ref.col;
            let edited_is_formula = is_formula_input(&input);
            for row in origin.row..=end.row {
                for col in origin.col..=end.col {
                    // Skip the origin cell (top-left); we only need to clear spill outputs.
                    if row == origin.row && col == origin.col {
                        continue;
                    }
                    // If the user overwrote a spill output cell with a literal value, don't emit a
                    // spill-clear change for that cell; the caller already knows its new input.
                    if !edited_is_formula && row == edited_row && col == edited_col {
                        continue;
                    }
                    self.pending_spill_clears
                        .insert(FormulaCellKey::new(sheet.clone(), CellRef::new(row, col)));
                }
            }
        }

        let sheet_cells = self
            .sheets
            .get_mut(&sheet)
            .expect("sheet just ensured must exist");

        // `null` represents an empty cell in the JS protocol. Preserve sparse semantics by
        // removing the stored entry instead of storing an explicit blank.
        if input.is_null() {
            self.engine
                .clear_cell(&sheet, &address)
                .map_err(|err| js_err(err.to_string()))?;

            sheet_cells.remove(&address);
            // If this cell was previously tracked as part of a spill-clear batch, drop it so we
            // don't report direct input edits as recalc changes.
            self.pending_spill_clears
                .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
            self.pending_formula_baselines
                .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
            return Ok(());
        }

        if is_formula_input(&input) {
            let raw = input.as_str().expect("formula input must be string");
            let canonical = raw.trim_start().to_string();

            let key = FormulaCellKey::new(sheet.clone(), cell_ref);
            self.pending_formula_baselines
                .entry(key)
                .or_insert_with(|| {
                    engine_value_to_json(self.engine.get_cell_value(&sheet, &address))
                });

            // Reset the stored value to blank so `getCell` returns null until the next recalc,
            // matching the existing `formula-core` semantics.
            self.engine
                .set_cell_value(&sheet, &address, EngineValue::Blank)
                .map_err(|err| js_err(err.to_string()))?;
            self.engine
                .set_cell_formula(&sheet, &address, &canonical)
                .map_err(|err| js_err(err.to_string()))?;

            sheet_cells.insert(address.clone(), JsonValue::String(canonical));
            return Ok(());
        }

        // Non-formula scalar value.
        self.engine
            .set_cell_value(&sheet, &address, json_to_engine_value(&input))
            .map_err(|err| js_err(err.to_string()))?;

        sheet_cells.insert(address.clone(), input);
        // If this cell was previously tracked as part of a spill-clear batch (e.g. a multi-cell
        // paste over a spill range), drop it so we don't report direct input edits as recalc
        // changes.
        self.pending_spill_clears
            .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
        self.pending_formula_baselines
            .remove(&FormulaCellKey::new(sheet.clone(), cell_ref));
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

        let sheet_filter = sheet
            .and_then(|s| self.resolve_sheet(s))
            .map(str::to_string);

        let recalc_changes = self.engine.recalculate_with_value_changes_single_threaded();
        let mut by_cell: BTreeMap<FormulaCellKey, JsonValue> = BTreeMap::new();

        for change in recalc_changes {
            if let Some(filter) = &sheet_filter {
                if &change.sheet != filter {
                    continue;
                }
            }
            by_cell.insert(
                FormulaCellKey {
                    sheet: change.sheet,
                    row: change.addr.row,
                    col: change.addr.col,
                },
                engine_value_to_json(change.value),
            );
        }

        if let Some(filter) = &sheet_filter {
            let keys: Vec<FormulaCellKey> = self
                .pending_spill_clears
                .iter()
                .filter(|k| &k.sheet == filter)
                .cloned()
                .collect();
            for key in keys {
                self.pending_spill_clears.remove(&key);
                if by_cell.contains_key(&key) {
                    continue;
                }
                let address = key.address();
                let value = engine_value_to_json(self.engine.get_cell_value(&key.sheet, &address));
                by_cell.insert(key, value);
            }
        } else {
            let pending = std::mem::take(&mut self.pending_spill_clears);
            for key in pending {
                if by_cell.contains_key(&key) {
                    continue;
                }
                let address = key.address();
                let value = engine_value_to_json(self.engine.get_cell_value(&key.sheet, &address));
                by_cell.insert(key, value);
            }
        }

        if let Some(filter) = &sheet_filter {
            let keys: Vec<FormulaCellKey> = self
                .pending_formula_baselines
                .keys()
                .filter(|k| &k.sheet == filter)
                .cloned()
                .collect();
            for key in keys {
                let Some(before) = self.pending_formula_baselines.remove(&key) else {
                    continue;
                };
                if by_cell.contains_key(&key) {
                    continue;
                }
                let address = key.address();
                let after =
                    engine_value_to_json(self.engine.get_cell_value(&key.sheet, &address));
                if after != before {
                    by_cell.insert(key, after);
                }
            }
        } else {
            let pending = std::mem::take(&mut self.pending_formula_baselines);
            for (key, before) in pending {
                if by_cell.contains_key(&key) {
                    continue;
                }
                let address = key.address();
                let after =
                    engine_value_to_json(self.engine.get_cell_value(&key.sheet, &address));
                if after != before {
                    by_cell.insert(key, after);
                }
            }
        }

        let changes = by_cell
            .into_iter()
            .map(|(key, value)| {
                let address = key.address();
                CellChange {
                    sheet: key.sheet,
                    address,
                    value,
                }
            })
            .collect();

        Ok(changes)
    }
}

fn json_scalar_to_js(value: &JsonValue) -> JsValue {
    match value {
        JsonValue::Null => JsValue::NULL,
        JsonValue::Bool(b) => JsValue::from_bool(*b),
        JsonValue::Number(n) => n.as_f64().map(JsValue::from_f64).unwrap_or(JsValue::NULL),
        JsonValue::String(s) => JsValue::from_str(s),
        // The engine protocol only supports scalars; fall back to `null` for any
        // unexpected values to avoid surfacing `undefined`.
        _ => JsValue::NULL,
    }
}

fn object_set(obj: &Object, key: &str, value: &JsValue) -> Result<(), JsValue> {
    Reflect::set(obj, &JsValue::from_str(key), value).map(|_| ())
}

fn cell_data_to_js(cell: &CellData) -> Result<JsValue, JsValue> {
    let obj = Object::new();
    object_set(&obj, "sheet", &JsValue::from_str(&cell.sheet))?;
    object_set(&obj, "address", &JsValue::from_str(&cell.address))?;
    object_set(&obj, "input", &json_scalar_to_js(&cell.input))?;
    object_set(&obj, "value", &json_scalar_to_js(&cell.value))?;
    Ok(obj.into())
}

fn cell_change_to_js(change: &CellChange) -> Result<JsValue, JsValue> {
    let obj = Object::new();
    object_set(&obj, "sheet", &JsValue::from_str(&change.sheet))?;
    object_set(&obj, "address", &JsValue::from_str(&change.address))?;
    object_set(&obj, "value", &json_scalar_to_js(&change.value))?;
    Ok(obj.into())
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
                if input.is_null() {
                    // `null` cells are treated as absent (sparse semantics).
                    continue;
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
        struct WorkbookJson {
            sheets: BTreeMap<String, SheetJson>,
        }

        #[derive(Serialize)]
        struct SheetJson {
            cells: BTreeMap<String, JsonValue>,
        }

        let mut sheets = BTreeMap::new();
        for (sheet_name, cells) in &self.inner.sheets {
            let mut out_cells = BTreeMap::new();
            for (address, input) in cells {
                // Ensure we never serialize explicit `null` cells; empty cells are
                // omitted from the sparse workbook representation.
                if input.is_null() {
                    continue;
                }
                out_cells.insert(address.clone(), input.clone());
            }
            sheets.insert(sheet_name.clone(), SheetJson { cells: out_cells });
        }

        serde_json::to_string(&WorkbookJson { sheets })
            .map_err(|err| js_err(format!("invalid workbook json: {}", err)))
    }

    #[wasm_bindgen(js_name = "getCell")]
    pub fn get_cell(&self, address: String, sheet: Option<String>) -> Result<JsValue, JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        let cell = self.inner.get_cell_data(sheet, &address)?;
        cell_data_to_js(&cell)
    }

    #[wasm_bindgen(js_name = "setCell")]
    pub fn set_cell(
        &mut self,
        address: String,
        input: JsValue,
        sheet: Option<String>,
    ) -> Result<(), JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        if input.is_null() {
            return self
                .inner
                .set_cell_internal(sheet, &address, JsonValue::Null);
        }
        let input: JsonValue =
            serde_wasm_bindgen::from_value(input).map_err(|err| js_err(err.to_string()))?;
        self.inner.set_cell_internal(sheet, &address, input)
    }

    #[wasm_bindgen(js_name = "getRange")]
    pub fn get_range(&self, range: String, sheet: Option<String>) -> Result<JsValue, JsValue> {
        let sheet = sheet.as_deref().unwrap_or(DEFAULT_SHEET);
        let sheet = self.inner.require_sheet(sheet)?.to_string();
        let range = WorkbookState::parse_range(&range)?;

        let outer = Array::new();
        for row in range.start.row..=range.end.row {
            let inner = Array::new();
            for col in range.start.col..=range.end.col {
                let addr = CellRef::new(row, col).to_a1();
                let cell = self.inner.get_cell_data(&sheet, &addr)?;
                inner.push(&cell_data_to_js(&cell)?);
            }
            outer.push(&inner);
        }

        Ok(outer.into())
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
        let out = Array::new();
        for change in changes {
            out.push(&cell_change_to_js(&change)?);
        }
        Ok(out.into())
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

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn recalculate_includes_spill_output_cells() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=SEQUENCE(1,2)"))
            .unwrap();

        let changes = wb.recalculate_internal(None).unwrap();
        assert_eq!(
            changes,
            vec![
                CellChange {
                    sheet: DEFAULT_SHEET.to_string(),
                    address: "A1".to_string(),
                    value: json!(1.0),
                },
                CellChange {
                    sheet: DEFAULT_SHEET.to_string(),
                    address: "B1".to_string(),
                    value: json!(2.0),
                },
            ]
        );
    }

    #[test]
    fn recalculate_reports_spill_clears_when_spill_origin_is_edited() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=SEQUENCE(1,2)"))
            .unwrap();
        let _ = wb.recalculate_internal(None).unwrap();

        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=1"))
            .unwrap();
        let changes = wb.recalculate_internal(None).unwrap();
        assert_eq!(
            changes,
            vec![
                CellChange {
                    sheet: DEFAULT_SHEET.to_string(),
                    address: "A1".to_string(),
                    value: json!(1.0),
                },
                CellChange {
                    sheet: DEFAULT_SHEET.to_string(),
                    address: "B1".to_string(),
                    value: JsonValue::Null,
                },
            ]
        );
    }

    #[test]
    fn recalculate_reports_spill_clears_when_spill_cell_is_overwritten() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=SEQUENCE(1,3)"))
            .unwrap();
        let _ = wb.recalculate_internal(None).unwrap();

        wb.set_cell_internal(DEFAULT_SHEET, "B1", json!(5.0))
            .unwrap();
        let changes = wb.recalculate_internal(None).unwrap();
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

    #[test]
    fn recalculate_reports_formula_edit_to_blank_value() {
        let mut wb = WorkbookState::new_with_default_sheet();
        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=1")).unwrap();
        let _ = wb.recalculate_internal(None).unwrap();

        wb.set_cell_internal(DEFAULT_SHEET, "A1", json!("=A2")).unwrap();
        let changes = wb.recalculate_internal(None).unwrap();
        assert_eq!(
            changes,
            vec![CellChange {
                sheet: DEFAULT_SHEET.to_string(),
                address: "A1".to_string(),
                value: JsonValue::Null,
            }]
        );
    }

    #[test]
    #[cfg(not(target_arch = "wasm32"))]
    fn to_json_preserves_engine_workbook_schema() {
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

        assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A1"], json!(1.0));
        assert_eq!(parsed["sheets"]["Sheet1"]["cells"]["A2"], json!("=A1*2"));

        let wb2 = WasmWorkbook::from_json(&json_str).unwrap();
        let json_str2 = wb2.to_json().unwrap();
        let parsed2: serde_json::Value = serde_json::from_str(&json_str2).unwrap();
        assert_eq!(parsed2["sheets"]["Sheet1"]["cells"]["A2"], json!("=A1*2"));
    }
}
