use chrono::NaiveDate;

use formula_model::Range;

use crate::eval::CellAddr;
use crate::pivot::source::coerce_pivot_value_with_number_format;
use crate::pivot::{PivotCache, PivotConfig, PivotEngine, PivotError, PivotResult, PivotValue};
use crate::value::{ErrorKind, Value};

use super::{CellKey, Engine, SheetId};

impl Engine {
    /// Build a [`PivotCache`] directly from the engine's current workbook state.
    ///
    /// This is similar to [`Engine::calculate_pivot_from_range`], but returns the cache so callers
    /// can inspect schema/unique values or run calculations separately.
    pub fn pivot_cache_from_range(
        &self,
        sheet: &str,
        range: Range,
    ) -> Result<PivotCache, PivotError> {
        let sheet_id = self
            .workbook
            .sheet_id(sheet)
            .ok_or_else(|| PivotError::SheetNotFound(sheet.to_string()))?;

        let source = materialize_range_as_pivot_values(self, sheet_id, range);
        PivotCache::from_range(&source)
    }

    /// Calculate a pivot table directly from the engine's current workbook state.
    ///
    /// This avoids marshalling the source range through JS/IPC: the engine materializes the
    /// rectangular `range` into a `Vec<Vec<PivotValue>>`, builds a [`PivotCache`], and then runs
    /// [`PivotEngine::calculate`].
    ///
    /// Notes:
    /// - The first row of `range` is treated as the header row.
    /// - Headers are derived from the cell's **current value** (not the formatted display
    ///   string). Field names use the pivot cache's internal `PivotValue` → display-name
    ///   conversion logic.
    /// - All values are read via the same code path as [`Engine::get_cell_value`] (including spill
    ///   resolution and external value providers), but without A1 parsing overhead.
    pub fn calculate_pivot_from_range(
        &self,
        sheet: &str,
        range: Range,
        cfg: &PivotConfig,
    ) -> Result<PivotResult, PivotError> {
        let sheet_id = self
            .workbook
            .sheet_id(sheet)
            .ok_or_else(|| PivotError::SheetNotFound(sheet.to_string()))?;

        let source = materialize_range_as_pivot_values(self, sheet_id, range);
        let cache = PivotCache::from_range(&source)?;
        PivotEngine::calculate(&cache, cfg)
    }
}

fn materialize_range_as_pivot_values(
    engine: &Engine,
    sheet_id: SheetId,
    range: Range,
) -> Vec<Vec<PivotValue>> {
    let width = range.width() as usize;
    let height = range.height() as usize;

    let mut out = Vec::with_capacity(height);
    for row in range.start.row..=range.end.row {
        let mut row_out = Vec::with_capacity(width);
        for col in range.start.col..=range.end.col {
            let addr = CellAddr { row, col };
            let value = get_cell_value_at(engine, sheet_id, addr);
            let pivot_value = engine_value_to_pivot_value(&value);
            let number_format = number_format_at(engine, sheet_id, addr);
            row_out.push(coerce_pivot_value_with_number_format(
                pivot_value,
                number_format,
                engine.date_system(),
            ));
        }
        out.push(row_out);
    }
    out
}

fn number_format_at<'a>(engine: &'a Engine, sheet_id: SheetId, addr: CellAddr) -> Option<&'a str> {
    // Pivot source typing should follow the same layered number-format resolution semantics as
    // "precision as displayed" rounding: explicit per-cell overrides and spill-origin semantics,
    // then sheet < col < row < cell style layers (per-property).
    engine.number_format_pattern_for_rounding(CellKey {
        sheet: sheet_id,
        addr,
    })
}

fn get_cell_value_at(engine: &Engine, sheet_id: SheetId, addr: CellAddr) -> Value {
    // This is equivalent to `Engine::get_cell_value`, but works with already-parsed coordinates so
    // callers can iterate large source ranges without per-cell string allocations/parsing.
    if let Some(sheet_state) = engine.workbook.sheets.get(sheet_id) {
        if addr.row >= sheet_state.row_count || addr.col >= sheet_state.col_count {
            return Value::Error(ErrorKind::Ref);
        }
    }

    let key = CellKey {
        sheet: sheet_id,
        addr,
    };
    if let Some(v) = engine.spilled_cell_value(key) {
        return v;
    }
    if let Some(cell) = engine.workbook.get_cell(key) {
        // Match `Engine::get_cell_value`: allow provider-backed values to flow through
        // style-only blank cell records.
        if cell.formula.is_some() || cell.value != Value::Blank {
            return cell.value.clone();
        }
    }

    if let Some(provider) = &engine.external_value_provider {
        // Use the workbook's canonical display name to keep provider lookups stable even when
        // callers pass a different sheet-name casing.
        if let Some(sheet_name) = engine.workbook.sheet_name(sheet_id) {
            if let Some(v) = provider.get(sheet_name, addr) {
                return v;
            }
        }
    }

    Value::Blank
}

fn engine_value_to_pivot_value(value: &Value) -> PivotValue {
    match value {
        Value::Blank => PivotValue::Blank,
        Value::Number(n) => {
            // Pivot values are serialized across WASM/JSON boundaries (schema/field items). JSON
            // cannot represent non-finite numbers, and Excel surfaces these as `#NUM!`.
            if n.is_finite() {
                PivotValue::Number(*n)
            } else {
                PivotValue::Text(ErrorKind::Num.as_code().to_string())
            }
        }
        Value::Bool(b) => PivotValue::Bool(*b),
        Value::Text(s) => match parse_text_as_date(s) {
            Some(d) => PivotValue::Date(d),
            None => PivotValue::Text(s.clone()),
        },
        Value::Error(e) => PivotValue::Text(e.as_code().to_string()),
        Value::Entity(e) => PivotValue::Text(e.display.clone()),
        Value::Record(r) => {
            // Records may specify a `display_field` (Excel rich value displayField semantics).
            // Mirror grid behavior by degrading to the display-field value when present.
            if let Some(display_field) = r.display_field.as_deref() {
                if let Some(value) = r.get_field_case_insensitive(display_field) {
                    let text = value
                        .coerce_to_string()
                        .unwrap_or_else(|e| e.as_code().to_string());
                    return PivotValue::Text(text);
                }
            }
            PivotValue::Text(r.display.clone())
        }
        Value::Array(a) => {
            // If a dynamic array somehow lands in a cell value, match Excel's visible behavior
            // (top-left value shown in the origin cell).
            let top_left = a.top_left();
            engine_value_to_pivot_value(&top_left)
        }
        Value::Reference(_) | Value::ReferenceUnion(_) | Value::Lambda(_) | Value::Spill { .. } => {
            PivotValue::Blank
        }
    }
}

fn parse_text_as_date(s: &str) -> Option<NaiveDate> {
    let s = s.trim();
    if s.is_empty() {
        return None;
    }

    // Best-effort: support a handful of common date encodings seen in spreadsheet data sources.
    //
    // This is intentionally conservative: pivot typing should prefer "treat as text" over
    // accidentally converting arbitrary strings into dates.
    const FORMATS: &[&str] = &["%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m/%d/%y"];
    for fmt in FORMATS {
        if let Ok(date) = NaiveDate::parse_from_str(s, fmt) {
            return Some(date);
        }
    }
    None
}
