use crate::eval::parse_a1;
use crate::{Engine, EngineError, RecalcMode, Value};
use std::borrow::Cow;

use super::{CellRef, CellValue, WhatIfModel};

/// Adapter that exposes [`crate::Engine`] through the [`WhatIfModel`] trait.
///
/// What-If analysis typically operates on the "active sheet". This adapter
/// provides a configurable `default_sheet` used when cell references are given
/// without an explicit `Sheet!A1` prefix.
pub struct EngineWhatIfModel<'a> {
    engine: &'a mut Engine,
    default_sheet: String,
    recalc_mode: RecalcMode,
}

impl<'a> EngineWhatIfModel<'a> {
    pub fn new(engine: &'a mut Engine, default_sheet: impl Into<String>) -> Self {
        Self {
            engine,
            default_sheet: default_sheet.into(),
            // For iterative algorithms (goal seek / Monte Carlo) we default to
            // single-threaded recalculation to avoid per-iteration threadpool
            // coordination overhead.
            recalc_mode: RecalcMode::SingleThreaded,
        }
    }

    pub fn with_recalc_mode(mut self, mode: RecalcMode) -> Self {
        self.recalc_mode = mode;
        self
    }

    fn split_cell_ref<'sheet, 'cell>(
        default_sheet: &'sheet str,
        cell: &'cell CellRef,
    ) -> Result<(SheetRef<'sheet, 'cell>, &'cell str), EngineError> {
        let raw = cell.as_str().trim();

        if let Some((sheet_raw, addr_raw)) = raw.split_once('!') {
            let addr = addr_raw.trim();
            // Validate upfront so errors surface as `EngineError::Address`
            // rather than being converted into a non-numeric `#REF!` value.
            let _ = parse_a1(addr)?;

            let sheet_raw = sheet_raw.trim();
            let sheet = if let Some(inner) =
                formula_model::unquote_excel_single_quoted_identifier_lenient(sheet_raw)
            {
                if inner.is_empty() {
                    SheetRef::Default(default_sheet)
                } else {
                    match inner {
                        Cow::Borrowed(s) => SheetRef::Explicit(s),
                        Cow::Owned(s) => SheetRef::Owned(s),
                    }
                }
            } else if sheet_raw.is_empty() {
                SheetRef::Default(default_sheet)
            } else {
                SheetRef::Explicit(sheet_raw)
            };

            Ok((sheet, addr))
        } else {
            let _ = parse_a1(raw)?;
            Ok((SheetRef::Default(default_sheet), raw))
        }
    }

    fn map_to_engine_value(value: CellValue) -> Value {
        match value {
            CellValue::Number(v) => Value::Number(v),
            CellValue::Text(v) => Value::Text(v),
            CellValue::Bool(v) => Value::Bool(v),
            CellValue::Blank => Value::Blank,
        }
    }

    fn map_from_engine_value(value: Value) -> CellValue {
        match value {
            Value::Number(v) => CellValue::Number(v),
            Value::Text(v) => CellValue::Text(v),
            Value::Bool(v) => CellValue::Bool(v),
            Value::Entity(entity) => CellValue::Text(entity.display),
            Value::Record(record) => CellValue::Text(record.display),
            Value::Blank => CellValue::Blank,
            Value::Error(e) => CellValue::Text(e.as_code().to_string()),
            Value::Array(arr) => Self::map_from_engine_value(arr.top_left()),
            // The What-If API intentionally only supports scalar values, so anything richer
            // (references, lambdas, spill markers, entities/records, etc.) is degraded to the
            // engine's display string.
            other => CellValue::Text(other.to_string()),
        }
    }
}

enum SheetRef<'a, 'b> {
    Default(&'a str),
    Explicit(&'b str),
    Owned(String),
}

impl SheetRef<'_, '_> {
    fn as_str(&self) -> &str {
        match self {
            SheetRef::Default(v) | SheetRef::Explicit(v) => v,
            SheetRef::Owned(v) => v.as_str(),
        }
    }
}

impl WhatIfModel for EngineWhatIfModel<'_> {
    type Error = EngineError;

    fn get_cell_value(&self, cell: &CellRef) -> Result<CellValue, Self::Error> {
        let default_sheet = self.default_sheet.as_str();
        let (sheet, addr) = Self::split_cell_ref(default_sheet, cell)?;
        Ok(Self::map_from_engine_value(
            self.engine.get_cell_value(sheet.as_str(), addr),
        ))
    }

    fn set_cell_value(&mut self, cell: &CellRef, value: CellValue) -> Result<(), Self::Error> {
        let default_sheet = self.default_sheet.as_str();
        let (sheet, addr) = Self::split_cell_ref(default_sheet, cell)?;
        self.engine
            .set_cell_value(sheet.as_str(), addr, Self::map_to_engine_value(value))
    }

    fn recalculate(&mut self) -> Result<(), Self::Error> {
        match self.recalc_mode {
            RecalcMode::SingleThreaded => self.engine.recalculate_single_threaded(),
            RecalcMode::MultiThreaded => self.engine.recalculate_multi_threaded(),
        }
        Ok(())
    }
}
