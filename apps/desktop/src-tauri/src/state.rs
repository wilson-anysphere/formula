use crate::file_io::Workbook;
use crate::macros::{
    execute_invocation, MacroExecutionOptions, MacroExecutionOutcome, MacroHost, MacroHostError,
    MacroInfo, MacroInvocation, MacroRuntimeContext,
};
use crate::persistence::{
    autosave_db_path_for_workbook, open_memory_manager, open_storage, workbook_from_model,
    workbook_to_model, PersistentWorkbookState, WorkbookPersistenceLocation,
};
use chrono::Datelike;
use formula_columnar::{ColumnType as ColumnarType, ColumnarTable, Value as ColumnarValue};
use formula_engine::eval::{parse_a1, CellAddr};
use formula_engine::metadata::FormatRun;
use formula_engine::pivot::{PivotConfig, PivotEngine, PivotTable as EnginePivotTable, PivotValue};
use formula_engine::style_bridge::ui_style_to_model_style;
use formula_engine::what_if::goal_seek::{GoalSeek, GoalSeekParams, GoalSeekResult};
use formula_engine::what_if::monte_carlo::{MonteCarloEngine, SimulationConfig, SimulationResult};
use formula_engine::what_if::scenario_manager::{
    Scenario, ScenarioId, ScenarioManager, SummaryReport,
};
use formula_engine::what_if::{
    CellRef as WhatIfCellRef, CellValue as WhatIfCellValue, EngineWhatIfModel, WhatIfModel,
};
use formula_engine::{
    Engine as FormulaEngine, ErrorKind, ExternalValueProvider, NameDefinition, NameScope,
    PrecedentNode, RecalcMode, Value as EngineValue,
};
use formula_format::{format_value, FormatOptions, Value as FormatValue};
use formula_storage::{
    AutoSaveConfig, AutoSaveManager, CellChange, CellData as StorageCellData,
    CellRange as StorageCellRange, ImportModelWorkbookOptions,
};
use formula_model::{CellRef as ModelCellRef, Range as ModelRange, SheetVisibility, TabColor};
use formula_xlsx::print::{
    CellRange as PrintCellRange, ManualPageBreaks, PageSetup, SheetPrintSettings,
};
use crate::resource_limits::{MAX_ORIGIN_XLSX_BYTES, MAX_RANGE_CELLS_PER_CALL, MAX_RANGE_DIM};
use crate::sheet_name::sheet_name_eq_case_insensitive;
use serde_json::Value as JsonValue;
use std::collections::{HashMap, HashSet};
use std::io::Cursor;
use std::path::Path;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Arc;
use uuid::Uuid;

#[derive(Debug, thiserror::Error)]
pub enum AppStateError {
    #[error("no workbook loaded")]
    NoWorkbookLoaded,
    #[error("no undo history")]
    NoUndoHistory,
    #[error("no redo history")]
    NoRedoHistory,
    #[error("cannot delete the last remaining sheet")]
    CannotDeleteLastSheet,
    #[error("unknown sheet id: {0}")]
    UnknownSheet(String),
    #[error("invalid sheet index: {to_index} (sheet count {sheet_count})")]
    InvalidSheetIndex {
        to_index: usize,
        sheet_count: usize,
    },
    #[error("invalid range: start ({start_row},{start_col}) end ({end_row},{end_col})")]
    InvalidRange {
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    },
    #[error("range too large: {rows}x{cols} cells exceeds limit {limit}")]
    RangeTooLarge {
        rows: usize,
        cols: usize,
        limit: usize,
    },
    #[error("range too large: {rows}x{cols} exceeds max dimension {limit}")]
    RangeDimensionTooLarge {
        rows: usize,
        cols: usize,
        limit: usize,
    },
    #[error("unknown pivot id: {0}")]
    UnknownPivot(String),
    #[error("pivot table failed: {0}")]
    Pivot(String),
    #[error("what-if analysis failed: {0}")]
    WhatIf(String),
    #[error("persistence failed: {0}")]
    Persistence(String),
    #[error("formula engine error: {0}")]
    Engine(String),
    #[error("too many auditing results for {kind} (limit {limit})")]
    AuditingTooLarge { kind: String, limit: usize },
}

#[derive(Clone, Debug, PartialEq)]
pub enum CellScalar {
    Empty,
    Number(f64),
    Text(String),
    Bool(bool),
    Error(String),
}

impl CellScalar {
    pub fn display(&self) -> String {
        match self {
            CellScalar::Empty => String::new(),
            CellScalar::Number(_) => format_scalar_for_display_with_date_system(
                self,
                None,
                formula_format::DateSystem::Excel1900,
            ),
            CellScalar::Text(s) => s.clone(),
            CellScalar::Bool(true) => "TRUE".to_string(),
            CellScalar::Bool(false) => "FALSE".to_string(),
            CellScalar::Error(e) => e.clone(),
        }
    }

    pub fn as_json(&self) -> Option<JsonValue> {
        match self {
            CellScalar::Empty => None,
            CellScalar::Number(n) => Some(JsonValue::from(*n)),
            CellScalar::Text(s) => Some(JsonValue::from(s.clone())),
            CellScalar::Bool(b) => Some(JsonValue::from(*b)),
            CellScalar::Error(e) => Some(JsonValue::from(e.clone())),
        }
    }

    pub fn from_json(value: &JsonValue) -> Self {
        match value {
            JsonValue::Null => CellScalar::Empty,
            JsonValue::Bool(b) => CellScalar::Bool(*b),
            JsonValue::Number(n) => n
                .as_f64()
                .map(CellScalar::Number)
                .unwrap_or_else(|| CellScalar::Error("#NUM!".to_string())),
            JsonValue::String(s) => CellScalar::Text(s.clone()),
            other => CellScalar::Text(other.to_string()),
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct Cell {
    pub(crate) input_value: Option<CellScalar>,
    pub(crate) formula: Option<String>,
    pub(crate) computed_value: CellScalar,
    pub(crate) number_format: Option<String>,
}

impl Cell {
    pub(crate) fn empty() -> Self {
        Self {
            input_value: None,
            formula: None,
            computed_value: CellScalar::Empty,
            number_format: None,
        }
    }

    pub(crate) fn from_literal(value: Option<CellScalar>) -> Self {
        let computed_value = value.clone().unwrap_or(CellScalar::Empty);
        Self {
            input_value: value,
            formula: None,
            computed_value,
            number_format: None,
        }
    }

    pub(crate) fn from_formula(formula: String) -> Self {
        Self {
            input_value: None,
            formula: Some(formula),
            computed_value: CellScalar::Empty,
            number_format: None,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct CellData {
    pub value: CellScalar,
    pub formula: Option<String>,
    pub display_value: String,
}

#[derive(Clone, Debug, PartialEq)]
pub struct CellUpdateData {
    pub sheet_id: String,
    pub row: usize,
    pub col: usize,
    pub value: CellScalar,
    pub formula: Option<String>,
    pub display_value: String,
}

#[derive(Clone, Debug, PartialEq)]
pub struct SheetInfoData {
    pub id: String,
    pub name: String,
    pub visibility: SheetVisibility,
    pub tab_color: Option<TabColor>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct WorkbookInfoData {
    pub path: Option<String>,
    pub origin_path: Option<String>,
    pub sheets: Vec<SheetInfoData>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellRect {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

impl CellRect {
    fn contains(&self, row: usize, col: usize) -> bool {
        row >= self.start_row && row <= self.end_row && col >= self.start_col && col <= self.end_col
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PivotDestination {
    pub sheet_id: String,
    pub row: usize,
    pub col: usize,
}

#[derive(Clone, Debug)]
struct PivotRegistration {
    id: String,
    name: String,
    source_sheet_id: String,
    source_range: CellRect,
    destination: PivotDestination,
    config: PivotConfig,
    last_output_range: Option<CellRect>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PivotTableSummary {
    pub id: String,
    pub name: String,
    pub source_sheet_id: String,
    pub source_range: CellRect,
    pub destination: PivotDestination,
}

#[derive(Default)]
struct PivotManager {
    pivots: Vec<PivotRegistration>,
}

impl PivotManager {
    fn next_id() -> String {
        static NEXT_ID: AtomicU64 = AtomicU64::new(1);
        format!("pivot-{}", NEXT_ID.fetch_add(1, Ordering::Relaxed))
    }
}

#[derive(Clone, Debug, PartialEq)]
struct CellInputSnapshot {
    sheet_id: String,
    row: usize,
    col: usize,
    value: Option<CellScalar>,
    formula: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
struct UndoEntry {
    before: Vec<CellInputSnapshot>,
    after: Vec<CellInputSnapshot>,
}

pub type SharedAppState = std::sync::Arc<std::sync::Mutex<AppState>>;

const AUDITING_RESULT_LIMIT: usize = 2_000;
const FORMULA_UI_FORMATTING_METADATA_KEY: &str = "formula_ui_formatting";

pub struct AppState {
    workbook: Option<Workbook>,
    persistent: Option<PersistentWorkbookState>,
    engine: FormulaEngine,
    dirty: bool,
    undo_stack: Vec<UndoEntry>,
    redo_stack: Vec<UndoEntry>,
    scenario_manager: ScenarioManager,
    macro_host: MacroHost,
    pivots: PivotManager,
}

impl Default for AppState {
    fn default() -> Self {
        Self::new()
    }
}

impl AppState {
    pub fn new() -> Self {
        Self {
            workbook: None,
            persistent: None,
            engine: FormulaEngine::new(),
            dirty: false,
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
            scenario_manager: ScenarioManager::new(),
            macro_host: MacroHost::default(),
            pivots: PivotManager::default(),
        }
    }

    pub fn has_unsaved_changes(&self) -> bool {
        self.dirty
    }

    pub fn mark_dirty(&mut self) {
        self.dirty = true;
        self.redo_stack.clear();
    }

    pub fn workbook_info(&self) -> Result<WorkbookInfoData, AppStateError> {
        let workbook = self
            .workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        Ok(WorkbookInfoData {
            path: workbook.path.clone(),
            origin_path: workbook.origin_path.clone(),
            sheets: workbook
                .sheets
                .iter()
                .map(|sheet| SheetInfoData {
                    id: sheet.id.clone(),
                    name: sheet.name.clone(),
                    visibility: sheet.visibility,
                    tab_color: sheet.tab_color.clone(),
                })
                .collect(),
        })
    }

    pub fn create_sheet(&mut self, name: String) -> Result<String, AppStateError> {
        let (candidate_id, sheet_name, position) = {
            let workbook = self.get_workbook()?;
            let trimmed = name.trim();
            formula_model::validate_sheet_name(trimmed)
                .map_err(|e| AppStateError::WhatIf(e.to_string()))?;
            if workbook
                .sheets
                .iter()
                .any(|sheet| sheet_name_eq_case_insensitive(&sheet.name, trimmed))
            {
                return Err(AppStateError::WhatIf(
                    formula_model::SheetNameError::DuplicateName.to_string(),
                ));
            }

            let base_id = trimmed.to_string();
            let mut candidate = base_id.clone();
            let mut counter = 1usize;
            while workbook.sheets.iter().any(|s| s.id.eq_ignore_ascii_case(&candidate)) {
                counter += 1;
                candidate = format!("{base_id}-{counter}");
            }

            // Append by default.
            let position = workbook.sheets.len() as i64;

            (candidate, trimmed.to_string(), position)
        };

        if let Some(persistent) = self.persistent.as_mut() {
            let sheet = persistent
                .storage
                .create_sheet(persistent.workbook_id, &sheet_name, position, None)
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;
            persistent.sheet_map.insert(candidate_id.clone(), sheet.id);
        }

        {
            let workbook = self.get_workbook_mut()?;
            workbook.sheets.push(crate::file_io::Sheet::new(
                candidate_id.clone(),
                sheet_name.clone(),
            ));

            // Adding sheets is a structural XLSX edit. The patch-based save path (which relies on
            // `origin_xlsx_bytes`) can only patch existing worksheet parts; it cannot add new sheets.
            // Drop the origin bytes so the next save/export regenerates from storage.
            workbook.origin_xlsx_bytes = None;
            workbook.origin_xlsb_path = None;
        }

        self.engine.ensure_sheet(&sheet_name);

        self.dirty = true;
        self.redo_stack.clear();
        Ok(candidate_id)
    }

    pub fn add_sheet(
        &mut self,
        name: String,
        sheet_id: Option<String>,
        after_sheet_id: Option<String>,
        index: Option<usize>,
    ) -> Result<SheetInfoData, AppStateError> {
        let ctx = self.macro_host.runtime_context();
        let (candidate_id, candidate_name, insert_index, sheet_count_before) = {
            let workbook = self.get_workbook()?;
            let sheet_count_before = workbook.sheets.len();
            let base = name.trim();
            formula_model::validate_sheet_name(base)
                .map_err(|e| AppStateError::WhatIf(e.to_string()))?;
            let insert_index = after_sheet_id
                .as_deref()
                .map(str::trim)
                .filter(|id| !id.is_empty())
                .and_then(|after_id| {
                    workbook
                        .sheets
                        .iter()
                        .position(|sheet| sheet.id.eq_ignore_ascii_case(after_id))
                        .map(|idx| idx.saturating_add(1))
                })
                .unwrap_or_else(|| index.unwrap_or(sheet_count_before))
                .min(sheet_count_before);

            if let Some(sheet_id) = sheet_id {
                let trimmed_id = sheet_id.trim();
                if trimmed_id.is_empty() {
                    return Err(AppStateError::WhatIf(
                        "sheet id must not be empty".to_string(),
                    ));
                }
                if workbook
                    .sheets
                    .iter()
                    .any(|sheet| sheet.id.eq_ignore_ascii_case(trimmed_id))
                {
                    return Err(AppStateError::WhatIf(format!(
                        "duplicate sheet id: {trimmed_id}"
                    )));
                }
                if workbook
                    .sheets
                    .iter()
                    .any(|sheet| sheet_name_eq_case_insensitive(&sheet.name, base))
                {
                    return Err(AppStateError::WhatIf(
                        formula_model::SheetNameError::DuplicateName.to_string(),
                    ));
                }

                (
                    trimmed_id.to_string(),
                    base.to_string(),
                    insert_index,
                    sheet_count_before,
                )
            } else {
                let mut candidate_name = base.to_string();
                let mut counter = 1usize;
                while workbook
                    .sheets
                    .iter()
                    .any(|sheet| sheet_name_eq_case_insensitive(&sheet.name, &candidate_name))
                {
                    counter += 1;
                    let suffix = format!(" {counter}");
                    let suffix_len = suffix.encode_utf16().count();
                    let max_base_len =
                        formula_model::EXCEL_MAX_SHEET_NAME_LEN.saturating_sub(suffix_len);
                    let mut used_len = 0usize;
                    let mut truncated = String::new();
                    for ch in base.chars() {
                        let ch_len = ch.len_utf16();
                        if used_len + ch_len > max_base_len {
                            break;
                        }
                        used_len += ch_len;
                        truncated.push(ch);
                    }
                    candidate_name = format!("{truncated}{suffix}");
                }
                formula_model::validate_sheet_name(&candidate_name)
                    .map_err(|e| AppStateError::WhatIf(e.to_string()))?;

                let base_id = candidate_name.clone();
                let mut candidate_id = base_id.clone();
                let mut id_counter = 1usize;
                while workbook
                    .sheets
                    .iter()
                    .any(|sheet| sheet.id.eq_ignore_ascii_case(&candidate_id))
                {
                    id_counter += 1;
                    candidate_id = format!("{base_id}-{id_counter}");
                }

                (candidate_id, candidate_name, insert_index, sheet_count_before)
            }
        };

        if let Some(persistent) = self.persistent.as_mut() {
            let sheet = persistent
                .storage
                .create_sheet(
                    persistent.workbook_id,
                    &candidate_name,
                    insert_index as i64,
                    None,
                )
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;
            persistent.sheet_map.insert(candidate_id.clone(), sheet.id);
        }

        {
            let workbook = self.get_workbook_mut()?;
            workbook.sheets.insert(
                insert_index,
                crate::file_io::Sheet::new(candidate_id.clone(), candidate_name.clone()),
            );

            // The preserved pivot cache parts store the original workbook sheet list (including the
            // sheet indices used to rewrite `worksheetSource sheet="..."` references on apply).
            // Keep those indices aligned with structural edits so later renames still map correctly.
            if let Some(preserved) = workbook.preserved_pivot_parts.as_mut() {
                for sheet in &mut preserved.workbook_sheets {
                    if sheet.index >= insert_index {
                        sheet.index = sheet.index.saturating_add(1);
                    }
                }
            }

            // Adding sheets is a structural XLSX edit. The patch-based save path (which relies on
            // `origin_xlsx_bytes`) can only patch existing worksheet parts; it cannot add new sheets.
            // Drop the origin bytes so the next save/export regenerates from storage.
            workbook.origin_xlsx_bytes = None;
            workbook.origin_xlsb_path = None;
        }

        // The macro runtime context stores sheet indices; keep them stable so it continues to point
        // at the same sheet(s) after inserting a new sheet before them.
        if insert_index <= ctx.active_sheet || ctx.selection.is_some() {
            let mut next_ctx = ctx;
            if insert_index <= ctx.active_sheet {
                next_ctx.active_sheet = ctx.active_sheet.saturating_add(1);
            }
            if let Some(sel) = ctx.selection {
                let mut next_sel = sel;
                if sel.sheet >= insert_index {
                    next_sel.sheet = sel.sheet.saturating_add(1);
                }
                next_ctx.selection = Some(next_sel);
            }
            if next_ctx != ctx {
                // Avoid `self.get_workbook()` here: it borrows the entire `AppState` immutably for
                // the lifetime of the returned reference, which prevents us from mutably borrowing
                // the disjoint `macro_host` field below.
                let workbook = self
                    .workbook
                    .as_ref()
                    .ok_or(AppStateError::NoWorkbookLoaded)?;
                let macro_host = &mut self.macro_host;
                macro_host.sync_with_workbook(workbook);
                macro_host.set_runtime_context(next_ctx);
            }
        }

        // The formula engine indexes sheets by insertion order. When we insert a new sheet into
        // the middle of the workbook, we need to rebuild the engine so:
        // - sheet indices match the workbook order (important for 3D references)
        // - dependency graphs are updated to include the newly inserted sheet
        if insert_index < sheet_count_before {
            self.rebuild_engine_from_workbook()?;
            let recalc_changes = self.engine.recalculate_with_value_changes_multi_threaded();
            let _ = self.refresh_computed_values_from_recalc_changes(&recalc_changes)?;
        } else {
            self.engine.ensure_sheet(&candidate_name);
        }

        self.dirty = true;
        self.redo_stack.clear();
        Ok(SheetInfoData {
            id: candidate_id,
            name: candidate_name,
            visibility: SheetVisibility::Visible,
            tab_color: None,
        })
    }

    pub fn add_sheet_with_id(
        &mut self,
        sheet_id: String,
        name: String,
        after_sheet_id: Option<String>,
        index: Option<usize>,
    ) -> Result<(), AppStateError> {
        let ctx = self.macro_host.runtime_context();
        let (sheet_id, name, insert_index, sheet_count_before) = {
            let workbook = self.get_workbook()?;
            let sheet_count_before = workbook.sheets.len();

            let sheet_id = sheet_id.trim();
            if sheet_id.is_empty() {
                return Err(AppStateError::WhatIf(
                    "sheet id must be non-empty".to_string(),
                ));
            }
            if workbook
                .sheets
                .iter()
                .any(|sheet| sheet.id.eq_ignore_ascii_case(sheet_id))
            {
                return Err(AppStateError::WhatIf("sheet id must be unique".to_string()));
            }

            let trimmed = name.trim();
            formula_model::validate_sheet_name(trimmed)
                .map_err(|e| AppStateError::WhatIf(e.to_string()))?;
            if workbook
                .sheets
                .iter()
                .any(|sheet| sheet_name_eq_case_insensitive(&sheet.name, trimmed))
            {
                return Err(AppStateError::WhatIf(
                    formula_model::SheetNameError::DuplicateName.to_string(),
                ));
            }

            let insert_index = after_sheet_id
                .as_deref()
                .map(str::trim)
                .filter(|id| !id.is_empty())
                .and_then(|after_id| {
                    workbook
                        .sheets
                        .iter()
                        .position(|sheet| sheet.id.eq_ignore_ascii_case(after_id))
                        .map(|idx| idx.saturating_add(1))
                })
                .unwrap_or_else(|| index.unwrap_or(sheet_count_before))
                .min(sheet_count_before);

            (sheet_id.to_string(), trimmed.to_string(), insert_index, sheet_count_before)
        };

        if let Some(persistent) = self.persistent.as_mut() {
            let sheet = persistent
                .storage
                .create_sheet(
                    persistent.workbook_id,
                    &name,
                    insert_index as i64,
                    None,
                )
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;
            persistent.sheet_map.insert(sheet_id.clone(), sheet.id);
        }

        {
            let workbook = self.get_workbook_mut()?;
            workbook.sheets.insert(
                insert_index,
                crate::file_io::Sheet::new(sheet_id.clone(), name.clone()),
            );

            // Keep preserved pivot cache sheet indices aligned with the inserted sheet so
            // index-based worksheetSource rewrites remain correct.
            if let Some(preserved) = workbook.preserved_pivot_parts.as_mut() {
                for sheet in &mut preserved.workbook_sheets {
                    if sheet.index >= insert_index {
                        sheet.index = sheet.index.saturating_add(1);
                    }
                }
            }

            // Adding sheets is a structural XLSX edit. The patch-based save path (which relies on
            // `origin_xlsx_bytes`) can only patch existing worksheet parts; it cannot add new sheets.
            // Drop the origin bytes so the next save/export regenerates from storage.
            workbook.origin_xlsx_bytes = None;
            workbook.origin_xlsb_path = None;
        }

        // The macro runtime context stores sheet indices; keep them stable so it continues to point
        // at the same sheet(s) after inserting a new sheet before them.
        if insert_index <= ctx.active_sheet || ctx.selection.is_some() {
            let mut next_ctx = ctx;
            if insert_index <= ctx.active_sheet {
                next_ctx.active_sheet = ctx.active_sheet.saturating_add(1);
            }
            if let Some(sel) = ctx.selection {
                let mut next_sel = sel;
                if sel.sheet >= insert_index {
                    next_sel.sheet = sel.sheet.saturating_add(1);
                }
                next_ctx.selection = Some(next_sel);
            }
            if next_ctx != ctx {
                // Avoid `self.get_workbook()` here: it borrows the entire `AppState` immutably for
                // the lifetime of the returned reference, which prevents us from mutably borrowing
                // the disjoint `macro_host` field below.
                let workbook = self
                    .workbook
                    .as_ref()
                    .ok_or(AppStateError::NoWorkbookLoaded)?;
                let macro_host = &mut self.macro_host;
                macro_host.sync_with_workbook(workbook);
                macro_host.set_runtime_context(next_ctx);
            }
        }

        // The formula engine indexes sheets by insertion order. When we insert a new sheet into
        // the middle of the workbook, we need to rebuild the engine so:
        // - sheet indices match the workbook order (important for 3D references)
        // - dependency graphs are updated to include the newly inserted sheet
        if insert_index < sheet_count_before {
            self.rebuild_engine_from_workbook()?;
            let recalc_changes = self.engine.recalculate_with_value_changes_multi_threaded();
            let _ = self.refresh_computed_values_from_recalc_changes(&recalc_changes)?;
        } else {
            self.engine.ensure_sheet(&name);
        }

        self.dirty = true;
        self.redo_stack.clear();
        Ok(())
    }

    pub fn reorder_sheets(&mut self, sheet_ids: Vec<String>) -> Result<(), AppStateError> {
        let desired: Vec<String> = sheet_ids
            .into_iter()
            .map(|id| id.trim().to_string())
            .filter(|id| !id.is_empty())
            .collect();
        if desired.is_empty() {
            return Ok(());
        }

        let (ctx, active_sheet_id, selection_sheet_id, before_order, ordered_ids) = {
            let workbook = self.get_workbook()?;
            if workbook.sheets.len() <= 1 {
                return Ok(());
            }

            let before_order: Vec<String> = workbook.sheets.iter().map(|s| s.id.clone()).collect();

            let mut seen = HashSet::new();
            let mut ordered_ids: Vec<String> = Vec::with_capacity(before_order.len());

            for raw in &desired {
                if raw.is_empty() {
                    continue;
                }
                if seen.contains(raw) {
                    continue;
                }
                if !workbook.sheets.iter().any(|sheet| sheet.id == *raw) {
                    continue;
                }
                seen.insert(raw.clone());
                ordered_ids.push(raw.clone());
            }

            for id in &before_order {
                if seen.contains(id) {
                    continue;
                }
                seen.insert(id.clone());
                ordered_ids.push(id.clone());
            }

            if ordered_ids.len() <= 1
                || (ordered_ids.len() == before_order.len()
                    && ordered_ids
                        .iter()
                        .zip(before_order.iter())
                        .all(|(a, b)| a == b))
            {
                return Ok(());
            }

            let ctx = self.macro_host.runtime_context();
            let active_sheet_id = workbook.sheets.get(ctx.active_sheet).map(|s| s.id.clone());
            let selection_sheet_id = ctx
                .selection
                .as_ref()
                .and_then(|sel| workbook.sheets.get(sel.sheet).map(|s| s.id.clone()));

            (
                ctx,
                active_sheet_id,
                selection_sheet_id,
                before_order,
                ordered_ids,
            )
        };

        // Update persistent storage first so we can fail fast without mutating the in-memory workbook.
        if let Some(persistent) = self.persistent.as_ref() {
            let mut sheet_uuids = Vec::with_capacity(ordered_ids.len());
            for sheet_id in &ordered_ids {
                let sheet_uuid = persistent.sheet_uuid(sheet_id).ok_or_else(|| {
                    AppStateError::Persistence(format!(
                        "missing persistence mapping for sheet id {sheet_id}"
                    ))
                })?;
                sheet_uuids.push(sheet_uuid);
            }

            persistent
                .storage
                .reorder_sheets(persistent.workbook_id, &sheet_uuids)
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;
        }

        {
            let workbook = self.get_workbook_mut()?;

            let mut remaining = std::mem::take(&mut workbook.sheets);
            let mut next = Vec::with_capacity(remaining.len());
            for id in &ordered_ids {
                if let Some(idx) = remaining.iter().position(|s| s.id == *id) {
                    next.push(remaining.remove(idx));
                }
            }
            next.extend(remaining);
            workbook.sheets = next;

            // Keep preserved pivot cache sheet indices aligned with the reordered sheet list so
            // index-based worksheetSource rewrites remain correct after a reorder+rename sequence.
            if let Some(preserved) = workbook.preserved_pivot_parts.as_mut() {
                let mut old_idx_by_sheet_id: HashMap<&str, usize> =
                    HashMap::with_capacity(before_order.len());
                for (idx, sheet_id) in before_order.iter().enumerate() {
                    old_idx_by_sheet_id.insert(sheet_id.as_str(), idx);
                }
                let mut new_idx_by_old: HashMap<usize, usize> =
                    HashMap::with_capacity(before_order.len());
                for (new_idx, sheet_id) in ordered_ids.iter().enumerate() {
                    if let Some(old_idx) = old_idx_by_sheet_id.get(sheet_id.as_str()) {
                        new_idx_by_old.insert(*old_idx, new_idx);
                    }
                }

                for sheet in &mut preserved.workbook_sheets {
                    if let Some(new_idx) = new_idx_by_old.get(&sheet.index) {
                        sheet.index = *new_idx;
                    }
                }
            }

            // NOTE: Reordering sheets is a structural edit for XLSB inputs. The `.xlsb` writer
            // cannot currently reorder workbook metadata, so treat this as a conversion to XLSX
            // on the next save (consistent with other structural edits like add/delete/rename).
            if workbook.origin_xlsb_path.is_some() {
                workbook.origin_xlsb_path = None;
            }
        }

        // The macro runtime context stores sheet indices; keep them stable so it continues to
        // point at the same sheet(s) after reordering.
        {
            let workbook = self
                .workbook
                .as_ref()
                .ok_or(AppStateError::NoWorkbookLoaded)?;
            let mut new_ctx = ctx;

            if let Some(active_sheet_id) = active_sheet_id {
                if let Some(idx) = workbook
                    .sheets
                    .iter()
                    .position(|s| s.id == active_sheet_id)
                {
                    new_ctx.active_sheet = idx;
                }
            }

            if let (Some(sel), Some(selection_sheet_id)) =
                (new_ctx.selection.as_mut(), selection_sheet_id)
            {
                if let Some(idx) = workbook
                    .sheets
                    .iter()
                    .position(|s| s.id == selection_sheet_id)
                {
                    sel.sheet = idx;
                }
            }

            self.macro_host.sync_with_workbook(workbook);
            self.macro_host.set_runtime_context(new_ctx);
        }

        // Sheet order affects some Excel semantics (e.g. 3D references like `Sheet1:Sheet3!A1`).
        // Rebuild to ensure evaluations reflect the updated sheet ordering.
        self.rebuild_engine_from_workbook()?;
        self.engine.recalculate();
        let _ = self.refresh_computed_values()?;

        // Defensive: the reorder list we applied may have ignored some unknown ids; keep the
        // caller-visible ordering in sync with the actual workbook model.
        debug_assert_ne!(before_order, ordered_ids);

        self.dirty = true;
        self.redo_stack.clear();
        Ok(())
    }

    pub fn rename_sheet(&mut self, sheet_id: &str, name: String) -> Result<(), AppStateError> {
        let trimmed = name.trim();
        formula_model::validate_sheet_name(trimmed)
            .map_err(|e| AppStateError::WhatIf(e.to_string()))?;

        {
            let workbook = self.get_workbook()?;
            if workbook.sheets.iter().any(|sheet| {
                sheet.id != sheet_id && sheet_name_eq_case_insensitive(&sheet.name, trimmed)
            }) {
                return Err(AppStateError::WhatIf(
                    formula_model::SheetNameError::DuplicateName.to_string(),
                ));
            }
        }

        let (old_name, new_name) = {
            let workbook = self.get_workbook()?;
            let sheet = workbook
                .sheet(sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
            let old_name = sheet.name.clone();
            if old_name == trimmed {
                return Ok(());
            }

            (old_name, trimmed.to_string())
        };

        // When the workbook is backed by SQLite persistence, rename the sheet in storage first so
        // validation/rewrite failures don't leave the in-memory model in a partially updated state.
        if let Some(persistent) = self.persistent.as_ref() {
            let sheet_uuid = persistent.sheet_uuid(sheet_id).ok_or_else(|| {
                AppStateError::Persistence(format!(
                    "missing persistence mapping for sheet id {sheet_id}"
                ))
            })?;

            // Ensure all pending in-memory edits are reflected in SQLite so the rewrite pass runs
            // against the latest cell formulas.
            persistent
                .memory
                .flush_dirty_pages()
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;

            persistent
                .storage
                .rename_sheet(sheet_uuid, &new_name)
                .map_err(|e| match e {
                    formula_storage::StorageError::EmptySheetName => {
                        AppStateError::WhatIf("sheet name cannot be empty".to_string())
                    }
                    formula_storage::StorageError::InvalidSheetName(err) => {
                        AppStateError::WhatIf(err.to_string())
                    }
                    formula_storage::StorageError::DuplicateSheetName(name) => {
                        AppStateError::WhatIf(format!("sheet name already exists: {name}"))
                    }
                    formula_storage::StorageError::SheetNotFound(_) => {
                        AppStateError::UnknownSheet(sheet_id.to_string())
                    }
                    other => AppStateError::Persistence(other.to_string()),
                })?;

            // Sheet renames rewrite formulas directly in SQLite; clear the page cache so subsequent
            // viewport reads observe updated formula text.
            persistent.memory.clear_cache();
        }

        {
            let workbook = self.get_workbook_mut()?;
            let sheet = workbook
                .sheet_mut(sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
            sheet.name = new_name.clone();

            // Rewrite cross-sheet references in formulas and workbook-level metadata.
            for sheet in &mut workbook.sheets {
                for ((row, col), cell) in sheet.cells.iter_mut() {
                    let Some(formula) = cell.formula.as_mut() else {
                        continue;
                    };
                    let rewritten =
                        formula_model::rewrite_sheet_names_in_formula(formula, &old_name, &new_name);
                    if rewritten != *formula {
                        *formula = rewritten;
                        // Track as an input edit so patch-based XLSX saves include updated formulas.
                        sheet.dirty_cells.insert((*row, *col));
                    }
                }
            }

            for name in &mut workbook.defined_names {
                name.refers_to = formula_model::rewrite_sheet_names_in_formula(
                    &name.refers_to,
                    &old_name,
                    &new_name,
                );
                if let Some(sheet_key) = name.sheet_id.as_mut() {
                    if sheet_name_eq_case_insensitive(sheet_key, &old_name) {
                        *sheet_key = new_name.clone();
                    }
                }
            }

            for table in &mut workbook.tables {
                if sheet_name_eq_case_insensitive(&table.sheet_id, &old_name) {
                    table.sheet_id = new_name.clone();
                }
            }

            // Preserve print settings keyed by sheet name.
            for settings in &mut workbook.print_settings.sheets {
                if sheet_name_eq_case_insensitive(&settings.sheet_name, &old_name) {
                    settings.sheet_name = new_name.clone();
                }
            }
            for settings in &mut workbook.original_print_settings.sheets {
                if sheet_name_eq_case_insensitive(&settings.sheet_name, &old_name) {
                    settings.sheet_name = new_name.clone();
                }
            }

            // Keep preserved XLSX artifacts (pivots/drawings) keyed by sheet name aligned with any
            // in-app rename. These maps are used when regenerating an XLSX package after structural
            // edits (renames drop `origin_xlsx_bytes`). If we don't update the keys, downstream
            // preservation logic may fall back to stale sheet indices and incorrectly attach parts
            // after subsequent sheet mutations (delete/reorder).
            if let Some(preserved) = workbook.preserved_pivot_parts.as_mut() {
                let key = preserved
                    .sheet_pivot_tables
                    .keys()
                    .find(|k| sheet_name_eq_case_insensitive(k, &old_name))
                    .cloned();
                if let Some(key) = key {
                    if let Some(value) = preserved.sheet_pivot_tables.remove(&key) {
                        preserved.sheet_pivot_tables.insert(new_name.clone(), value);
                    }
                }
                // NOTE: `preserved.workbook_sheets` intentionally retains the sheet names from the
                // original workbook.xml. These names are used to rewrite pivot cache
                // `worksheetSource sheet="..."` references when a sheet is renamed in-app. The
                // index field is kept aligned with structural edits (insert/delete/reorder), so we
                // can still map old names -> current sheet names by index at save time.
            }

            if let Some(preserved) = workbook.preserved_drawing_parts.as_mut() {
                fn rename_preserved_map_key<V>(
                    map: &mut std::collections::BTreeMap<String, V>,
                    old_name: &str,
                    new_name: &str,
                ) {
                    let key = map
                        .keys()
                        .find(|k| sheet_name_eq_case_insensitive(k, old_name))
                        .cloned();
                    if let Some(key) = key {
                        if let Some(value) = map.remove(&key) {
                            map.insert(new_name.to_string(), value);
                        }
                    }
                }

                rename_preserved_map_key(&mut preserved.sheet_drawings, &old_name, &new_name);
                rename_preserved_map_key(&mut preserved.sheet_pictures, &old_name, &new_name);
                rename_preserved_map_key(&mut preserved.sheet_ole_objects, &old_name, &new_name);
                rename_preserved_map_key(&mut preserved.sheet_controls, &old_name, &new_name);
                rename_preserved_map_key(&mut preserved.sheet_drawing_hfs, &old_name, &new_name);
                rename_preserved_map_key(&mut preserved.chart_sheets, &old_name, &new_name);
            }

            // Renaming sheets is a structural XLSX edit. The patch-based save path cannot rewrite
            // workbook.xml relationships/sheet names, so drop origin bytes to force regeneration
            // from storage on the next save.
            workbook.origin_xlsx_bytes = None;
            workbook.origin_xlsb_path = None;
        }

        // The formula engine indexes sheets by name and does not support renames in-place.
        // Rebuild to ensure subsequent evaluations resolve the updated sheet name.
        self.rebuild_engine_from_workbook()?;
        self.engine.recalculate();
        let _ = self.refresh_computed_values()?;

        self.dirty = true;
        self.redo_stack.clear();
        Ok(())
    }

    pub fn delete_sheet(&mut self, sheet_id: &str) -> Result<(), AppStateError> {
        let (deleted_id, deleted_name, deleted_index, sheet_order) = {
            let workbook = self.get_workbook()?;
            if workbook.sheets.len() <= 1 {
                return Err(AppStateError::CannotDeleteLastSheet);
            }

            let idx = workbook
                .sheets
                .iter()
                .position(|s| s.id == sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
            let sheet = &workbook.sheets[idx];
            let order = workbook.sheets.iter().map(|s| s.name.clone()).collect::<Vec<_>>();
            (sheet.id.clone(), sheet.name.clone(), idx, order)
        };

        // When the workbook is backed by SQLite persistence, delete the sheet in storage first so
        // a failure doesn't leave the in-memory workbook in a partially updated state. The storage
        // layer also rewrites cross-sheet references to Excel-like `#REF!` semantics.
        if let Some(persistent) = self.persistent.as_mut() {
            persistent
                .memory
                .flush_dirty_pages()
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;

            let sheet_uuid = persistent.sheet_uuid(&deleted_id).ok_or_else(|| {
                AppStateError::Persistence(format!(
                    "missing persistence mapping for sheet id {deleted_id}"
                ))
            })?;
            persistent
                .storage
                .delete_sheet(sheet_uuid)
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;

            // Deleting a sheet triggers workbook-wide reference rewrites in storage. Clear the
            // paging cache so subsequent viewport reads see the updated formulas/values.
            persistent.sheet_map.remove(&deleted_id);
            persistent.memory.clear_cache();
        }

        let remaining_sheet_count = {
            let workbook = self.get_workbook_mut()?;
            if workbook.sheets.len() <= 1 {
                return Err(AppStateError::CannotDeleteLastSheet);
            }

            workbook
                .remove_sheet(&deleted_id)
                .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;

            // Drop print settings for the deleted worksheet.
            workbook
                .print_settings
                .sheets
                .retain(|s| !sheet_name_eq_case_insensitive(&s.sheet_name, &deleted_name));
            workbook
                .original_print_settings
                .sheets
                .retain(|s| !sheet_name_eq_case_insensitive(&s.sheet_name, &deleted_name));

            // Drop any names scoped to the deleted worksheet.
            workbook.defined_names.retain(|name| match name.sheet_id.as_deref() {
                None => true,
                Some(scope) => {
                    !scope.eq_ignore_ascii_case(&deleted_id)
                        && !sheet_name_eq_case_insensitive(scope, &deleted_name)
                }
            });

            // Drop sheet-scoped tables that no longer have a home.
            workbook.tables.retain(|table| {
                !table.sheet_id.eq_ignore_ascii_case(&deleted_id)
                    && !sheet_name_eq_case_insensitive(&table.sheet_id, &deleted_name)
            });

            // Drop preserved worksheet artifacts (pivots/drawings) keyed by the deleted sheet so
            // regeneration-based XLSX saves don't re-attach them to a different sheet.
            if let Some(preserved) = workbook.preserved_pivot_parts.as_mut() {
                let keys: Vec<String> = preserved
                    .sheet_pivot_tables
                    .keys()
                    .filter(|k| sheet_name_eq_case_insensitive(k, &deleted_name))
                    .cloned()
                    .collect();
                for key in keys {
                    preserved.sheet_pivot_tables.remove(&key);
                }

                // The preserved pivot cache definition parts reference worksheet sources by name.
                // We rewrite those names using the original sheet list indices, so keep the
                // indices aligned with structural edits to avoid mapping deleted sheets to the
                // wrong remaining sheet.
                preserved.workbook_sheets.retain(|s| s.index != deleted_index);
                for sheet in &mut preserved.workbook_sheets {
                    if sheet.index > deleted_index {
                        sheet.index = sheet.index.saturating_sub(1);
                    }
                }
            }

            if let Some(preserved) = workbook.preserved_drawing_parts.as_mut() {
                fn remove_preserved_map_keys<V>(
                    map: &mut std::collections::BTreeMap<String, V>,
                    deleted_name: &str,
                ) {
                    let keys: Vec<String> = map
                        .keys()
                        .filter(|k| sheet_name_eq_case_insensitive(k, deleted_name))
                        .cloned()
                        .collect();
                    for key in keys {
                        map.remove(&key);
                    }
                }

                remove_preserved_map_keys(&mut preserved.sheet_drawings, &deleted_name);
                remove_preserved_map_keys(&mut preserved.sheet_pictures, &deleted_name);
                remove_preserved_map_keys(&mut preserved.sheet_ole_objects, &deleted_name);
                remove_preserved_map_keys(&mut preserved.sheet_controls, &deleted_name);
                remove_preserved_map_keys(&mut preserved.sheet_drawing_hfs, &deleted_name);
                // `chart_sheets` represent chartsheet tabs (not worksheets), but remove any
                // preserved entry if it somehow matches the deleted name to avoid index fallback.
                remove_preserved_map_keys(&mut preserved.chart_sheets, &deleted_name);
            }

            // Rewrite formulas that referenced the deleted sheet.
            for sheet in &mut workbook.sheets {
                for ((row, col), cell) in sheet.cells.iter_mut() {
                    let Some(formula) = cell.formula.as_mut() else {
                        continue;
                    };
                    let rewritten = formula_model::rewrite_deleted_sheet_references_in_formula(
                        formula,
                        &deleted_name,
                        &sheet_order,
                    );
                    if rewritten != *formula {
                        *formula = rewritten;
                        sheet.dirty_cells.insert((*row, *col));
                    }
                }
            }

            for name in &mut workbook.defined_names {
                name.refers_to = formula_model::rewrite_deleted_sheet_references_in_formula(
                    &name.refers_to,
                    &deleted_name,
                    &sheet_order,
                );
            }

            // Deleting sheets is a structural XLSX edit. The patch-based save path cannot rewrite
            // workbook.xml sheet order/relationships, so drop origin bytes to force regeneration
            // from storage on the next save.
            workbook.origin_xlsx_bytes = None;
            workbook.origin_xlsb_path = None;

            workbook.sheets.len()
        };

        // Adjust macro runtime context if it points at (or after) the deleted sheet.
        if remaining_sheet_count > 0 {
            let ctx = self.macro_host.runtime_context();
            let mut next_ctx = ctx;

            let fallback_index = if deleted_index < remaining_sheet_count {
                deleted_index
            } else {
                remaining_sheet_count.saturating_sub(1)
            };

            if ctx.active_sheet == deleted_index {
                next_ctx.active_sheet = fallback_index;
                next_ctx.selection = None;
            } else if ctx.active_sheet > deleted_index {
                next_ctx.active_sheet = ctx.active_sheet.saturating_sub(1);
            }

            if let Some(sel) = ctx.selection {
                if sel.sheet == deleted_index {
                    next_ctx.selection = None;
                } else if sel.sheet > deleted_index {
                    let mut next_sel = sel;
                    next_sel.sheet = sel.sheet.saturating_sub(1);
                    next_ctx.selection = Some(next_sel);
                }
            }

            // Avoid `self.get_workbook()` here: it borrows the entire `AppState` immutably for the
            // lifetime of the returned reference, which prevents us from mutably borrowing the
            // disjoint `macro_host` field below.
            let workbook = self
                .workbook
                .as_ref()
                .ok_or(AppStateError::NoWorkbookLoaded)?;
            let macro_host = &mut self.macro_host;
            macro_host.sync_with_workbook(workbook);
            macro_host.set_runtime_context(next_ctx);
        }

        // Drop any pivot registrations that referenced the deleted sheet.
        self.pivots.pivots.retain(|pivot| {
            !(pivot.source_sheet_id.eq_ignore_ascii_case(&deleted_id)
                || pivot.destination.sheet_id.eq_ignore_ascii_case(&deleted_id))
        });

        // Purge undo entries that touch the deleted sheet so undo/redo cannot reference a sheet
        // that no longer exists.
        self.undo_stack.retain(|entry| {
            !entry
                .before
                .iter()
                .any(|snap| snap.sheet_id.eq_ignore_ascii_case(&deleted_id))
                && !entry
                    .after
                    .iter()
                    .any(|snap| snap.sheet_id.eq_ignore_ascii_case(&deleted_id))
        });

        self.redo_stack.clear();

        // Rebuild after deletion so the formula engine drops the removed sheet.
        self.rebuild_engine_from_workbook()?;
        let recalc_changes = self.engine.recalculate_with_value_changes_multi_threaded();
        let _ = self.refresh_computed_values_from_recalc_changes(&recalc_changes)?;

        self.dirty = true;
        Ok(())
    }

    pub fn move_sheet(&mut self, sheet_id: &str, to_index: usize) -> Result<(), AppStateError> {
        let (from_index, sheet_count, ctx, active_sheet_id, selection_sheet_id) = {
            let workbook = self.get_workbook()?;
            let sheet_count = workbook.sheets.len();
            if to_index >= sheet_count {
                return Err(AppStateError::InvalidSheetIndex {
                    to_index,
                    sheet_count,
                });
            }

            let from_index = workbook
                .sheets
                .iter()
                .position(|sheet| sheet.id == sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
            if from_index == to_index {
                return Ok(());
            }

            let ctx = self.macro_host.runtime_context();
            let active_sheet_id = workbook.sheets.get(ctx.active_sheet).map(|s| s.id.clone());
            let selection_sheet_id = ctx
                .selection
                .as_ref()
                .and_then(|sel| workbook.sheets.get(sel.sheet).map(|s| s.id.clone()));

            (from_index, sheet_count, ctx, active_sheet_id, selection_sheet_id)
        };

        // Update persistent storage first so we can fail fast without mutating the in-memory workbook.
        if let Some(persistent) = self.persistent.as_ref() {
            let sheet_uuid = persistent.sheet_uuid(sheet_id).ok_or_else(|| {
                AppStateError::Persistence(format!(
                    "missing persistence mapping for sheet id {sheet_id}"
                ))
            })?;
            persistent
                .storage
                .reorder_sheet(sheet_uuid, to_index as i64)
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;
        }

        {
            let workbook = self.get_workbook_mut()?;
            if sheet_count != workbook.sheets.len() {
                // Shouldn't happen (sheet list mutates between validation and reorder), but avoid
                // panicking on `remove`/`insert`.
                return Err(AppStateError::InvalidSheetIndex {
                    to_index,
                    sheet_count: workbook.sheets.len(),
                });
            }

            let sheet = workbook.sheets.remove(from_index);
            workbook.sheets.insert(to_index, sheet);

            // Keep preserved pivot cache sheet indices aligned with the moved sheet so
            // index-based worksheetSource rewrites remain correct after move+rename sequences.
            if let Some(preserved) = workbook.preserved_pivot_parts.as_mut() {
                for sheet in &mut preserved.workbook_sheets {
                    if sheet.index == from_index {
                        sheet.index = to_index;
                    } else if from_index < to_index {
                        // Moving forward: intermediate sheets shift left by 1.
                        if sheet.index > from_index && sheet.index <= to_index {
                            sheet.index = sheet.index.saturating_sub(1);
                        }
                    } else {
                        // Moving backward: intermediate sheets shift right by 1.
                        if sheet.index >= to_index && sheet.index < from_index {
                            sheet.index = sheet.index.saturating_add(1);
                        }
                    }
                }
            }

            // NOTE: Reordering sheets is a structural edit for XLSB inputs. The `.xlsb` writer
            // cannot currently reorder workbook metadata, so treat this as a conversion to XLSX
            // on the next save (consistent with other structural edits like add/delete/rename).
            if workbook.origin_xlsb_path.is_some() {
                workbook.origin_xlsb_path = None;
            }
        }

        // The macro runtime context stores sheet indices; keep them stable so it continues to
        // point at the same sheet(s) after reordering.
        {
            // Avoid `self.get_workbook()` here: it borrows the entire `AppState` immutably for the
            // lifetime of the returned reference, which prevents us from mutably borrowing the
            // disjoint `macro_host` field below.
            let workbook = self
                .workbook
                .as_ref()
                .ok_or(AppStateError::NoWorkbookLoaded)?;
            let mut new_ctx = ctx;

            if let Some(active_sheet_id) = active_sheet_id {
                if let Some(idx) = workbook
                    .sheets
                    .iter()
                    .position(|s| s.id == active_sheet_id)
                {
                    new_ctx.active_sheet = idx;
                }
            }

            if let (Some(sel), Some(selection_sheet_id)) =
                (new_ctx.selection.as_mut(), selection_sheet_id)
            {
                if let Some(idx) = workbook
                    .sheets
                    .iter()
                    .position(|s| s.id == selection_sheet_id)
                {
                    sel.sheet = idx;
                }
            }

            let macro_host = &mut self.macro_host;
            macro_host.sync_with_workbook(workbook);
            macro_host.set_runtime_context(new_ctx);
        }

        // Sheet order affects some Excel semantics (e.g. 3D references like `Sheet1:Sheet3!A1`).
        // Rebuild to ensure evaluations reflect the updated sheet ordering.
        self.rebuild_engine_from_workbook()?;
        self.engine.recalculate();
        let _ = self.refresh_computed_values()?;

        self.dirty = true;
        self.redo_stack.clear();
        Ok(())
    }

    pub fn set_sheet_visibility(
        &mut self,
        sheet_id: &str,
        visibility: SheetVisibility,
    ) -> Result<(), AppStateError> {
        let (current, visible_count) = {
            let workbook = self.get_workbook()?;
            let sheet = workbook
                .sheet(sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
            let current = sheet.visibility;
            if current == visibility {
                return Ok(());
            }
            let visible_count = workbook
                .sheets
                .iter()
                .filter(|s| matches!(s.visibility, SheetVisibility::Visible))
                .count();
            (current, visible_count)
        };

        // Note: the desktop UI only exposes "visible" / "hidden", but we still allow setting
        // `veryHidden` from the host APIs so:
        // - imported workbooks can round-trip their original visibility state, and
        // - advanced integrations (extensions, automation) can opt into the VBA-style behavior.
        // Excel invariant: cannot hide the last visible sheet.
        if matches!(current, SheetVisibility::Visible)
            && !matches!(visibility, SheetVisibility::Visible)
            && visible_count <= 1
        {
            return Err(AppStateError::WhatIf(
                "cannot hide the last visible sheet".to_string(),
            ));
        }

        // Persist first so we don't leave the in-memory workbook in a partially-updated state if
        // the storage backend rejects the write.
        if let Some(persistent) = self.persistent.as_ref() {
            let sheet_uuid = persistent.sheet_uuid(sheet_id).ok_or_else(|| {
                AppStateError::Persistence(format!(
                    "missing persistence mapping for sheet id {sheet_id}"
                ))
            })?;
            let storage_visibility = match visibility {
                SheetVisibility::Visible => formula_storage::SheetVisibility::Visible,
                SheetVisibility::Hidden => formula_storage::SheetVisibility::Hidden,
                SheetVisibility::VeryHidden => formula_storage::SheetVisibility::VeryHidden,
            };
            persistent
                .storage
                .set_sheet_visibility(sheet_uuid, storage_visibility)
                .map_err(|e| match e {
                    formula_storage::StorageError::SheetNotFound(_) => {
                        AppStateError::UnknownSheet(sheet_id.to_string())
                    }
                    other => AppStateError::Persistence(other.to_string()),
                })?;
        }

        {
            let workbook = self.get_workbook_mut()?;
            {
                let sheet = workbook
                    .sheet_mut(sheet_id)
                    .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
                sheet.visibility = visibility;
            }

            // Sheet visibility is workbook.xml metadata. Our patch-based XLSX save path can
            // rewrite this, but XLSB round-trip saves cannot update workbook metadata. Drop XLSB
            // provenance so subsequent saves use the XLSX path.
            workbook.origin_xlsb_path = None;
        }

        self.dirty = true;
        self.redo_stack.clear();
        Ok(())
    }

    pub fn set_sheet_tab_color(
        &mut self,
        sheet_id: &str,
        tab_color: Option<TabColor>,
    ) -> Result<(), AppStateError> {
        let current = {
            let workbook = self.get_workbook()?;
            let sheet = workbook
                .sheet(sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
            sheet.tab_color.clone()
        };

        let tab_color = match tab_color {
            None => None,
            Some(mut color) => {
                // Normalize `rgb` values to Excel-compatible uppercased ARGB hex (AARRGGBB). Leave
                // theme/indexed/tint/auto values as-is so we can persist and round-trip workbook
                // metadata that isn't directly settable via the desktop UI.
                if let Some(rgb) = color.rgb.as_deref() {
                    let trimmed = rgb.trim();
                    if trimmed.is_empty() {
                        color.rgb = None;
                    } else {
                        let hex = trimmed.strip_prefix('#').unwrap_or(trimmed);
                        if hex.len() == 6 && hex.chars().all(|c| c.is_ascii_hexdigit()) {
                            color.rgb = Some(format!("FF{}", hex.to_ascii_uppercase()));
                        } else if hex.len() == 8 && hex.chars().all(|c| c.is_ascii_hexdigit()) {
                            color.rgb = Some(hex.to_ascii_uppercase());
                        } else {
                            return Err(AppStateError::WhatIf(
                                "tab color rgb must be 6-digit (RRGGBB) or 8-digit ARGB (AARRGGBB) hex"
                                    .to_string(),
                            ));
                        }
                    }
                }

                if let Some(tint) = color.tint {
                    if !tint.is_finite() || !(-1.0..=1.0).contains(&tint) {
                        return Err(AppStateError::WhatIf(
                            "tab color tint must be a finite number between -1.0 and 1.0"
                                .to_string(),
                        ));
                    }
                }
                // Treat an all-empty payload as clearing the tab color.
                if color.rgb.is_none()
                    && color.theme.is_none()
                    && color.indexed.is_none()
                    && color.tint.is_none()
                    && color.auto.is_none()
                {
                    None
                } else {
                    Some(color)
                }
            }
        };

        if current == tab_color {
            return Ok(());
        }

        // Persist first so we don't leave the in-memory workbook in a partially-updated state if
        // the storage backend rejects the write.
        if let Some(persistent) = self.persistent.as_ref() {
            let sheet_uuid = persistent.sheet_uuid(sheet_id).ok_or_else(|| {
                AppStateError::Persistence(format!(
                    "missing persistence mapping for sheet id {sheet_id}"
                ))
            })?;

            persistent
                .storage
                .set_sheet_tab_color(sheet_uuid, tab_color.as_ref())
                .map_err(|e| match e {
                    formula_storage::StorageError::SheetNotFound(_) => {
                        AppStateError::UnknownSheet(sheet_id.to_string())
                    }
                    other => AppStateError::Persistence(other.to_string()),
                })?;
        }

        {
            let workbook = self.get_workbook_mut()?;
            {
                let sheet = workbook
                    .sheet_mut(sheet_id)
                    .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
                sheet.tab_color = tab_color;
            }

            // Tab color lives in worksheet metadata (`sheetPr/tabColor`). Our patch-based XLSX save
            // path can rewrite this, but XLSB round-trip saves cannot update worksheet metadata.
            // Drop XLSB provenance so subsequent saves use the XLSX path.
            workbook.origin_xlsb_path = None;
        }

        self.dirty = true;
        self.redo_stack.clear();
        Ok(())
    }

    pub fn set_range_number_format(
        &mut self,
        sheet_id: &str,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        number_format: Option<String>,
    ) -> Result<(), AppStateError> {
        if start_row > end_row || start_col > end_col {
            return Err(AppStateError::InvalidRange {
                start_row,
                start_col,
                end_row,
                end_col,
            });
        }

        let row_count = end_row
            .checked_sub(start_row)
            .and_then(|d| d.checked_add(1))
            .unwrap_or(usize::MAX);
        let col_count = end_col
            .checked_sub(start_col)
            .and_then(|d| d.checked_add(1))
            .unwrap_or(usize::MAX);

        if row_count > MAX_RANGE_DIM || col_count > MAX_RANGE_DIM {
            return Err(AppStateError::RangeDimensionTooLarge {
                rows: row_count,
                cols: col_count,
                limit: MAX_RANGE_DIM,
            });
        }

        let cell_count = (row_count as u128) * (col_count as u128);
        if cell_count > MAX_RANGE_CELLS_PER_CALL as u128 {
            return Err(AppStateError::RangeTooLarge {
                rows: row_count,
                cols: col_count,
                limit: MAX_RANGE_CELLS_PER_CALL,
            });
        }

        let (sheet_name, coords) = {
            let workbook = self.get_workbook_mut()?;
            let sheet = workbook
                .sheet_mut(sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
            let sheet_name = sheet.name.clone();

            let mut coords = Vec::with_capacity(row_count.saturating_mul(col_count));
            for row in start_row..=end_row {
                for col in start_col..=end_col {
                    let mut cell = sheet
                        .cells
                        .get(&(row, col))
                        .cloned()
                        .unwrap_or_else(Cell::empty);
                    cell.number_format = number_format.clone();
                    sheet.set_cell(row, col, cell);
                    coords.push((row, col));
                }
            }

            (sheet_name, coords)
        };

        let mut style_ids_by_format: HashMap<String, u32> = HashMap::new();
        let style_id = number_format
            .as_deref()
            .map(str::trim)
            .filter(|s| !s.is_empty())
            .map(|fmt| {
                if let Some(existing) = style_ids_by_format.get(fmt) {
                    *existing
                } else {
                    let fmt = fmt.to_string();
                    let style_id = self.engine.intern_style(formula_model::Style {
                        number_format: Some(fmt.clone()),
                        ..Default::default()
                    });
                    style_ids_by_format.insert(fmt, style_id);
                    style_id
                }
            })
            .unwrap_or(0);

        for (row, col) in coords {
            let addr = coord_to_a1(row, col);
            let _ = self.engine.set_cell_style_id(&sheet_name, &addr, style_id);
        }

        self.dirty = true;
        self.redo_stack.clear();
        Ok(())
    }

    #[cfg(any(feature = "desktop", test))]
    pub(crate) fn apply_sheet_formatting_deltas_to_engine(
        &mut self,
        payload: &crate::commands::ApplySheetFormattingDeltasRequest,
    ) -> Result<(), AppStateError> {
        use crate::commands::{
            LimitedSheetCellFormatDeltas, LimitedSheetColFormatDeltas,
            LimitedSheetFormatRunsByColDeltas, LimitedSheetRowFormatDeltas,
        };

        let workbook = self.get_workbook()?;
        let sheet = resolve_sheet_case_insensitive(workbook, &payload.sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(payload.sheet_id.clone()))?;
        let sheet_name = sheet.name.clone();

        // Sheet default formatting.
        //
        // UI formatting payloads can include camelCase keys (e.g. `numberFormat`) and "explicit
        // clear" semantics (e.g. `numberFormat: null` overriding an imported `number_format`).
        // Use `ui_style_to_model_style` so `CELL()` formatting queries match the UI interpretation.
        if let Some(default_format) = payload.default_format.as_ref() {
            let style_id = match default_format {
                None => None,
                Some(format) if format.is_null() => None,
                Some(format) => {
                    let id = self.engine.intern_style(ui_style_to_model_style(format));
                    (id != 0).then_some(id)
                }
            };
            self.engine.set_sheet_default_style_id(&sheet_name, style_id);
        }

        if let Some(LimitedSheetRowFormatDeltas(deltas)) = payload.row_formats.as_ref() {
            for delta in deltas {
                if delta.row < 0 {
                    continue;
                }
                let Ok(row) = u32::try_from(delta.row) else {
                    continue;
                };
                let style_id = if delta.format.is_null() {
                    None
                } else {
                    let id = self.engine.intern_style(ui_style_to_model_style(&delta.format));
                    (id != 0).then_some(id)
                };
                self.engine.set_row_style_id(&sheet_name, row, style_id);
            }
        }

        if let Some(LimitedSheetColFormatDeltas(deltas)) = payload.col_formats.as_ref() {
            for delta in deltas {
                if delta.col < 0 {
                    continue;
                }
                let Ok(col) = u32::try_from(delta.col) else {
                    continue;
                };
                let style_id = if delta.format.is_null() {
                    None
                } else {
                    let id = self.engine.intern_style(ui_style_to_model_style(&delta.format));
                    (id != 0).then_some(id)
                };
                self.engine.set_col_style_id(&sheet_name, col, style_id);
            }
        }

        if let Some(LimitedSheetFormatRunsByColDeltas(deltas)) = payload.format_runs_by_col.as_ref()
        {
            for delta in deltas {
                if delta.col < 0 {
                    continue;
                }
                let Ok(col) = u32::try_from(delta.col) else {
                    continue;
                };
                if col >= formula_model::EXCEL_MAX_COLS {
                    continue;
                }

                let mut runs: Vec<FormatRun> = Vec::new();
                for run in delta.runs.0.iter() {
                    if run.start_row < 0 || run.end_row_exclusive <= run.start_row {
                        continue;
                    }
                    let Ok(start_row) = u32::try_from(run.start_row) else {
                        continue;
                    };
                    let Ok(end_row_exclusive) = u32::try_from(run.end_row_exclusive) else {
                        continue;
                    };
                    if end_row_exclusive <= start_row {
                        continue;
                    }
                    let style_id = if run.format.is_null() {
                        0
                    } else {
                        self.engine.intern_style(ui_style_to_model_style(&run.format))
                    };
                    runs.push(FormatRun {
                        start_row,
                        end_row_exclusive,
                        style_id,
                    });
                }
                let _ = self.engine.set_format_runs_by_col(&sheet_name, col, runs);
            }
        }

        if let Some(LimitedSheetCellFormatDeltas(deltas)) = payload.cell_formats.as_ref() {
            for delta in deltas {
                if delta.row < 0 || delta.col < 0 {
                    continue;
                }
                let Ok(row) = usize::try_from(delta.row) else {
                    continue;
                };
                let Ok(col) = usize::try_from(delta.col) else {
                    continue;
                };
                let addr = coord_to_a1(row, col);
                if delta.format.is_null() {
                    let _ = self.engine.set_cell_style_id(&sheet_name, &addr, 0);
                    continue;
                }

                let style_id = self
                    .engine
                    .intern_style(ui_style_to_model_style(&delta.format));
                let _ = self.engine.set_cell_style_id(&sheet_name, &addr, style_id);
            }
        }

        Ok(())
    }

    pub fn load_workbook(&mut self, mut workbook: Workbook) -> WorkbookInfoData {
        workbook.ensure_sheet_ids();
        self.workbook = Some(workbook);
        self.persistent = None;
        self.engine = FormulaEngine::new();
        self.scenario_manager = ScenarioManager::new();
        self.macro_host.invalidate();
        self.pivots = PivotManager::default();

        // Best effort: rebuild and calculate. Unsupported formulas become #NAME? via the engine.
        let _ = self.rebuild_engine_from_workbook();
        let recalc_changes = self.engine.recalculate_with_value_changes_multi_threaded();
        let _ = self.refresh_computed_values_from_recalc_changes(&recalc_changes);

        self.dirty = false;
        self.undo_stack.clear();
        self.redo_stack.clear();
        self.workbook_info()
            .expect("workbook_info should succeed right after load")
    }

    pub fn load_workbook_persistent(
        &mut self,
        workbook: Workbook,
        location: WorkbookPersistenceLocation,
    ) -> Result<WorkbookInfoData, AppStateError> {
        if let WorkbookPersistenceLocation::OnDisk(path) = &location {
            if let Some(parent) = path.parent() {
                std::fs::create_dir_all(parent)
                    .map_err(|e| AppStateError::Persistence(e.to_string()))?;
            }
        }

        let storage = open_storage(&location).map_err(|e| AppStateError::Persistence(e.to_string()))?;

        let existing = storage
            .list_workbooks()
            .map_err(|e| AppStateError::Persistence(e.to_string()))?;

        let (workbook_id, sheet_metas, workbook_for_state) = if let Some(existing_meta) = existing.first()
        {
            let model = storage
                .export_model_workbook(existing_meta.id)
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;
            let mut recovered =
                workbook_from_model(&model).map_err(|e| AppStateError::Persistence(e.to_string()))?;

            // Carry over workbook-level metadata captured from the file we opened.
            recovered.path = workbook.path.clone();
            recovered.origin_path = workbook.origin_path.clone();
            recovered.origin_xlsx_bytes = workbook.origin_xlsx_bytes.clone();
            recovered.origin_xlsb_path = workbook.origin_xlsb_path.clone();
            recovered.power_query_xml = workbook.power_query_xml.clone();
            recovered.original_power_query_xml = workbook.original_power_query_xml.clone();
            recovered.vba_project_bin = workbook.vba_project_bin.clone();
            recovered.vba_project_signature_bin = workbook.vba_project_signature_bin.clone();
            recovered.macro_fingerprint = workbook.macro_fingerprint.clone();
            recovered.preserved_drawing_parts = workbook.preserved_drawing_parts.clone();
            recovered.preserved_pivot_parts = workbook.preserved_pivot_parts.clone();
            recovered.theme_palette = workbook.theme_palette.clone();
            recovered.print_settings = workbook.print_settings.clone();
            recovered.original_print_settings = workbook.original_print_settings.clone();
            recovered.defined_names = workbook.defined_names.clone();
            recovered.tables = workbook.tables.clone();

            // Keep columnar-backed sheet data from the freshly loaded workbook (e.g. CSV imports).
            for recovered_sheet in &mut recovered.sheets {
                if let Some(src) = workbook
                    .sheets
                    .iter()
                    .find(|s| sheet_name_eq_case_insensitive(&s.name, &recovered_sheet.name))
                {
                    recovered_sheet.columnar = src.columnar.clone();
                }
            }

            // Preserve XLSB sheet ordinals for round-tripping after crash recovery / autosave
            // restores. These indices are used as a fallback when a sheet was renamed in-app.
            for (idx, recovered_sheet) in recovered.sheets.iter_mut().enumerate() {
                if recovered_sheet.origin_ordinal.is_some() {
                    continue;
                }
                recovered_sheet.origin_ordinal = workbook
                    .sheets
                    .get(idx)
                    .and_then(|sheet| sheet.origin_ordinal);
            }

            let sheet_metas = storage
                .list_sheets(existing_meta.id)
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;
            (existing_meta.id, sheet_metas, recovered)
        } else {
            let model = if let Some(origin_bytes) = workbook.origin_xlsx_bytes.as_deref() {
                formula_xlsx::read_workbook_from_reader(Cursor::new(origin_bytes))
                    .map_err(|e| AppStateError::Persistence(e.to_string()))?
            } else {
                workbook_to_model(&workbook)
                    .map_err(|e| AppStateError::Persistence(e.to_string()))?
            };

            let name = workbook
                .origin_path
                .as_deref()
                .or(workbook.path.as_deref())
                .and_then(|p| Path::new(p).file_stem().and_then(|s| s.to_str()))
                .filter(|s| !s.trim().is_empty())
                .unwrap_or("Workbook");

            let workbook_meta = storage
                .import_model_workbook(&model, ImportModelWorkbookOptions::new(name))
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;
            let sheet_metas = storage
                .list_sheets(workbook_meta.id)
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;

            seed_sheet_formatting_metadata_from_model(&storage, &sheet_metas, &model)
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;
            (workbook_meta.id, sheet_metas, workbook)
        };

        let info = self.load_workbook(workbook_for_state);

        let memory = open_memory_manager(storage.clone());
        let autosave = tokio::runtime::Handle::try_current()
            .ok()
            .map(|_| Arc::new(AutoSaveManager::spawn(memory.clone(), AutoSaveConfig::default())));

        let mut sheet_map = HashMap::<String, Uuid>::new();
        if let Some(workbook) = self.workbook.as_ref() {
            for sheet in &workbook.sheets {
                if let Some(meta) = sheet_metas
                    .iter()
                    .find(|m| sheet_name_eq_case_insensitive(&m.name, &sheet.name))
                {
                    sheet_map.insert(sheet.id.clone(), meta.id);
                }
            }
        }

        self.persistent = Some(PersistentWorkbookState {
            location,
            storage,
            memory,
            autosave,
            workbook_id,
            sheet_map,
        });

        // Best-effort: apply persisted UI formatting metadata to the engine's style table so
        // functions like `CELL("protect")` can consult it once style-aware behavior is added.
        //
        // Formatting metadata is stored in the persistent sheet metadata as UI-friendly JSON
        // (`formula_ui_formatting`, typically camelCase). Convert it to `formula_model::Style`
        // and store the resulting style ids in the engine.
        //
        // Note: `load_workbook` already performed an initial recalc pass. If we applied any style
        // metadata here, trigger one additional recalc so volatile formulas (e.g. `CELL()`) can
        // observe the new metadata.
        if self.apply_persistent_ui_formatting_metadata_to_engine() {
            let recalc_changes = self.engine.recalculate_with_value_changes_multi_threaded();
            let _ = self.refresh_computed_values_from_recalc_changes(&recalc_changes);
        }

        Ok(info)
    }

    pub(crate) fn apply_ui_row_format_delta_to_engine(
        &mut self,
        sheet_id: &str,
        row: i64,
        format: &JsonValue,
    ) -> Result<(), AppStateError> {
        if row < 0 {
            return Ok(());
        }
        let Ok(row_0based) = u32::try_from(row) else {
            return Ok(());
        };
        if row_0based >= i32::MAX as u32 {
            return Ok(());
        }

        let sheet_name = {
            let workbook = self.get_workbook()?;
            workbook
                .sheet(sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?
                .name
                .clone()
        };

        let style_id = if format.is_null() {
            None
        } else {
            let id = self.engine.intern_style(ui_style_to_model_style(format));
            (id != 0).then_some(id)
        };
        self.engine.set_row_style_id(&sheet_name, row_0based, style_id);
        Ok(())
    }

    pub(crate) fn apply_ui_col_format_delta_to_engine(
        &mut self,
        sheet_id: &str,
        col: i64,
        format: &JsonValue,
    ) -> Result<(), AppStateError> {
        if col < 0 {
            return Ok(());
        }
        let Ok(col_0based) = u32::try_from(col) else {
            return Ok(());
        };
        if col_0based >= formula_model::EXCEL_MAX_COLS {
            return Ok(());
        }

        let sheet_name = {
            let workbook = self.get_workbook()?;
            workbook
                .sheet(sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?
                .name
                .clone()
        };

        let style_id = if format.is_null() {
            None
        } else {
            let id = self.engine.intern_style(ui_style_to_model_style(format));
            (id != 0).then_some(id)
        };
        self.engine.set_col_style_id(&sheet_name, col_0based, style_id);
        Ok(())
    }

    pub(crate) fn apply_ui_cell_format_delta_to_engine(
        &mut self,
        sheet_id: &str,
        row: i64,
        col: i64,
        format: &JsonValue,
    ) -> Result<(), AppStateError> {
        if row < 0 || col < 0 {
            return Ok(());
        }
        let Ok(row_0based) = u32::try_from(row) else {
            return Ok(());
        };
        if row_0based >= i32::MAX as u32 {
            return Ok(());
        }
        let Ok(col_0based) = u32::try_from(col) else {
            return Ok(());
        };
        if col_0based >= formula_model::EXCEL_MAX_COLS {
            return Ok(());
        }

        let sheet_name = {
            let workbook = self.get_workbook()?;
            workbook
                .sheet(sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?
                .name
                .clone()
        };

        let style_id = if format.is_null() {
            0
        } else {
            self.engine.intern_style(ui_style_to_model_style(format))
        };
        let addr = coord_to_a1(row_0based as usize, col_0based as usize);
        self.engine
            .set_cell_style_id(&sheet_name, &addr, style_id)
            .map_err(|e| AppStateError::Engine(e.to_string()))?;
        Ok(())
    }

    fn apply_persistent_ui_formatting_metadata_to_engine(&mut self) -> bool {
        fn parse_non_negative_i64(raw: Option<&JsonValue>) -> Option<i64> {
            let v = raw?;
            let n = v
                .as_i64()
                .or_else(|| v.as_u64().and_then(|u| i64::try_from(u).ok()))?;
            (n >= 0).then_some(n)
        }

        let Some(storage) = self.persistent_storage() else {
            return false;
        };
        // Explicit `Vec<String>` annotation avoids type inference falling back to an unsized `str`
        // element type (which then fails in the `for sheet_id in sheet_ids` loop below).
        let sheet_ids: Vec<String> = match self.get_workbook() {
            Ok(workbook) => workbook.sheets.iter().map(|s| s.id.clone()).collect(),
            Err(_) => return false,
        };

        let mut changed = false;
        for sheet_id in sheet_ids {
            let sheet_name = match self.get_workbook() {
                Ok(workbook) => workbook
                    .sheet(&sheet_id)
                    .map(|s| s.name.clone())
                    .unwrap_or_default(),
                Err(_) => String::new(),
            };
            if sheet_name.is_empty() {
                continue;
            }

            let sheet_uuid = match self.persistent_sheet_uuid(&sheet_id) {
                Ok(v) => v,
                Err(_) => continue,
            };
            let meta = match storage.get_sheet_meta(sheet_uuid) {
                Ok(v) => v,
                Err(_) => continue,
            };
            let raw = meta
                .metadata
                .as_ref()
                .and_then(|metadata| metadata.get(FORMULA_UI_FORMATTING_METADATA_KEY));
            let Some(obj) = raw.and_then(|v| v.as_object()) else {
                continue;
            };

            // Sheet default format.
            //
            // `defaultFormat` is always present in persisted UI metadata, but is often `null`.
            // Avoid flagging a change (and triggering a redundant recalc) unless we actually apply
            // a non-default style.
            if let Some(default_format) = obj
                .get("defaultFormat")
                .or_else(|| obj.get("default_format"))
            {
                if default_format.is_null() {
                    self.engine.set_sheet_default_style_id(&sheet_name, None);
                } else {
                    let id = self
                        .engine
                        .intern_style(ui_style_to_model_style(default_format));
                    let style_id = (id != 0).then_some(id);
                    self.engine
                        .set_sheet_default_style_id(&sheet_name, style_id);
                    if style_id.is_some() {
                        changed = true;
                    }
                }
            }

            // Range-run formats.
            if let Some(cols) = obj
                .get("formatRunsByCol")
                .or_else(|| obj.get("format_runs_by_col"))
                .and_then(|v| v.as_array())
            {
                for entry in cols {
                    let Some(col) =
                        parse_non_negative_i64(entry.get("col").or_else(|| entry.get("index")))
                    else {
                        continue;
                    };
                    let Ok(col_0based) = u32::try_from(col) else {
                        continue;
                    };
                    if col_0based >= formula_model::EXCEL_MAX_COLS {
                        continue;
                    }

                    let Some(runs_array) = entry.get("runs").and_then(|v| v.as_array()) else {
                        continue;
                    };

                    let mut runs: Vec<FormatRun> = Vec::new();
                    let mut any_non_default = false;
                    for run in runs_array {
                        let Some(start_row) = parse_non_negative_i64(
                            run.get("startRow").or_else(|| run.get("start_row")),
                        ) else {
                            continue;
                        };
                        let Some(end_row_exclusive) = parse_non_negative_i64(
                            run.get("endRowExclusive")
                                .or_else(|| run.get("end_row_exclusive")),
                        )
                        .or_else(|| {
                            parse_non_negative_i64(run.get("endRow").or_else(|| run.get("end_row")))
                                .map(|end| end.saturating_add(1))
                        }) else {
                            continue;
                        };
                        if end_row_exclusive <= start_row {
                            continue;
                        }

                        let Ok(start_row) = u32::try_from(start_row) else {
                            continue;
                        };
                        let Ok(end_row_exclusive) = u32::try_from(end_row_exclusive) else {
                            continue;
                        };
                        if start_row >= i32::MAX as u32 || end_row_exclusive > i32::MAX as u32 {
                            continue;
                        }

                        let style_id = match run.get("format") {
                            Some(format) if !format.is_null() => {
                                let id = self
                                    .engine
                                    .intern_style(ui_style_to_model_style(format));
                                if id != 0 {
                                    any_non_default = true;
                                }
                                id
                            }
                            _ => 0,
                        };
                        runs.push(FormatRun {
                            start_row,
                            end_row_exclusive,
                            style_id,
                        });
                    }

                    if any_non_default {
                        let _ = self
                            .engine
                            .set_format_runs_by_col(&sheet_name, col_0based, runs);
                        changed = true;
                    }
                }
            }

            // Row formats.
            if let Some(rows) = obj
                .get("rowFormats")
                .or_else(|| obj.get("row_formats"))
                .and_then(|v| v.as_array())
            {
                for entry in rows {
                    let Some(row) =
                        parse_non_negative_i64(entry.get("row").or_else(|| entry.get("index")))
                    else {
                        continue;
                    };
                    let Some(format) = entry.get("format") else {
                        continue;
                    };
                    if self
                        .apply_ui_row_format_delta_to_engine(&sheet_id, row, format)
                        .is_ok()
                    {
                        changed = true;
                    }
                }
            }

            // Col formats.
            if let Some(cols) = obj
                .get("colFormats")
                .or_else(|| obj.get("col_formats"))
                .and_then(|v| v.as_array())
            {
                for entry in cols {
                    let Some(col) =
                        parse_non_negative_i64(entry.get("col").or_else(|| entry.get("index")))
                    else {
                        continue;
                    };
                    let Some(format) = entry.get("format") else {
                        continue;
                    };
                    if self
                        .apply_ui_col_format_delta_to_engine(&sheet_id, col, format)
                        .is_ok()
                    {
                        changed = true;
                    }
                }
            }

            // Cell formats.
            if let Some(cells) = obj
                .get("cellFormats")
                .or_else(|| obj.get("cell_formats"))
                .and_then(|v| v.as_array())
            {
                for entry in cells {
                    let Some(row) = parse_non_negative_i64(entry.get("row")) else {
                        continue;
                    };
                    let Some(col) = parse_non_negative_i64(entry.get("col")) else {
                        continue;
                    };
                    let Some(format) = entry.get("format") else {
                        continue;
                    };
                    if self
                        .apply_ui_cell_format_delta_to_engine(&sheet_id, row, col, format)
                        .is_ok()
                    {
                        changed = true;
                    }
                }
            }
        }

        changed
    }

    pub fn autosave_manager(&self) -> Option<Arc<AutoSaveManager>> {
        self.persistent.as_ref().and_then(|p| p.autosave.clone())
    }

    pub fn persistent_workbook_id(&self) -> Option<Uuid> {
        self.persistent.as_ref().map(|p| p.workbook_id)
    }

    pub fn persistent_sheet_uuid(&self, sheet_id: &str) -> Result<Uuid, AppStateError> {
        let persistent = self.persistent.as_ref().ok_or_else(|| {
            AppStateError::Persistence("workbook is not backed by persistent storage".to_string())
        })?;
        persistent.sheet_uuid(sheet_id).ok_or_else(|| {
            AppStateError::Persistence(format!(
                "missing persistence mapping for sheet id {sheet_id}"
            ))
        })
    }

    pub fn persistent_storage(&self) -> Option<formula_storage::Storage> {
        self.persistent.as_ref().map(|p| p.storage.clone())
    }

    pub fn persistent_memory_manager(&self) -> Option<formula_storage::MemoryManager> {
        self.persistent.as_ref().map(|p| p.memory.clone())
    }

    fn is_xlsx_family_extension(ext: &str) -> bool {
        ext.eq_ignore_ascii_case("xlsx")
            || ext.eq_ignore_ascii_case("xlsm")
            || ext.eq_ignore_ascii_case("xltx")
            || ext.eq_ignore_ascii_case("xltm")
            || ext.eq_ignore_ascii_case("xlam")
    }

    fn is_macro_free_extension(ext: &str) -> bool {
        ext.eq_ignore_ascii_case("xlsx") || ext.eq_ignore_ascii_case("xltx")
    }

    pub fn mark_saved(
        &mut self,
        new_path: Option<String>,
        new_origin_xlsx_bytes: Option<Arc<[u8]>>,
    ) -> Result<(), AppStateError> {
        let requested_path = new_path.clone();
        let is_real_save = requested_path.is_some() || new_origin_xlsx_bytes.is_some();
        let metadata_before = self
            .workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)
            .map(workbook_file_metadata)?;
        let (directory, filename) = {
            let workbook = self
                .workbook
                .as_mut()
                .ok_or(AppStateError::NoWorkbookLoaded)?;

            if let Some(path) = new_path {
                workbook.path = Some(path);
            }

            let ext = workbook
                .path
                .as_deref()
                .and_then(|p| std::path::Path::new(p).extension().and_then(|s| s.to_str()));

            // Allow operators to tighten the retention cap via `FORMULA_MAX_ORIGIN_XLSX_BYTES`.
            // Never allow relaxing it above the built-in default (even in debug builds).
            let max_origin_xlsx_bytes =
                crate::resource_limits::max_origin_xlsx_bytes().min(MAX_ORIGIN_XLSX_BYTES);

            if ext.is_some_and(|ext| ext.eq_ignore_ascii_case("xlsb")) {
                workbook.origin_xlsb_path = workbook.path.clone();
                workbook.origin_xlsx_bytes = None;
            } else if ext.is_some_and(Self::is_xlsx_family_extension) {
                workbook.origin_xlsb_path = None;
                if let Some(bytes) = new_origin_xlsx_bytes {
                    if bytes.len() <= max_origin_xlsx_bytes {
                        workbook.origin_xlsx_bytes = Some(bytes);
                    } else {
                        // Defense-in-depth: avoid retaining arbitrarily large baseline snapshots in
                        // memory after save. Dropping the baseline forces subsequent saves to use the
                        // regeneration-based path (instead of patching from stale bytes).
                        eprintln!(
                            "[save] dropping origin_xlsx_bytes baseline after save: snapshot {} bytes exceeds origin retention limit ({})",
                            bytes.len(),
                            max_origin_xlsx_bytes
                        );
                        workbook.origin_xlsx_bytes = None;
                    }
                }
            } else if let Some(bytes) = new_origin_xlsx_bytes {
                if bytes.len() <= max_origin_xlsx_bytes {
                    workbook.origin_xlsx_bytes = Some(bytes);
                } else {
                    eprintln!(
                        "[save] dropping origin_xlsx_bytes baseline after save: snapshot {} bytes exceeds origin retention limit ({})",
                        bytes.len(),
                        max_origin_xlsx_bytes
                    );
                    workbook.origin_xlsx_bytes = None;
                }
            }

            // Saving establishes a new baseline for "net" changes. Clear the per-cell baseline so
            // subsequent edits are tracked against this saved state (not the previously opened or
            // previously saved workbook bytes).
            workbook.cell_input_baseline.clear();
            if is_real_save {
                workbook.original_print_settings = workbook.print_settings.clone();
                workbook.original_power_query_xml = workbook.power_query_xml.clone();
            }
            for sheet in &mut workbook.sheets {
                sheet.clear_dirty_cells();
            }

            // If the saved file is macro-free (`.xlsx`/`.xltx`), macros are not preserved; clear any
            // in-memory macro payloads so the UI doesn't continue to treat the workbook as
            // macro-enabled.
            if ext.is_some_and(Self::is_macro_free_extension) {
                workbook.vba_project_bin = None;
                workbook.vba_project_signature_bin = None;
                workbook.macro_fingerprint = None;
            }
            self.dirty = false;

            workbook_file_metadata(workbook)
        };

        if metadata_before != (directory.clone(), filename.clone()) {
            // Update workbook file metadata on the formula engine so worksheet information functions
            // like `CELL("filename")` and `INFO("directory")` reflect the latest save path.
            self.engine
                .set_workbook_file_metadata(directory.as_deref(), filename.as_deref());
            let changes = self.engine.recalculate_with_value_changes_multi_threaded();
            let _ = self.refresh_computed_values_from_recalc_changes(&changes)?;
        }

        // If the user saved under a new path (Save As), re-key the autosave database to the new file
        // identity so crash recovery uses the correct autosave DB for subsequent opens.
        //
        // This is best-effort: a failure to re-key should never fail the save itself.
        if let Some(saved_path) = requested_path.as_deref() {
            let _ = self.rekey_autosave_db_after_save(saved_path);
        }

        Ok(())
    }

    fn rekey_autosave_db_after_save(&mut self, saved_path: &str) -> Result<(), AppStateError> {
        let Some(persistent) = self.persistent.as_mut() else {
            return Ok(());
        };
        let current_db_path = match &persistent.location {
            WorkbookPersistenceLocation::OnDisk(path) => path.clone(),
            WorkbookPersistenceLocation::InMemory => return Ok(()),
        };

        let Some(desired_db_path) = autosave_db_path_for_workbook(saved_path) else {
            return Ok(());
        };
        if desired_db_path == current_db_path {
            return Ok(());
        }

        if let Some(parent) = desired_db_path.parent() {
            std::fs::create_dir_all(parent)
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;
        }

        // Ensure all pending in-memory edits are durably persisted before exporting.
        persistent
            .memory
            .flush_dirty_pages()
            .map_err(|e| AppStateError::Persistence(e.to_string()))?;

        let model = persistent
            .storage
            .export_model_workbook(persistent.workbook_id)
            .map_err(|e| AppStateError::Persistence(e.to_string()))?;

        let new_storage = formula_storage::Storage::open_path(&desired_db_path)
            .map_err(|e| AppStateError::Persistence(e.to_string()))?;

        let workbook_name = Path::new(saved_path)
            .file_stem()
            .and_then(|s| s.to_str())
            .filter(|s| !s.trim().is_empty())
            .unwrap_or("Workbook");

        let new_workbook_meta = new_storage
            .import_model_workbook(&model, ImportModelWorkbookOptions::new(workbook_name))
            .map_err(|e| AppStateError::Persistence(e.to_string()))?;
        let new_sheet_metas = new_storage
            .list_sheets(new_workbook_meta.id)
            .map_err(|e| AppStateError::Persistence(e.to_string()))?;

        let new_memory = open_memory_manager(new_storage.clone());
        let new_autosave = tokio::runtime::Handle::try_current()
            .ok()
            .map(|_| Arc::new(AutoSaveManager::spawn(new_memory.clone(), AutoSaveConfig::default())));

        let mut sheet_map = HashMap::<String, Uuid>::new();
        if let Some(workbook) = self.workbook.as_ref() {
            for sheet in &workbook.sheets {
                if let Some(meta) = new_sheet_metas
                    .iter()
                    .find(|m| sheet_name_eq_case_insensitive(&m.name, &sheet.name))
                {
                    sheet_map.insert(sheet.id.clone(), meta.id);
                }
            }
        }

        persistent.location = WorkbookPersistenceLocation::OnDisk(desired_db_path);
        persistent.storage = new_storage;
        persistent.memory = new_memory;
        persistent.autosave = new_autosave;
        persistent.workbook_id = new_workbook_meta.id;
        persistent.sheet_map = sheet_map;

        Ok(())
    }

    pub fn get_workbook(&self) -> Result<&Workbook, AppStateError> {
        self.workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)
    }

    pub fn get_workbook_mut(&mut self) -> Result<&mut Workbook, AppStateError> {
        self.workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)
    }

    pub fn sheet_print_settings(
        &self,
        sheet_id: &str,
    ) -> Result<SheetPrintSettings, AppStateError> {
        let workbook = self
            .workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let sheet = workbook
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;

        let settings = workbook
            .print_settings
            .sheets
            .iter()
            .find(|s| sheet_name_eq_case_insensitive(&s.sheet_name, &sheet.name));

        let mut out = settings
            .cloned()
            .unwrap_or_else(|| default_sheet_print_settings(sheet.name.clone()));
        out.sheet_name = sheet.name.clone();
        Ok(out)
    }

    pub fn set_sheet_print_area(
        &mut self,
        sheet_id: &str,
        print_area: Option<Vec<PrintCellRange>>,
    ) -> Result<(), AppStateError> {
        let workbook = self
            .workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let sheet = workbook
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
        let sheet_name = sheet.name.clone();

        let settings =
            ensure_sheet_print_settings(&mut workbook.print_settings.sheets, &sheet_name);
        settings.print_area = print_area;

        self.dirty = true;
        self.redo_stack.clear();
        Ok(())
    }

    pub fn set_sheet_page_setup(
        &mut self,
        sheet_id: &str,
        page_setup: PageSetup,
    ) -> Result<(), AppStateError> {
        let workbook = self
            .workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let sheet = workbook
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
        let sheet_name = sheet.name.clone();

        let settings =
            ensure_sheet_print_settings(&mut workbook.print_settings.sheets, &sheet_name);
        settings.page_setup = page_setup;

        self.dirty = true;
        self.redo_stack.clear();
        Ok(())
    }

    pub fn get_cell(
        &self,
        sheet_id: &str,
        row: usize,
        col: usize,
    ) -> Result<CellData, AppStateError> {
        let workbook = self
            .workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let date_system = workbook_date_system(workbook);
        let sheet = workbook
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
        let cell = sheet.get_cell(row, col);

        let formula = if let Some(persistent) = self.persistent.as_ref() {
            let sheet_uuid = persistent.sheet_uuid(sheet_id).ok_or_else(|| {
                AppStateError::Persistence(format!(
                    "missing persistence mapping for sheet id {sheet_id}"
                ))
            })?;
            let viewport = persistent
                .memory
                .load_viewport(
                    sheet_uuid,
                    StorageCellRange::new(row as i64, row as i64, col as i64, col as i64),
                )
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;
            let cached = viewport
                .get(row as i64, col as i64)
                .and_then(|c| c.formula.as_ref())
                .and_then(|f| {
                    let display = formula_model::display_formula_text(f);
                    if display.is_empty() {
                        None
                    } else {
                        Some(display)
                    }
                });
            cached.or(cell.formula.clone())
        } else {
            cell.formula.clone()
        };

        let addr = coord_to_a1(row, col);
        let value = engine_value_to_scalar(self.engine.get_cell_value(&sheet.name, &addr));
        let display_value =
            format_scalar_for_display_with_date_system(&value, cell.number_format.as_deref(), date_system);
        Ok(CellData {
            value,
            formula,
            display_value,
        })
    }

    pub fn get_range(
        &self,
        sheet_id: &str,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> Result<Vec<Vec<CellData>>, AppStateError> {
        if start_row > end_row || start_col > end_col {
            return Err(AppStateError::InvalidRange {
                start_row,
                start_col,
                end_row,
                end_col,
            });
        }

        let row_count = end_row
            .checked_sub(start_row)
            .and_then(|d| d.checked_add(1))
            .unwrap_or(usize::MAX);
        let col_count = end_col
            .checked_sub(start_col)
            .and_then(|d| d.checked_add(1))
            .unwrap_or(usize::MAX);

        if row_count > MAX_RANGE_DIM || col_count > MAX_RANGE_DIM {
            return Err(AppStateError::RangeDimensionTooLarge {
                rows: row_count,
                cols: col_count,
                limit: MAX_RANGE_DIM,
            });
        }

        let cell_count = (row_count as u128) * (col_count as u128);
        if cell_count > MAX_RANGE_CELLS_PER_CALL as u128 {
            return Err(AppStateError::RangeTooLarge {
                rows: row_count,
                cols: col_count,
                limit: MAX_RANGE_CELLS_PER_CALL,
            });
        }

        let workbook = self
            .workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let date_system = workbook_date_system(workbook);
        let sheet = workbook
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;

        // Fetch workbook cell metadata (formulas/number formats) in bulk. This is columnar-aware and
        // avoids per-cell lookups.
        let cells = sheet.get_range_cells(start_row, start_col, end_row, end_col);

        // Fetch computed engine values in bulk to avoid per-cell A1 conversions and engine lookups.
        //
        // The engine uses `u32` coordinates; preserve prior behavior for out-of-range indices by
        // filling with `#REF!` (or blank when the sheet does not exist in the engine).
        let engine_sheet_exists = self.engine.sheet_id(&sheet.name).is_some();
        let values: Vec<Vec<EngineValue>> = if !engine_sheet_exists {
            vec![vec![EngineValue::Blank; col_count]; row_count]
        } else {
            match (u32::try_from(start_row), u32::try_from(start_col)) {
                (Ok(start_row_u32), Ok(start_col_u32)) => {
                    // Clamp the representable portion of the range to `u32::MAX` and fill anything
                    // beyond with `#REF!` to match the per-cell `get_cell_value` semantics when A1
                    // parsing would overflow.
                    let max_row = u64::from(u32::MAX);
                    let max_col = u64::from(u32::MAX);
                    let remaining_rows =
                        max_row.saturating_sub(u64::from(start_row_u32)).saturating_add(1);
                    let remaining_cols =
                        max_col.saturating_sub(u64::from(start_col_u32)).saturating_add(1);
                    let in_bounds_rows = remaining_rows.min(row_count as u64) as usize;
                    let in_bounds_cols = remaining_cols.min(col_count as u64) as usize;

                    let end_row_u32 = start_row_u32 + (in_bounds_rows as u32).saturating_sub(1);
                    let end_col_u32 = start_col_u32 + (in_bounds_cols as u32).saturating_sub(1);

                    let range = formula_model::Range::new(
                        formula_model::CellRef::new(start_row_u32, start_col_u32),
                        formula_model::CellRef::new(end_row_u32, end_col_u32),
                    );

                    let mut values = self
                        .engine
                        .get_range_values(&sheet.name, range)
                        .map_err(|e| AppStateError::Engine(e.to_string()))?;

                    if in_bounds_cols < col_count {
                        for row in &mut values {
                            row.extend(
                                std::iter::repeat(EngineValue::Error(ErrorKind::Ref))
                                    .take(col_count - in_bounds_cols),
                            );
                        }
                    }
                    if in_bounds_rows < row_count {
                        values.extend(
                            std::iter::repeat(vec![EngineValue::Error(ErrorKind::Ref); col_count])
                                .take(row_count - in_bounds_rows),
                        );
                    }

                    values
                }
                _ => vec![vec![EngineValue::Error(ErrorKind::Ref); col_count]; row_count],
            }
        };

        let viewport = if let Some(persistent) = self.persistent.as_ref() {
            let sheet_uuid = persistent.sheet_uuid(sheet_id).ok_or_else(|| {
                AppStateError::Persistence(format!(
                    "missing persistence mapping for sheet id {sheet_id}"
                ))
            })?;
            let viewport = persistent
                .memory
                .load_viewport(
                    sheet_uuid,
                    StorageCellRange::new(
                        start_row as i64,
                        end_row as i64,
                        start_col as i64,
                        end_col as i64,
                    ),
                )
                .map_err(|e| AppStateError::Persistence(e.to_string()))?;
            Some(viewport)
        } else {
            None
        };

        let mut rows_out = Vec::with_capacity(row_count);
        for (r_off, (value_row, cell_row)) in values.into_iter().zip(cells.into_iter()).enumerate()
        {
            let mut row_out = Vec::with_capacity(col_count);
            for (c_off, (engine_value, cell)) in
                value_row.into_iter().zip(cell_row.into_iter()).enumerate()
            {
                let value = engine_value_to_scalar(engine_value);
                let display_value = format_scalar_for_display_with_date_system(
                    &value,
                    cell.number_format.as_deref(),
                    date_system,
                );

                let r = start_row.saturating_add(r_off);
                let c = start_col.saturating_add(c_off);

                let formula = if let Some(viewport) = viewport.as_ref() {
                    let cached = viewport
                        .get(r as i64, c as i64)
                        .and_then(|c| c.formula.as_ref())
                        .and_then(|f| {
                            let display = formula_model::display_formula_text(f);
                            if display.is_empty() {
                                None
                            } else {
                                Some(display)
                            }
                        });
                    cached.or(cell.formula)
                } else {
                    cell.formula
                };

                row_out.push(CellData {
                    value,
                    formula,
                    display_value,
                });
            }
            rows_out.push(row_out);
        }
        Ok(rows_out)
    }

    pub fn get_precedents(
        &self,
        sheet_id: &str,
        row: usize,
        col: usize,
        transitive: bool,
    ) -> Result<Vec<String>, AppStateError> {
        let workbook = self.get_workbook()?;
        let sheet = resolve_sheet_case_insensitive(workbook, sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
        let sheet_name = sheet.name.clone();

        let addr = coord_to_a1(row, col);
        let nodes = if transitive {
            self.engine
                .precedents_transitive(&sheet_name, &addr)
                .map_err(|e| AppStateError::Engine(e.to_string()))?
        } else {
            self.engine
                .precedents(&sheet_name, &addr)
                .map_err(|e| AppStateError::Engine(e.to_string()))?
        };

        let cells = expand_precedent_nodes_to_cells(&nodes, AUDITING_RESULT_LIMIT, "precedents")?;
        Ok(format_auditing_cells(workbook, &sheet_name, cells))
    }

    pub fn get_dependents(
        &self,
        sheet_id: &str,
        row: usize,
        col: usize,
        transitive: bool,
    ) -> Result<Vec<String>, AppStateError> {
        let workbook = self.get_workbook()?;
        let sheet = resolve_sheet_case_insensitive(workbook, sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
        let sheet_name = sheet.name.clone();

        let addr = coord_to_a1(row, col);
        let nodes = if transitive {
            self.engine
                .dependents_transitive(&sheet_name, &addr)
                .map_err(|e| AppStateError::Engine(e.to_string()))?
        } else {
            self.engine
                .dependents(&sheet_name, &addr)
                .map_err(|e| AppStateError::Engine(e.to_string()))?
        };

        let cells = expand_precedent_nodes_to_cells(&nodes, AUDITING_RESULT_LIMIT, "dependents")?;
        Ok(format_auditing_cells(workbook, &sheet_name, cells))
    }

    pub fn set_cell(
        &mut self,
        sheet_id: &str,
        row: usize,
        col: usize,
        value: Option<JsonValue>,
        formula: Option<String>,
    ) -> Result<Vec<CellUpdateData>, AppStateError> {
        let before = self.snapshot_cell(sheet_id, row, col)?;
        let after_cell = CellInputSnapshot {
            sheet_id: sheet_id.to_string(),
            row,
            col,
            value: value.as_ref().map(CellScalar::from_json),
            formula: normalize_formula(formula),
        };

        if before == after_cell {
            return Ok(Vec::new());
        }

        if let Some(workbook) = self.workbook.as_mut() {
            if workbook.origin_xlsx_bytes.is_some() || workbook.origin_xlsb_path.is_some() {
                workbook
                    .cell_input_baseline
                    .entry((before.sheet_id.clone(), before.row, before.col))
                    .or_insert_with(|| (before.value.clone(), before.formula.clone()));
            }
        }

        self.apply_snapshots(&[after_cell.clone()])?;
        let mut updates = self.recalculate_with_pivots(vec![(sheet_id.to_string(), row, col)])?;

        if !updates
            .iter()
            .any(|u| u.sheet_id == sheet_id && u.row == row && u.col == col)
        {
            let cell = self.get_cell(sheet_id, row, col)?;
            updates.push(CellUpdateData {
                sheet_id: sheet_id.to_string(),
                row,
                col,
                value: cell.value,
                formula: cell.formula,
                display_value: cell.display_value,
            });
        }

        self.dirty = true;
        self.redo_stack.clear();
        self.undo_stack.push(UndoEntry {
            before: vec![before],
            after: vec![after_cell],
        });

        Ok(updates)
    }

    pub fn set_range(
        &mut self,
        sheet_id: &str,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        values: Vec<Vec<(Option<JsonValue>, Option<String>)>>,
    ) -> Result<Vec<CellUpdateData>, AppStateError> {
        if start_row > end_row || start_col > end_col {
            return Err(AppStateError::InvalidRange {
                start_row,
                start_col,
                end_row,
                end_col,
            });
        }

        let row_count = end_row
            .checked_sub(start_row)
            .and_then(|d| d.checked_add(1))
            .unwrap_or(usize::MAX);
        let col_count = end_col
            .checked_sub(start_col)
            .and_then(|d| d.checked_add(1))
            .unwrap_or(usize::MAX);

        if row_count > MAX_RANGE_DIM || col_count > MAX_RANGE_DIM {
            return Err(AppStateError::RangeDimensionTooLarge {
                rows: row_count,
                cols: col_count,
                limit: MAX_RANGE_DIM,
            });
        }

        let cell_count = (row_count as u128) * (col_count as u128);
        if cell_count > MAX_RANGE_CELLS_PER_CALL as u128 {
            return Err(AppStateError::RangeTooLarge {
                rows: row_count,
                cols: col_count,
                limit: MAX_RANGE_CELLS_PER_CALL,
            });
        }

        let expected_rows = row_count;
        let expected_cols = col_count;
        if values.len() != expected_rows || values.iter().any(|row| row.len() != expected_cols) {
            return Err(AppStateError::InvalidRange {
                start_row,
                start_col,
                end_row,
                end_col,
            });
        }

        let mut before = Vec::new();
        let mut after = Vec::new();
        let mut changed = Vec::new();

        for (r_off, row_values) in values.iter().enumerate() {
            for (c_off, (value, formula)) in row_values.iter().enumerate() {
                let row = start_row + r_off;
                let col = start_col + c_off;

                let snapshot_before = self.snapshot_cell(sheet_id, row, col)?;
                let snapshot_after = CellInputSnapshot {
                    sheet_id: sheet_id.to_string(),
                    row,
                    col,
                    value: value.as_ref().map(CellScalar::from_json),
                    formula: normalize_formula(formula.clone()),
                };

                if snapshot_before != snapshot_after {
                    before.push(snapshot_before);
                    after.push(snapshot_after.clone());
                    changed.push((row, col));
                }
            }
        }

        if changed.is_empty() {
            return Ok(Vec::new());
        }

        if let Some(workbook) = self.workbook.as_mut() {
            if workbook.origin_xlsx_bytes.is_some() || workbook.origin_xlsb_path.is_some() {
                for snap in &before {
                    workbook
                        .cell_input_baseline
                        .entry((snap.sheet_id.clone(), snap.row, snap.col))
                        .or_insert_with(|| (snap.value.clone(), snap.formula.clone()));
                }
            }
        }

        self.apply_snapshots(&after)?;
        let direct_changes = changed
            .iter()
            .map(|(row, col)| (sheet_id.to_string(), *row, *col))
            .collect::<Vec<_>>();
        let mut updates = self.recalculate_with_pivots(direct_changes)?;

        for (row, col) in changed {
            if !updates
                .iter()
                .any(|u| u.sheet_id == sheet_id && u.row == row && u.col == col)
            {
                let cell_data = self.get_cell(sheet_id, row, col)?;
                updates.push(CellUpdateData {
                    sheet_id: sheet_id.to_string(),
                    row,
                    col,
                    value: cell_data.value,
                    formula: cell_data.formula,
                    display_value: cell_data.display_value,
                });
            }
        }

        self.dirty = true;
        self.redo_stack.clear();
        self.undo_stack.push(UndoEntry { before, after });
        Ok(updates)
    }

    pub fn clear_range(
        &mut self,
        sheet_id: &str,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> Result<Vec<CellUpdateData>, AppStateError> {
        if start_row > end_row || start_col > end_col {
            return Err(AppStateError::InvalidRange {
                start_row,
                start_col,
                end_row,
                end_col,
            });
        }

        let row_count = end_row
            .checked_sub(start_row)
            .and_then(|d| d.checked_add(1))
            .unwrap_or(usize::MAX);
        let col_count = end_col
            .checked_sub(start_col)
            .and_then(|d| d.checked_add(1))
            .unwrap_or(usize::MAX);

        if row_count > MAX_RANGE_DIM || col_count > MAX_RANGE_DIM {
            return Err(AppStateError::RangeDimensionTooLarge {
                rows: row_count,
                cols: col_count,
                limit: MAX_RANGE_DIM,
            });
        }

        let cell_count = (row_count as u128) * (col_count as u128);
        if cell_count > MAX_RANGE_CELLS_PER_CALL as u128 {
            return Err(AppStateError::RangeTooLarge {
                rows: row_count,
                cols: col_count,
                limit: MAX_RANGE_CELLS_PER_CALL,
            });
        }

        let coords_to_clear = {
            let workbook = self.get_workbook()?;
            let sheet = workbook
                .sheet(sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;

            sheet
                .cells_iter()
                .filter_map(|((row, col), cell)| {
                    if row < start_row
                        || row > end_row
                        || col < start_col
                        || col > end_col
                        || (cell.input_value.is_none() && cell.formula.is_none())
                    {
                        None
                    } else {
                        Some((row, col))
                    }
                })
                .collect::<Vec<_>>()
        };

        if coords_to_clear.is_empty() {
            return Ok(Vec::new());
        }

        let mut before = Vec::new();
        let mut after = Vec::new();
        let mut changed = Vec::new();

        for (row, col) in coords_to_clear {
            let snapshot_before = self.snapshot_cell(sheet_id, row, col)?;
            let snapshot_after = CellInputSnapshot {
                sheet_id: sheet_id.to_string(),
                row,
                col,
                value: None,
                formula: None,
            };

            if snapshot_before != snapshot_after {
                before.push(snapshot_before);
                after.push(snapshot_after);
                changed.push((row, col));
            }
        }

        if changed.is_empty() {
            return Ok(Vec::new());
        }

        if let Some(workbook) = self.workbook.as_mut() {
            if workbook.origin_xlsx_bytes.is_some() || workbook.origin_xlsb_path.is_some() {
                for snap in &before {
                    workbook
                        .cell_input_baseline
                        .entry((snap.sheet_id.clone(), snap.row, snap.col))
                        .or_insert_with(|| (snap.value.clone(), snap.formula.clone()));
                }
            }
        }

        self.apply_snapshots(&after)?;
        let direct_changes = changed
            .iter()
            .map(|(row, col)| (sheet_id.to_string(), *row, *col))
            .collect::<Vec<_>>();
        let mut updates = self.recalculate_with_pivots(direct_changes)?;

        for (row, col) in changed {
            if !updates
                .iter()
                .any(|u| u.sheet_id == sheet_id && u.row == row && u.col == col)
            {
                let cell_data = self.get_cell(sheet_id, row, col)?;
                updates.push(CellUpdateData {
                    sheet_id: sheet_id.to_string(),
                    row,
                    col,
                    value: cell_data.value,
                    formula: cell_data.formula,
                    display_value: cell_data.display_value,
                });
            }
        }

        self.dirty = true;
        self.redo_stack.clear();
        self.undo_stack.push(UndoEntry { before, after });

        Ok(updates)
    }

    pub fn list_pivot_tables(&self) -> Vec<PivotTableSummary> {
        self.pivots
            .pivots
            .iter()
            .map(|pivot| PivotTableSummary {
                id: pivot.id.clone(),
                name: pivot.name.clone(),
                source_sheet_id: pivot.source_sheet_id.clone(),
                source_range: pivot.source_range.clone(),
                destination: pivot.destination.clone(),
            })
            .collect()
    }

    pub fn create_pivot_table(
        &mut self,
        name: String,
        source_sheet_id: String,
        source_range: CellRect,
        destination: PivotDestination,
        config: PivotConfig,
    ) -> Result<(String, Vec<CellUpdateData>), AppStateError> {
        if source_range.start_row > source_range.end_row
            || source_range.start_col > source_range.end_col
        {
            return Err(AppStateError::InvalidRange {
                start_row: source_range.start_row,
                start_col: source_range.start_col,
                end_row: source_range.end_row,
                end_col: source_range.end_col,
            });
        }

        if self.workbook.is_none() {
            return Err(AppStateError::NoWorkbookLoaded);
        }

        // Validate sheet ids early so we don't register unusable pivots.
        let workbook = self.workbook.as_ref().expect("checked is_some above");
        if workbook.sheet(&source_sheet_id).is_none() {
            return Err(AppStateError::UnknownSheet(source_sheet_id));
        }
        if workbook.sheet(&destination.sheet_id).is_none() {
            return Err(AppStateError::UnknownSheet(destination.sheet_id));
        }

        let pivot_id = PivotManager::next_id();
        self.pivots.pivots.push(PivotRegistration {
            id: pivot_id.clone(),
            name,
            source_sheet_id,
            source_range,
            destination,
            config,
            last_output_range: None,
        });

        let updates = match self.refresh_pivot_table(&pivot_id) {
            Ok(updates) => updates,
            Err(err) => {
                // Best-effort rollback if refresh failed.
                self.pivots.pivots.retain(|p| p.id != pivot_id);
                return Err(err);
            }
        };

        self.dirty = true;
        self.redo_stack.clear();

        Ok((pivot_id, updates))
    }

    pub fn refresh_pivot_table(
        &mut self,
        pivot_id: &str,
    ) -> Result<Vec<CellUpdateData>, AppStateError> {
        let idx = self
            .pivots
            .pivots
            .iter()
            .position(|p| p.id == pivot_id)
            .ok_or_else(|| AppStateError::UnknownPivot(pivot_id.to_string()))?;

        let workbook = self
            .workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let engine = &mut self.engine;

        let pivot_updates =
            refresh_pivot_registration(workbook, engine, &mut self.pivots.pivots[idx])?;

        if !pivot_updates.is_empty() {
            self.dirty = true;
        }

        // Recalculate once more so formulas depending on the pivot output update.
        let recalc_changes = self.engine.recalculate_with_value_changes_multi_threaded();
        let mut updates = pivot_updates;
        updates.extend(self.refresh_computed_values_from_recalc_changes(&recalc_changes)?);

        Ok(dedupe_updates(updates))
    }

    pub fn refresh_all_pivots(&mut self) -> Result<Vec<CellUpdateData>, AppStateError> {
        if self.pivots.pivots.is_empty() {
            return Ok(Vec::new());
        }

        // Refresh pivots in workbook sheet order, then by top-left destination cell.
        // This mirrors how users think about pivots ("in sheet order") and ensures a stable refresh
        // sequence even when pivot ids are generated out-of-order.
        let pivot_updates = {
            let workbook = self
                .workbook
                .as_mut()
                .ok_or(AppStateError::NoWorkbookLoaded)?;
            let engine = &mut self.engine;

            let mut sheet_order: HashMap<&str, usize> = HashMap::new();
            for (idx, sheet) in workbook.sheets.iter().enumerate() {
                sheet_order.insert(sheet.id.as_str(), idx);
            }

            let mut indices: Vec<usize> = (0..self.pivots.pivots.len()).collect();
            indices.sort_by_key(|idx| {
                let pivot = &self.pivots.pivots[*idx];
                (
                    sheet_order
                        .get(pivot.destination.sheet_id.as_str())
                        .copied()
                        .unwrap_or(usize::MAX),
                    pivot.destination.row,
                    pivot.destination.col,
                    *idx,
                )
            });

            let mut pivot_updates = Vec::new();
            for idx in indices {
                pivot_updates.extend(refresh_pivot_registration(
                    workbook,
                    engine,
                    &mut self.pivots.pivots[idx],
                )?);
            }
            pivot_updates
        };

        if !pivot_updates.is_empty() {
            self.dirty = true;
        }

        // Recalculate once so formulas depending on any pivot outputs update.
        let recalc_changes = self.engine.recalculate_with_value_changes_multi_threaded();
        let mut updates = pivot_updates;
        updates.extend(self.refresh_computed_values_from_recalc_changes(&recalc_changes)?);
        Ok(dedupe_updates(updates))
    }

    pub fn recalculate_all(&mut self) -> Result<Vec<CellUpdateData>, AppStateError> {
        if self.workbook.is_none() {
            return Err(AppStateError::NoWorkbookLoaded);
        }
        // `formula_engine::Engine` only recalculates the current dirty set (plus volatile
        // closure). The desktop `recalculate()` command is intended as a user-triggered
        // "recalculate all formulas" action, so we re-feed existing formulas to mark them
        // dirty before recalculating.
        let formulas = {
            let workbook = self
                .workbook
                .as_ref()
                .expect("checked workbook is_some above");
            workbook
                .sheets
                .iter()
                .flat_map(|sheet| {
                    let sheet_name = sheet.name.clone();
                    sheet.cells_iter().filter_map(move |((row, col), cell)| {
                        cell.formula.as_ref().map(|formula| {
                            (sheet_name.clone(), coord_to_a1(row, col), formula.clone())
                        })
                    })
                })
                .collect::<Vec<_>>()
        };

        for (sheet_name, addr, formula) in formulas {
            if self
                .engine
                .set_cell_formula(&sheet_name, &addr, &formula)
                .is_err()
            {
                let _ = self.engine.set_cell_value(
                    &sheet_name,
                    &addr,
                    EngineValue::Error(ErrorKind::Name),
                );
            }
        }

        let recalc_changes = self.engine.recalculate_with_value_changes_multi_threaded();
        self.refresh_computed_values_from_recalc_changes(&recalc_changes)
    }

    pub fn goal_seek(
        &mut self,
        sheet_id: &str,
        params: GoalSeekParams,
    ) -> Result<(GoalSeekResult, Vec<CellUpdateData>), AppStateError> {
        let default_sheet_name = self
            .get_workbook()?
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?
            .name
            .clone();

        let (changing_sheet_id, changing_row, changing_col) = resolve_cell_ref(
            self.get_workbook()?,
            sheet_id,
            &default_sheet_name,
            &params.changing_cell,
        )?;

        let mut model = EngineWhatIfModel::new(&mut self.engine, default_sheet_name)
            .with_recalc_mode(RecalcMode::SingleThreaded);

        let result = GoalSeek::solve(&mut model, params)
            .map_err(|e| AppStateError::WhatIf(e.to_string()))?;

        let updates = self.set_cell(
            &changing_sheet_id,
            changing_row,
            changing_col,
            Some(JsonValue::from(result.solution)),
            None,
        )?;

        Ok((result, updates))
    }

    pub fn run_monte_carlo(
        &mut self,
        sheet_id: &str,
        config: SimulationConfig,
    ) -> Result<SimulationResult, AppStateError> {
        let default_sheet_name = self
            .get_workbook()?
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?
            .name
            .clone();

        let mut model = EngineWhatIfModel::new(&mut self.engine, default_sheet_name)
            .with_recalc_mode(RecalcMode::SingleThreaded);

        // Keep the engine consistent with the workbook after the simulation by
        // restoring the original input values (the workbook state is not mutated
        // during the simulation run).
        let mut base_inputs = std::collections::HashMap::<WhatIfCellRef, WhatIfCellValue>::new();
        for input in &config.input_distributions {
            if base_inputs.contains_key(&input.cell) {
                continue;
            }
            let base = model
                .get_cell_value(&input.cell)
                .map_err(|e| AppStateError::WhatIf(e.to_string()))?;
            base_inputs.insert(input.cell.clone(), base);
        }

        let result = MonteCarloEngine::run_simulation(&mut model, config);

        // Best-effort restore, even if the simulation fails.
        for (cell, value) in base_inputs {
            let _ = model.set_cell_value(&cell, value);
        }
        let _ = model.recalculate();

        result.map_err(|e| AppStateError::WhatIf(e.to_string()))
    }

    pub fn list_scenarios(&self) -> Vec<Scenario> {
        self.scenario_manager.scenarios().cloned().collect()
    }

    pub fn create_scenario(
        &mut self,
        sheet_id: &str,
        name: String,
        changing_cells: Vec<WhatIfCellRef>,
        created_by: String,
        comment: Option<String>,
    ) -> Result<Scenario, AppStateError> {
        let workbook = self.get_workbook()?;
        let default_sheet = workbook
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;

        let mut cells = Vec::with_capacity(changing_cells.len());
        let mut values = Vec::with_capacity(changing_cells.len());
        for cell in changing_cells {
            let (resolved_sheet_id, row, col) =
                resolve_cell_ref(workbook, sheet_id, &default_sheet.name, &cell)?;
            if workbook.cell_has_formula(&resolved_sheet_id, row, col) {
                return Err(AppStateError::WhatIf(format!(
                    "cannot create scenario: {cell} contains a formula"
                )));
            }

            let scalar = workbook.cell_value(&resolved_sheet_id, row, col);
            let sheet_name = workbook
                .sheet(&resolved_sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(resolved_sheet_id.clone()))?
                .name
                .clone();
            let canonical = WhatIfCellRef::new(format!(
                "{}!{}",
                quote_sheet_name(&sheet_name),
                coord_to_a1(row, col)
            ));

            cells.push(canonical);
            values.push(scalar_to_what_if_value(&scalar));
        }

        let id = self
            .scenario_manager
            .create_scenario(name, cells, values, created_by, comment)
            .map_err(|e| AppStateError::WhatIf(e.to_string()))?;

        self.scenario_manager
            .get(id)
            .cloned()
            .ok_or_else(|| AppStateError::WhatIf("scenario not found after creation".to_string()))
    }

    pub fn apply_scenario(
        &mut self,
        sheet_id: &str,
        scenario_id: ScenarioId,
    ) -> Result<Vec<CellUpdateData>, AppStateError> {
        if self.workbook.is_none() {
            return Err(AppStateError::NoWorkbookLoaded);
        }

        let default_sheet_name = self
            .get_workbook()?
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?
            .name
            .clone();

        // First, update the scenario manager bookkeeping and engine values.
        {
            let mut model = EngineWhatIfModel::new(&mut self.engine, default_sheet_name.clone())
                .with_recalc_mode(RecalcMode::SingleThreaded);
            self.scenario_manager
                .apply_scenario(&mut model, scenario_id)
                .map_err(|e| AppStateError::WhatIf(e.to_string()))?;
        }

        let scenario = self
            .scenario_manager
            .get(scenario_id)
            .cloned()
            .ok_or_else(|| AppStateError::WhatIf("scenario not found".to_string()))?;

        let workbook = self.get_workbook()?;
        let mut before = Vec::new();
        let mut after = Vec::new();
        let mut touched = Vec::new();

        for (cell, value) in scenario.values {
            let (resolved_sheet_id, row, col) =
                resolve_cell_ref(workbook, sheet_id, &default_sheet_name, &cell)?;

            before.push(self.snapshot_cell(&resolved_sheet_id, row, col)?);
            after.push(CellInputSnapshot {
                sheet_id: resolved_sheet_id.clone(),
                row,
                col,
                value: what_if_value_to_scalar(&value),
                formula: None,
            });
            touched.push((resolved_sheet_id, row, col));
        }

        if after.is_empty() {
            return Ok(Vec::new());
        }

        if let Some(workbook) = self.workbook.as_mut() {
            if workbook.origin_xlsx_bytes.is_some() || workbook.origin_xlsb_path.is_some() {
                for snap in &before {
                    workbook
                        .cell_input_baseline
                        .entry((snap.sheet_id.clone(), snap.row, snap.col))
                        .or_insert_with(|| (snap.value.clone(), snap.formula.clone()));
                }
            }
        }

        self.apply_snapshots(&after)?;
        let recalc_changes = self.engine.recalculate_with_value_changes_multi_threaded();
        let mut updates = self.refresh_computed_values_from_recalc_changes(&recalc_changes)?;

        for (resolved_sheet_id, row, col) in touched {
            if !updates
                .iter()
                .any(|u| u.sheet_id == resolved_sheet_id && u.row == row && u.col == col)
            {
                let cell = self.get_cell(&resolved_sheet_id, row, col)?;
                updates.push(CellUpdateData {
                    sheet_id: resolved_sheet_id,
                    row,
                    col,
                    value: cell.value,
                    formula: cell.formula,
                    display_value: cell.display_value,
                });
            }
        }

        self.dirty = true;
        self.redo_stack.clear();
        self.undo_stack.push(UndoEntry { before, after });

        Ok(updates)
    }

    pub fn restore_base_scenario(
        &mut self,
        sheet_id: &str,
    ) -> Result<Vec<CellUpdateData>, AppStateError> {
        if self.workbook.is_none() {
            return Err(AppStateError::NoWorkbookLoaded);
        }

        let default_sheet_name = self
            .get_workbook()?
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?
            .name
            .clone();

        // First, restore the engine values and scenario manager bookkeeping.
        {
            let mut model = EngineWhatIfModel::new(&mut self.engine, default_sheet_name.clone())
                .with_recalc_mode(RecalcMode::SingleThreaded);
            self.scenario_manager
                .restore_base(&mut model)
                .map_err(|e| AppStateError::WhatIf(e.to_string()))?;
        }

        let base_values = self.scenario_manager.base_values().clone();
        if base_values.is_empty() {
            return Ok(Vec::new());
        }

        let workbook = self.get_workbook()?;
        let mut before = Vec::new();
        let mut after = Vec::new();
        let mut touched = Vec::new();

        for (cell, value) in base_values {
            let (resolved_sheet_id, row, col) =
                resolve_cell_ref(workbook, sheet_id, &default_sheet_name, &cell)?;

            before.push(self.snapshot_cell(&resolved_sheet_id, row, col)?);
            after.push(CellInputSnapshot {
                sheet_id: resolved_sheet_id.clone(),
                row,
                col,
                value: what_if_value_to_scalar(&value),
                formula: None,
            });
            touched.push((resolved_sheet_id, row, col));
        }

        if let Some(workbook) = self.workbook.as_mut() {
            if workbook.origin_xlsx_bytes.is_some() || workbook.origin_xlsb_path.is_some() {
                for snap in &before {
                    workbook
                        .cell_input_baseline
                        .entry((snap.sheet_id.clone(), snap.row, snap.col))
                        .or_insert_with(|| (snap.value.clone(), snap.formula.clone()));
                }
            }
        }

        self.apply_snapshots(&after)?;
        let recalc_changes = self.engine.recalculate_with_value_changes_multi_threaded();
        let mut updates = self.refresh_computed_values_from_recalc_changes(&recalc_changes)?;

        for (resolved_sheet_id, row, col) in touched {
            if !updates
                .iter()
                .any(|u| u.sheet_id == resolved_sheet_id && u.row == row && u.col == col)
            {
                let cell = self.get_cell(&resolved_sheet_id, row, col)?;
                updates.push(CellUpdateData {
                    sheet_id: resolved_sheet_id,
                    row,
                    col,
                    value: cell.value,
                    formula: cell.formula,
                    display_value: cell.display_value,
                });
            }
        }

        self.dirty = true;
        self.redo_stack.clear();
        self.undo_stack.push(UndoEntry { before, after });

        Ok(updates)
    }

    pub fn generate_summary_report(
        &mut self,
        sheet_id: &str,
        result_cells: Vec<WhatIfCellRef>,
        scenario_ids: Vec<ScenarioId>,
    ) -> Result<SummaryReport, AppStateError> {
        if self.workbook.is_none() {
            return Err(AppStateError::NoWorkbookLoaded);
        }

        let default_sheet_name = self
            .get_workbook()?
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?
            .name
            .clone();

        // Snapshot current engine inputs for all scenario-changing cells so we
        // can restore the engine state after report generation.
        let mut cells_to_restore = std::collections::HashSet::<WhatIfCellRef>::new();
        for id in &scenario_ids {
            if let Some(s) = self.scenario_manager.get(*id) {
                for cell in &s.changing_cells {
                    cells_to_restore.insert(cell.clone());
                }
            }
        }

        let mut model = EngineWhatIfModel::new(&mut self.engine, default_sheet_name.clone())
            .with_recalc_mode(RecalcMode::SingleThreaded);

        let mut saved = std::collections::HashMap::<WhatIfCellRef, WhatIfCellValue>::new();
        for cell in &cells_to_restore {
            let value = model
                .get_cell_value(cell)
                .map_err(|e| AppStateError::WhatIf(e.to_string()))?;
            saved.insert(cell.clone(), value);
        }

        let mut manager = self.scenario_manager.clone();
        let report = manager
            .generate_summary_report(&mut model, result_cells, scenario_ids)
            .map_err(|e| AppStateError::WhatIf(e.to_string()));

        // Always restore, even if report generation fails.
        for (cell, value) in saved {
            let _ = model.set_cell_value(&cell, value);
        }
        let _ = model.recalculate();

        report
    }

    pub fn undo(&mut self) -> Result<Vec<CellUpdateData>, AppStateError> {
        if self.workbook.is_none() {
            return Err(AppStateError::NoWorkbookLoaded);
        }
        let entry = self.undo_stack.pop().ok_or(AppStateError::NoUndoHistory)?;
        let direct_changes = entry
            .before
            .iter()
            .map(|cell| (cell.sheet_id.clone(), cell.row, cell.col))
            .collect::<Vec<_>>();
        self.apply_snapshots(&entry.before)?;
        let mut updates = self.recalculate_with_pivots(direct_changes)?;

        for cell in &entry.before {
            if !updates
                .iter()
                .any(|u| u.sheet_id == cell.sheet_id && u.row == cell.row && u.col == cell.col)
            {
                let cell_data = self.get_cell(&cell.sheet_id, cell.row, cell.col)?;
                updates.push(CellUpdateData {
                    sheet_id: cell.sheet_id.clone(),
                    row: cell.row,
                    col: cell.col,
                    value: cell_data.value,
                    formula: cell_data.formula,
                    display_value: cell_data.display_value,
                });
            }
        }

        self.redo_stack.push(entry);
        self.dirty = true;
        Ok(updates)
    }

    pub fn redo(&mut self) -> Result<Vec<CellUpdateData>, AppStateError> {
        if self.workbook.is_none() {
            return Err(AppStateError::NoWorkbookLoaded);
        }
        let entry = self.redo_stack.pop().ok_or(AppStateError::NoRedoHistory)?;
        let direct_changes = entry
            .after
            .iter()
            .map(|cell| (cell.sheet_id.clone(), cell.row, cell.col))
            .collect::<Vec<_>>();
        self.apply_snapshots(&entry.after)?;
        let mut updates = self.recalculate_with_pivots(direct_changes)?;

        for cell in &entry.after {
            if !updates
                .iter()
                .any(|u| u.sheet_id == cell.sheet_id && u.row == cell.row && u.col == cell.col)
            {
                let cell_data = self.get_cell(&cell.sheet_id, cell.row, cell.col)?;
                updates.push(CellUpdateData {
                    sheet_id: cell.sheet_id.clone(),
                    row: cell.row,
                    col: cell.col,
                    value: cell_data.value,
                    formula: cell_data.formula,
                    display_value: cell_data.display_value,
                });
            }
        }

        self.undo_stack.push(entry);
        self.dirty = true;
        Ok(updates)
    }

    pub fn vba_project(&mut self) -> Result<Option<formula_vba::VBAProject>, MacroHostError> {
        let workbook = self
            .workbook
            .as_ref()
            .ok_or(MacroHostError::NoWorkbookLoaded)?;
        self.macro_host.project(workbook)
    }

    pub fn list_macros(&mut self) -> Result<Vec<MacroInfo>, MacroHostError> {
        let workbook = self
            .workbook
            .as_ref()
            .ok_or(MacroHostError::NoWorkbookLoaded)?;
        self.macro_host.list_macros(workbook)
    }

    pub fn macro_runtime_context(&self) -> MacroRuntimeContext {
        self.macro_host.runtime_context()
    }

    pub fn set_macro_runtime_context(&mut self, ctx: MacroRuntimeContext) -> Result<(), MacroHostError> {
        let workbook = self
            .workbook
            .as_ref()
            .ok_or(MacroHostError::NoWorkbookLoaded)?;

        if ctx.active_sheet >= workbook.sheets.len() {
            return Err(MacroHostError::Runtime(format!(
                "sheet index out of range: {}",
                ctx.active_sheet
            )));
        }

        if ctx.active_cell.0 == 0 || ctx.active_cell.1 == 0 {
            return Err(MacroHostError::Runtime("ActiveCell is 1-based".to_string()));
        }

        if let Some(sel) = ctx.selection {
            if sel.sheet >= workbook.sheets.len() {
                return Err(MacroHostError::Runtime(format!(
                    "selection sheet index out of range: {}",
                    sel.sheet
                )));
            }
            if sel.start_row == 0 || sel.start_col == 0 || sel.end_row == 0 || sel.end_col == 0 {
                return Err(MacroHostError::Runtime("Selection is 1-based".to_string()));
            }
            if sel.start_row > sel.end_row || sel.start_col > sel.end_col {
                return Err(MacroHostError::Runtime(format!(
                    "invalid selection range: start ({},{}) end ({},{})",
                    sel.start_row, sel.start_col, sel.end_row, sel.end_col
                )));
            }
        }

        // Ensure the macro host has already synchronized with the workbook so the context
        // is not reset the first time the VBA project hash is computed.
        self.macro_host.sync_with_workbook(workbook);
        self.macro_host.set_runtime_context(ctx);
        Ok(())
    }

    pub fn set_macro_ui_context(
        &mut self,
        sheet_id: &str,
        active_row: usize,
        active_col: usize,
        selection: Option<CellRect>,
    ) -> Result<(), MacroHostError> {
        let workbook = self
            .workbook
            .as_ref()
            .ok_or(MacroHostError::NoWorkbookLoaded)?;
        let sheet_index = workbook
            .sheets
            .iter()
            .position(|s| s.id == sheet_id)
            .ok_or_else(|| MacroHostError::Runtime(format!("unknown sheet id: {sheet_id}")))?;

        self.macro_host.sync_with_workbook(workbook);

        let row = u32::try_from(active_row.saturating_add(1))
            .map_err(|_| MacroHostError::Runtime("row index out of range".to_string()))?;
        let col = u32::try_from(active_col.saturating_add(1))
            .map_err(|_| MacroHostError::Runtime("col index out of range".to_string()))?;

        let selection = match selection {
            Some(rect) => {
                if rect.start_row > rect.end_row || rect.start_col > rect.end_col {
                    return Err(MacroHostError::Runtime(format!(
                        "invalid range: start ({},{}) end ({},{})",
                        rect.start_row, rect.start_col, rect.end_row, rect.end_col
                    )));
                }
                let start_row = u32::try_from(rect.start_row.saturating_add(1))
                    .map_err(|_| MacroHostError::Runtime("row index out of range".to_string()))?;
                let start_col = u32::try_from(rect.start_col.saturating_add(1))
                    .map_err(|_| MacroHostError::Runtime("col index out of range".to_string()))?;
                let end_row = u32::try_from(rect.end_row.saturating_add(1))
                    .map_err(|_| MacroHostError::Runtime("row index out of range".to_string()))?;
                let end_col = u32::try_from(rect.end_col.saturating_add(1))
                    .map_err(|_| MacroHostError::Runtime("col index out of range".to_string()))?;
                Some(formula_vba_runtime::VbaRangeRef {
                    sheet: sheet_index,
                    start_row,
                    start_col,
                    end_row,
                    end_col,
                })
            }
            None => None,
        };

        self.macro_host.set_runtime_context(MacroRuntimeContext {
            active_sheet: sheet_index,
            active_cell: (row, col),
            selection,
        });

        Ok(())
    }

    pub fn run_macro(
        &mut self,
        macro_id: &str,
        options: MacroExecutionOptions,
    ) -> Result<MacroExecutionOutcome, MacroHostError> {
        let (program, ctx, origin_path) = {
            let workbook = self
                .workbook
                .as_ref()
                .ok_or(MacroHostError::NoWorkbookLoaded)?;
            let program = self
                .macro_host
                .program(workbook)?
                .ok_or_else(|| MacroHostError::Runtime("no VBA macros in workbook".to_string()))?;
            let ctx = self.macro_host.runtime_context();
            let origin_path = workbook.origin_path.clone();
            (program, ctx, origin_path)
        };

        let (result, new_ctx) = execute_invocation(
            self,
            program,
            ctx,
            origin_path,
            MacroInvocation::Procedure {
                macro_id: macro_id.to_string(),
            },
            options,
        )?;
        self.macro_host.set_runtime_context(new_ctx);
        Ok(result)
    }

    pub fn fire_workbook_open(
        &mut self,
        options: MacroExecutionOptions,
    ) -> Result<MacroExecutionOutcome, MacroHostError> {
        self.fire_macro_event(MacroInvocation::WorkbookOpen, options)
    }

    pub fn fire_workbook_before_close(
        &mut self,
        options: MacroExecutionOptions,
    ) -> Result<MacroExecutionOutcome, MacroHostError> {
        self.fire_macro_event(MacroInvocation::WorkbookBeforeClose, options)
    }

    pub fn fire_worksheet_change(
        &mut self,
        sheet_id: &str,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        options: MacroExecutionOptions,
    ) -> Result<MacroExecutionOutcome, MacroHostError> {
        let target = self.range_for_event(sheet_id, start_row, start_col, end_row, end_col)?;
        self.fire_macro_event(MacroInvocation::WorksheetChange { target }, options)
    }

    pub fn fire_selection_change(
        &mut self,
        sheet_id: &str,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
        options: MacroExecutionOptions,
    ) -> Result<MacroExecutionOutcome, MacroHostError> {
        let target = self.range_for_event(sheet_id, start_row, start_col, end_row, end_col)?;
        self.fire_macro_event(MacroInvocation::SelectionChange { target }, options)
    }

    fn fire_macro_event(
        &mut self,
        invocation: MacroInvocation,
        options: MacroExecutionOptions,
    ) -> Result<MacroExecutionOutcome, MacroHostError> {
        let (program, ctx, origin_path) = {
            let workbook = self
                .workbook
                .as_ref()
                .ok_or(MacroHostError::NoWorkbookLoaded)?;
            let program = self.macro_host.program(workbook)?;
            let ctx = self.macro_host.runtime_context();
            let origin_path = workbook.origin_path.clone();
            (program, ctx, origin_path)
        };

        let Some(program) = program else {
            return Ok(MacroExecutionOutcome {
                ok: true,
                output: Vec::new(),
                updates: Vec::new(),
                error: None,
                permission_request: None,
            });
        };

        let (result, new_ctx) =
            execute_invocation(self, program, ctx, origin_path, invocation, options)?;
        self.macro_host.set_runtime_context(new_ctx);
        Ok(result)
    }

    fn range_for_event(
        &self,
        sheet_id: &str,
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    ) -> Result<formula_vba_runtime::VbaRangeRef, MacroHostError> {
        if start_row > end_row || start_col > end_col {
            return Err(MacroHostError::Runtime(format!(
                "invalid range: start ({start_row},{start_col}) end ({end_row},{end_col})"
            )));
        }
        let workbook = self
            .workbook
            .as_ref()
            .ok_or(MacroHostError::NoWorkbookLoaded)?;
        let sheet_index = workbook
            .sheets
            .iter()
            .position(|s| s.id == sheet_id)
            .ok_or_else(|| MacroHostError::Runtime(format!("unknown sheet id: {sheet_id}")))?;

        let start_row = u32::try_from(start_row.saturating_add(1))
            .map_err(|_| MacroHostError::Runtime("row index out of range".to_string()))?;
        let start_col = u32::try_from(start_col.saturating_add(1))
            .map_err(|_| MacroHostError::Runtime("col index out of range".to_string()))?;
        let end_row = u32::try_from(end_row.saturating_add(1))
            .map_err(|_| MacroHostError::Runtime("row index out of range".to_string()))?;
        let end_col = u32::try_from(end_col.saturating_add(1))
            .map_err(|_| MacroHostError::Runtime("col index out of range".to_string()))?;

        Ok(formula_vba_runtime::VbaRangeRef {
            sheet: sheet_index,
            start_row,
            start_col,
            end_row,
            end_col,
        })
    }

    fn recalculate_with_pivots(
        &mut self,
        mut pending_changes: Vec<(String, usize, usize)>,
    ) -> Result<Vec<CellUpdateData>, AppStateError> {
        const MAX_PIVOT_REFRESH_PASSES: usize = 2;

        let mut updates = Vec::new();
        let mut pivot_updates_in_last_pass = Vec::new();

        for pass in 0..MAX_PIVOT_REFRESH_PASSES {
            let recalc_changes = self.engine.recalculate_with_value_changes_multi_threaded();
            let formula_updates =
                self.refresh_computed_values_from_recalc_changes(&recalc_changes)?;

            let mut changed_for_pivots = pending_changes.clone();
            changed_for_pivots.extend(
                formula_updates
                    .iter()
                    .map(|u| (u.sheet_id.clone(), u.row, u.col)),
            );

            let pivot_updates = self.refresh_pivots_for_changed_cells(&changed_for_pivots)?;

            updates.extend(formula_updates);
            updates.extend(pivot_updates.clone());

            if pivot_updates.is_empty() {
                return Ok(dedupe_updates(updates));
            }

            pending_changes = pivot_updates
                .iter()
                .map(|u| (u.sheet_id.clone(), u.row, u.col))
                .collect();

            if pass == MAX_PIVOT_REFRESH_PASSES - 1 {
                pivot_updates_in_last_pass = pivot_updates;
            }
        }

        // If we hit the refresh limit, ensure we still run one final formula pass
        // so dependents of the last pivot output update.
        if !pivot_updates_in_last_pass.is_empty() {
            let recalc_changes = self.engine.recalculate_with_value_changes_multi_threaded();
            updates.extend(self.refresh_computed_values_from_recalc_changes(&recalc_changes)?);
        }

        Ok(dedupe_updates(updates))
    }

    fn refresh_pivots_for_changed_cells(
        &mut self,
        changed_cells: &[(String, usize, usize)],
    ) -> Result<Vec<CellUpdateData>, AppStateError> {
        if changed_cells.is_empty() || self.pivots.pivots.is_empty() {
            return Ok(Vec::new());
        }

        let workbook = self
            .workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let engine = &mut self.engine;

        let mut updates = Vec::new();
        for pivot in self.pivots.pivots.iter_mut() {
            let should_refresh = changed_cells.iter().any(|(sheet_id, row, col)| {
                sheet_id == &pivot.source_sheet_id && pivot.source_range.contains(*row, *col)
            });
            if !should_refresh {
                continue;
            }

            updates.extend(refresh_pivot_registration(workbook, engine, pivot)?);
        }

        if !updates.is_empty() {
            self.dirty = true;
        }

        Ok(updates)
    }

    fn snapshot_cell(
        &self,
        sheet_id: &str,
        row: usize,
        col: usize,
    ) -> Result<CellInputSnapshot, AppStateError> {
        let workbook = self
            .workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let sheet = workbook
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
        // Only snapshot the sparse overlay. Columnar-backed sheets can materialize
        // a temporary `Cell` for read access; treating those as stored values would
        // make undo/redo persist a full copy of the columnar data in the overlay.
        let (value, formula) = match sheet.cells.get(&(row, col)) {
            Some(cell) => (cell.input_value.clone(), cell.formula.clone()),
            None => (None, None),
        };

        Ok(CellInputSnapshot {
            sheet_id: sheet_id.to_string(),
            row,
            col,
            value,
            formula,
        })
    }

    fn apply_snapshots(&mut self, snapshots: &[CellInputSnapshot]) -> Result<(), AppStateError> {
        let workbook = self
            .workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let engine = &mut self.engine;

        for snap in snapshots {
            let sheet = workbook
                .sheet_mut(&snap.sheet_id)
                .ok_or_else(|| AppStateError::UnknownSheet(snap.sheet_id.clone()))?;
            let sheet_name = sheet.name.clone();
            let existing_format = sheet
                .cells
                .get(&(snap.row, snap.col))
                .and_then(|c| c.number_format.clone());
            let existing_computed_value = sheet
                .cells
                .get(&(snap.row, snap.col))
                .map(|c| c.computed_value.clone())
                .unwrap_or(CellScalar::Empty);

            let mut new_cell = match (&snap.formula, &snap.value) {
                (Some(formula), _) => {
                    let mut cell = Cell::from_formula(formula.clone());
                    // Preserve the previous computed value so caches remain stable until the
                    // next recalc pass updates the formula result.
                    //
                    // This keeps the workbook-side computed-value cache aligned with the
                    // engine's pre-recalc values, which is required for delta-based recalc
                    // updates.
                    cell.computed_value = existing_computed_value;
                    cell
                }
                (None, Some(value)) => Cell::from_literal(Some(value.clone())),
                (None, None) => Cell::empty(),
            };
            if new_cell.number_format.is_none() {
                new_cell.number_format = existing_format;
            }

            sheet.set_cell(snap.row, snap.col, new_cell);
            apply_snapshot_to_engine(
                engine,
                &sheet_name,
                snap.row,
                snap.col,
                &snap.value,
                &snap.formula,
            );

            if let Some(persistent) = self.persistent.as_ref() {
                let sheet_uuid = persistent.sheet_uuid(&snap.sheet_id).ok_or_else(|| {
                    AppStateError::Persistence(format!(
                        "missing persistence mapping for sheet id {}",
                        snap.sheet_id
                    ))
                })?;

                let value = snap
                    .value
                    .as_ref()
                    .map(scalar_to_storage_value)
                    .unwrap_or(formula_model::CellValue::Empty);
                 let formula = snap
                     .formula
                     .as_deref()
                     .and_then(formula_model::normalize_formula_text);
                 let data = StorageCellData {
                     value,
                     formula,
                     style: None,
                 };
                let change = CellChange {
                    sheet_id: sheet_uuid,
                    row: snap.row as i64,
                    col: snap.col as i64,
                    data,
                    user_id: None,
                };

                  if let Some(autosave) = persistent.autosave.as_ref() {
                      autosave
                          .record_change(change)
                          .map_err(|e| AppStateError::Persistence(e.to_string()))?;
                  } else {
                      persistent
                          .memory
                          .record_change(change)
                          .map_err(|e| AppStateError::Persistence(e.to_string()))?;
                      persistent
                          .memory
                          .flush_dirty_pages()
                          .map_err(|e| AppStateError::Persistence(e.to_string()))?;
                  }
              }
          }

        Ok(())
    }

    fn refresh_computed_values_from_recalc_changes(
        &mut self,
        changes: &[formula_engine::RecalcValueChange],
    ) -> Result<Vec<CellUpdateData>, AppStateError> {
        let workbook = self
            .workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let date_system = workbook_date_system(workbook);
        let mut updates = Vec::new();

        // Small number of sheets; keep this lookup simple and case-insensitive.
        for change in changes {
            let Some(sheet) = workbook
                .sheets
                .iter_mut()
                .find(|s| sheet_name_eq_case_insensitive(&s.name, &change.sheet))
            else {
                continue;
            };

            let row = change.addr.row as usize;
            let col = change.addr.col as usize;

            let Some(cell) = sheet.cells.get_mut(&(row, col)) else {
                continue;
            };

            // Workbook-side computed caches only track stored formula cells.
            if cell.formula.is_none() {
                continue;
            }

            let new_value = engine_value_to_scalar(change.value.clone());
            if new_value == cell.computed_value {
                continue;
            }

            cell.computed_value = new_value.clone();
            let display_value = format_scalar_for_display_with_date_system(
                &new_value,
                cell.number_format.as_deref(),
                date_system,
            );
            updates.push(CellUpdateData {
                sheet_id: sheet.id.clone(),
                row,
                col,
                value: new_value,
                formula: cell.formula.clone(),
                display_value,
            });
        }

        Ok(updates)
    }

    fn rebuild_engine_from_workbook(&mut self) -> Result<(), AppStateError> {
        {
            let workbook = self
                .workbook
                .as_ref()
                .ok_or(AppStateError::NoWorkbookLoaded)?;
            self.engine = FormulaEngine::new();
            let (directory, filename) = workbook_file_metadata(workbook);
            self.engine
                .set_workbook_file_metadata(directory.as_deref(), filename.as_deref());
            self.engine.set_date_system(match workbook.date_system {
                formula_model::DateSystem::Excel1900 => {
                    formula_engine::date::ExcelDateSystem::EXCEL_1900
                }
                formula_model::DateSystem::Excel1904 => {
                    formula_engine::date::ExcelDateSystem::Excel1904
                }
            });
            for sheet in &workbook.sheets {
                self.engine.ensure_sheet(&sheet.name);
            }
            self.engine.set_external_value_provider(
                ColumnarExternalValueProvider::from_workbook(workbook)
                    .map(|p| -> Arc<dyn ExternalValueProvider> { Arc::new(p) }),
            );

            // Create all sheets up-front so cross-sheet formula references resolve
            // regardless of workbook sheet ordering.
            for sheet in &workbook.sheets {
                self.engine.ensure_sheet(&sheet.name);
            }

            // Load workbook-level defined names into the calculation engine so formulas referencing
            // them evaluate correctly.
            //
            // This is best-effort: invalid/unsupported definitions are ignored rather than aborting
            // workbook load.
            for defined in &workbook.defined_names {
                let name = defined.name.trim();
                let refers_to = defined.refers_to.trim();
                if name.is_empty() || refers_to.is_empty() {
                    continue;
                }

                let scope = match defined.sheet_id.as_deref() {
                    None => NameScope::Workbook,
                    Some(sheet_key) => resolve_sheet_case_insensitive(workbook, sheet_key)
                        .map(|sheet| NameScope::Sheet(sheet.name.as_str()))
                        .unwrap_or(NameScope::Workbook),
                };

                // Preserve the "refers_to" text without forcing an '=' prefix; the formula engine
                // parser accepts either form.
                let _ = self.engine.define_name(
                    name,
                    scope,
                    NameDefinition::Formula(refers_to.to_string()),
                );
            }

            // Apply imported worksheet column metadata (width/hidden flags) so worksheet information
            // functions like `CELL("width")` can observe Excel column properties on desktop.
            for sheet in &workbook.sheets {
                let sheet_name = sheet.name.as_str();
                if sheet.default_col_width.is_some() {
                    let _ = self
                        .engine
                        .set_sheet_default_col_width(sheet_name, sheet.default_col_width);
                }
                for (col_0based, props) in &sheet.col_properties {
                    if let Some(width) = props.width {
                        let _ = self.engine.set_col_width(sheet_name, *col_0based, Some(width));
                    }
                    if props.hidden {
                        let _ = self
                            .engine
                            .set_col_hidden(sheet_name, *col_0based, props.hidden);
                    }
                    if let Some(style_id) = props.style_id {
                        let _ = self
                            .engine
                            .set_col_style_id(sheet_name, *col_0based, Some(style_id));
                    }
                }
            }

            let mut style_ids_by_format: HashMap<String, u32> = HashMap::new();
            for sheet in &workbook.sheets {
                let sheet_name = &sheet.name;
                for ((row, col), cell) in sheet.cells_iter() {
                    let addr = coord_to_a1(row, col);
                    if let Some(fmt) = cell
                        .number_format
                        .as_deref()
                        .map(str::trim)
                        .filter(|s| !s.is_empty())
                    {
                        let style_id = if let Some(existing) = style_ids_by_format.get(fmt) {
                            *existing
                        } else {
                            let fmt = fmt.to_string();
                            let style_id = self.engine.intern_style(formula_model::Style {
                                number_format: Some(fmt.clone()),
                                ..Default::default()
                            });
                            style_ids_by_format.insert(fmt, style_id);
                            style_id
                        };
                        let _ = self.engine.set_cell_style_id(sheet_name, &addr, style_id);
                    }
                    if let Some(formula) = &cell.formula {
                        if cell.computed_value != CellScalar::Empty {
                            let _ = self.engine.set_cell_value(
                                sheet_name,
                                &addr,
                                scalar_to_engine_value(&cell.computed_value),
                            );
                        }
                        if self
                            .engine
                            .set_cell_formula(sheet_name, &addr, formula)
                            .is_err()
                        {
                            let _ = self.engine.set_cell_value(
                                sheet_name,
                                &addr,
                                EngineValue::Error(ErrorKind::Name),
                            );
                        }
                        continue;
                    }
                    if let Some(value) = &cell.input_value {
                        let _ = self.engine.set_cell_value(
                            sheet_name,
                            &addr,
                            scalar_to_engine_value(value),
                        );
                    }
                }
            }
        }

        // Apply persisted UI formatting metadata (sheet/col/row/run/cell layers) to the engine so
        // style-aware worksheet information functions (e.g. `CELL("protect")`, `CELL("prefix")`)
        // observe the same formatting state as the UI after engine rebuilds.
        let _ = self.apply_persistent_ui_formatting_metadata_to_engine();

        Ok(())
    }

    fn refresh_computed_values(&mut self) -> Result<Vec<CellUpdateData>, AppStateError> {
        let workbook = self
            .workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let date_system = workbook_date_system(workbook);
        let mut updates = Vec::new();

        for sheet in &mut workbook.sheets {
            let sheet_name = sheet.name.clone();
            let sheet_id = sheet.id.clone();

            for ((row, col), cell) in sheet.cells.iter_mut() {
                let new_value = if cell.formula.is_some() {
                    let addr = coord_to_a1(*row, *col);
                    engine_value_to_scalar(self.engine.get_cell_value(&sheet_name, &addr))
                } else {
                    cell.input_value.clone().unwrap_or(CellScalar::Empty)
                };

                if new_value != cell.computed_value {
                    cell.computed_value = new_value.clone();
                    let display_value =
                        format_scalar_for_display_with_date_system(
                            &new_value,
                            cell.number_format.as_deref(),
                            date_system,
                        );
                    updates.push(CellUpdateData {
                        sheet_id: sheet_id.clone(),
                        row: *row,
                        col: *col,
                        value: new_value,
                        formula: cell.formula.clone(),
                        display_value,
                    });
                }
            }
        }

        Ok(updates)
    }
}

fn workbook_file_metadata(workbook: &Workbook) -> (Option<String>, Option<String>) {
    let path = workbook
        .path
        .as_deref()
        .filter(|p| !p.trim().is_empty())
        .or_else(|| workbook.origin_path.as_deref().filter(|p| !p.trim().is_empty()));
    let Some(path) = path else {
        return (None, None);
    };

    let path = Path::new(path);
    let filename = path
        .file_name()
        .and_then(|s| s.to_str())
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| s.to_string());
    let Some(filename) = filename else {
        return (None, None);
    };
    let directory = path
        .parent()
        .and_then(|p| p.to_str())
        .map(|s| s.trim())
        .filter(|s| !s.is_empty())
        .map(|s| s.to_string());

    (directory, Some(filename))
}

fn dedupe_updates(updates: Vec<CellUpdateData>) -> Vec<CellUpdateData> {
    let mut out: Vec<CellUpdateData> = Vec::new();
    let mut index_by_key = std::collections::HashMap::<(String, usize, usize), usize>::new();

    for update in updates {
        let key = (update.sheet_id.clone(), update.row, update.col);
        if let Some(idx) = index_by_key.get(&key).copied() {
            out[idx] = update;
        } else {
            index_by_key.insert(key, out.len());
            out.push(update);
        }
    }

    out
}

fn resolve_sheet_case_insensitive<'a>(
    workbook: &'a Workbook,
    sheet_id: &str,
) -> Option<&'a crate::file_io::Sheet> {
    workbook
        .sheets
        .iter()
        .find(|s| s.id.eq_ignore_ascii_case(sheet_id))
        .or_else(|| {
            workbook
                .sheets
                .iter()
                .find(|s| sheet_name_eq_case_insensitive(&s.name, sheet_id))
        })
}

fn format_auditing_cells(
    workbook: &Workbook,
    active_sheet_name: &str,
    cells: Vec<(usize, CellAddr)>,
) -> Vec<String> {
    let mut out = Vec::with_capacity(cells.len());
    for (sheet_idx, addr) in cells {
        let Some(sheet) = workbook.sheets.get(sheet_idx) else {
            continue;
        };
        let a1 = coord_to_a1(addr.row as usize, addr.col as usize);
        if sheet_name_eq_case_insensitive(&sheet.name, active_sheet_name) {
            out.push(a1);
        } else {
            out.push(format!("{}!{}", quote_sheet_name(&sheet.name), a1));
        }
    }
    out
}

fn expand_precedent_nodes_to_cells(
    nodes: &[PrecedentNode],
    limit: usize,
    kind: &str,
) -> Result<Vec<(usize, CellAddr)>, AppStateError> {
    let mut out: Vec<(usize, CellAddr)> = Vec::new();
    let mut seen: HashSet<(usize, u32, u32)> = HashSet::new();

    let mut push_cell = |sheet: usize, addr: CellAddr| -> Result<(), AppStateError> {
        if !seen.insert((sheet, addr.row, addr.col)) {
            return Ok(());
        }
        if out.len() >= limit {
            return Err(AppStateError::AuditingTooLarge {
                kind: kind.to_string(),
                limit,
            });
        }
        out.push((sheet, addr));
        Ok(())
    };

    for node in nodes {
        match *node {
            PrecedentNode::Cell { sheet, addr } => {
                push_cell(sheet, addr)?;
            }
            PrecedentNode::Range { sheet, start, end } => {
                let row_start = start.row.min(end.row);
                let row_end = start.row.max(end.row);
                let col_start = start.col.min(end.col);
                let col_end = start.col.max(end.col);

                let mut row = row_start;
                loop {
                    let mut col = col_start;
                    loop {
                        push_cell(sheet, CellAddr { row, col })?;
                        if col == col_end {
                            break;
                        }
                        col = col.saturating_add(1);
                    }
                    if row == row_end {
                        break;
                    }
                    row = row.saturating_add(1);
                }
            }
            // External workbook references (e.g. `[Book.xlsx]Sheet1!A1`) are not currently
            // surfaced in the desktop UI auditing overlays, which only highlight in-workbook
            // cells. Skip them rather than failing the entire request.
            PrecedentNode::ExternalCell { .. } | PrecedentNode::ExternalRange { .. } => {}
            PrecedentNode::SpillRange {
                sheet,
                origin,
                start,
                end,
            } => {
                push_cell(sheet, origin)?;

                let row_start = start.row.min(end.row);
                let row_end = start.row.max(end.row);
                let col_start = start.col.min(end.col);
                let col_end = start.col.max(end.col);

                let mut row = row_start;
                loop {
                    let mut col = col_start;
                    loop {
                        push_cell(sheet, CellAddr { row, col })?;
                        if col == col_end {
                            break;
                        }
                        col = col.saturating_add(1);
                    }
                    if row == row_end {
                        break;
                    }
                    row = row.saturating_add(1);
                }
            }
        }
    }

    out.sort_by_key(|(sheet, addr)| (*sheet, addr.row, addr.col));
    Ok(out)
}

fn pivot_value_to_scalar(value: &PivotValue, date_system: formula_engine::date::ExcelDateSystem) -> CellScalar {
    match value {
        PivotValue::Blank => CellScalar::Empty,
        PivotValue::Number(n) => CellScalar::Number(*n),
        PivotValue::Date(d) => {
            // Excel stores dates as serial numbers + number format. Preserve that invariant so
            // pivot label cells can be referenced by downstream formulas as numeric date serials.
            let excel_date = formula_engine::date::ExcelDate::new(d.year(), d.month() as u8, d.day() as u8);
            match formula_engine::date::ymd_to_serial(excel_date, date_system) {
                Ok(serial) => CellScalar::Number(serial as f64),
                Err(_) => CellScalar::Text(d.to_string()),
            }
        }
        PivotValue::Text(s) => CellScalar::Text(s.clone()),
        PivotValue::Bool(b) => CellScalar::Bool(*b),
    }
}

fn pivot_value_to_scalar_opt(
    value: &PivotValue,
    date_system: formula_engine::date::ExcelDateSystem,
) -> Option<CellScalar> {
    match value {
        PivotValue::Blank => None,
        other => Some(pivot_value_to_scalar(other, date_system)),
    }
}

fn normalize_pivot_grid(mut grid: Vec<Vec<PivotValue>>) -> Vec<Vec<PivotValue>> {
    if grid.is_empty() {
        return vec![vec![PivotValue::Blank]];
    }

    let mut cols = grid.iter().map(|row| row.len()).max().unwrap_or(0);
    cols = cols.max(1);

    for row in &mut grid {
        while row.len() < cols {
            row.push(PivotValue::Blank);
        }
    }

    grid
}

fn refresh_pivot_registration(
    workbook: &mut Workbook,
    engine: &mut FormulaEngine,
    pivot: &mut PivotRegistration,
) -> Result<Vec<CellUpdateData>, AppStateError> {
    if pivot.source_range.start_row > pivot.source_range.end_row
        || pivot.source_range.start_col > pivot.source_range.end_col
    {
        return Err(AppStateError::InvalidRange {
            start_row: pivot.source_range.start_row,
            start_col: pivot.source_range.start_col,
            end_row: pivot.source_range.end_row,
            end_col: pivot.source_range.end_col,
        });
    }

    let source_rows = pivot
        .source_range
        .end_row
        .checked_sub(pivot.source_range.start_row)
        .and_then(|d| d.checked_add(1))
        .unwrap_or(usize::MAX);
    let source_cols = pivot
        .source_range
        .end_col
        .checked_sub(pivot.source_range.start_col)
        .and_then(|d| d.checked_add(1))
        .unwrap_or(usize::MAX);

    if source_rows > MAX_RANGE_DIM || source_cols > MAX_RANGE_DIM {
        return Err(AppStateError::RangeDimensionTooLarge {
            rows: source_rows,
            cols: source_cols,
            limit: MAX_RANGE_DIM,
        });
    }

    let source_cells = (source_rows as u128) * (source_cols as u128);
    if source_cells > MAX_RANGE_CELLS_PER_CALL as u128 {
        return Err(AppStateError::RangeTooLarge {
            rows: source_rows,
            cols: source_cols,
            limit: MAX_RANGE_CELLS_PER_CALL,
        });
    }

    let (cache, source_col_number_formats) = {
        let source_sheet = workbook
            .sheet(&pivot.source_sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(pivot.source_sheet_id.clone()))?;
        let sheet_name = source_sheet.name.clone();

        let start_row = u32::try_from(pivot.source_range.start_row)
            .map_err(|_| AppStateError::Pivot("pivot source range start row out of range".to_string()))?;
        let start_col = u32::try_from(pivot.source_range.start_col)
            .map_err(|_| AppStateError::Pivot("pivot source range start col out of range".to_string()))?;
        let end_row = u32::try_from(pivot.source_range.end_row)
            .map_err(|_| AppStateError::Pivot("pivot source range end row out of range".to_string()))?;
        let end_col = u32::try_from(pivot.source_range.end_col)
            .map_err(|_| AppStateError::Pivot("pivot source range end col out of range".to_string()))?;

        let range = formula_model::Range::new(
            formula_model::CellRef::new(start_row, start_col),
            formula_model::CellRef::new(end_row, end_col),
        );

        let cache = engine
            .pivot_cache_from_range(&sheet_name, range)
            .map_err(|e| AppStateError::Pivot(e.to_string()))?;

        let cells = source_sheet.get_range_cells(
            pivot.source_range.start_row,
            pivot.source_range.start_col,
            pivot.source_range.end_row,
            pivot.source_range.end_col,
        );

        // Best-effort number format inference per source column (used to format pivot label dates).
        //
        // Note: This intentionally ignores the header row. We use the first explicit number format
        // found in the data rows.
        let mut col_number_formats = vec![None; source_cols];
        for c in 0..source_cols {
            for r in 1..source_rows {
                if let Some(fmt) = cells[r][c]
                    .number_format
                    .as_deref()
                    .map(str::trim)
                    .filter(|s| !s.is_empty())
                {
                    col_number_formats[c] = Some(fmt.to_string());
                    break;
                }
            }
        }

        (cache, col_number_formats)
    };

    // Infer date formats for row-field label columns so pivot date labels render as date serials
    // with a date number format (Excel semantics).
    let row_field_number_formats: Vec<Option<String>> = pivot
        .config
        .row_fields
        .iter()
        .map(|field| {
            let idx = cache.field_index_ref(&field.source_field)?;
            source_col_number_formats.get(idx).cloned().unwrap_or(None)
        })
        .collect();
    let result = PivotEngine::calculate(&cache, &pivot.config)
        .map_err(|e| AppStateError::Pivot(e.to_string()))?;

    let grid = normalize_pivot_grid(result.data);
    let row_count = grid.len();
    let col_count = grid[0].len();

    if row_count > MAX_RANGE_DIM || col_count > MAX_RANGE_DIM {
        return Err(AppStateError::RangeDimensionTooLarge {
            rows: row_count,
            cols: col_count,
            limit: MAX_RANGE_DIM,
        });
    }

    let output_cells = (row_count as u128) * (col_count as u128);
    if output_cells > MAX_RANGE_CELLS_PER_CALL as u128 {
        return Err(AppStateError::RangeTooLarge {
            rows: row_count,
            cols: col_count,
            limit: MAX_RANGE_CELLS_PER_CALL,
        });
    }

    let dest_start_row = pivot.destination.row;
    let dest_start_col = pivot.destination.col;

    let end_row = dest_start_row
        .checked_add(row_count.saturating_sub(1))
        .ok_or_else(|| AppStateError::Pivot("pivot output range row overflow".to_string()))?;
    let end_col = dest_start_col
        .checked_add(col_count.saturating_sub(1))
        .ok_or_else(|| AppStateError::Pivot("pivot output range col overflow".to_string()))?;

    let next_range = CellRect {
        start_row: dest_start_row,
        start_col: dest_start_col,
        end_row,
        end_col,
    };

    let dest_sheet_id = pivot.destination.sheet_id.clone();
    let dest_sheet_name = workbook
        .sheet(&dest_sheet_id)
        .ok_or_else(|| AppStateError::UnknownSheet(dest_sheet_id.clone()))?
        .name
        .clone();

    let prev_range = pivot.last_output_range.clone();
    let pivot_date_system = engine.date_system();
    let workbook_fmt_date_system = workbook_date_system(workbook);

    fn rect_to_model_range(rect: &CellRect) -> Result<ModelRange, AppStateError> {
        let start_row = u32::try_from(rect.start_row)
            .map_err(|_| AppStateError::Pivot("pivot output range row overflow".to_string()))?;
        let start_col = u32::try_from(rect.start_col)
            .map_err(|_| AppStateError::Pivot("pivot output range col overflow".to_string()))?;
        let end_row = u32::try_from(rect.end_row)
            .map_err(|_| AppStateError::Pivot("pivot output range row overflow".to_string()))?;
        let end_col = u32::try_from(rect.end_col)
            .map_err(|_| AppStateError::Pivot("pivot output range col overflow".to_string()))?;
        Ok(ModelRange::new(
            ModelCellRef::new(start_row, start_col),
            ModelCellRef::new(end_row, end_col),
        ))
    }

    fn stale_rects(prev: &CellRect, next: &CellRect) -> Vec<CellRect> {
        let inter_start_row = prev.start_row.max(next.start_row);
        let inter_start_col = prev.start_col.max(next.start_col);
        let inter_end_row = prev.end_row.min(next.end_row);
        let inter_end_col = prev.end_col.min(next.end_col);

        if inter_start_row > inter_end_row || inter_start_col > inter_end_col {
            return vec![prev.clone()];
        }

        let mut out = Vec::new();

        if prev.start_row < inter_start_row {
            out.push(CellRect {
                start_row: prev.start_row,
                end_row: inter_start_row.saturating_sub(1),
                start_col: prev.start_col,
                end_col: prev.end_col,
            });
        }
        if inter_end_row < prev.end_row {
            out.push(CellRect {
                start_row: inter_end_row.saturating_add(1),
                end_row: prev.end_row,
                start_col: prev.start_col,
                end_col: prev.end_col,
            });
        }
        if prev.start_col < inter_start_col {
            out.push(CellRect {
                start_row: inter_start_row,
                end_row: inter_end_row,
                start_col: prev.start_col,
                end_col: inter_start_col.saturating_sub(1),
            });
        }
        if inter_end_col < prev.end_col {
            out.push(CellRect {
                start_row: inter_start_row,
                end_row: inter_end_row,
                start_col: inter_end_col.saturating_add(1),
                end_col: prev.end_col,
            });
        }

        out
    }

    fn pivot_value_to_engine_value(
        value: &PivotValue,
        date_system: formula_engine::date::ExcelDateSystem,
    ) -> EngineValue {
        let scalar = pivot_value_to_scalar(value, date_system);
        scalar_to_engine_value(&scalar)
    }

    // Apply pivot output to the formula engine using bulk operations.
    //
    // Pivot refresh may touch thousands of cells; using per-cell engine APIs can trigger
    // recalculation repeatedly when calculation mode is automatic.
    if let Some(prev) = prev_range.as_ref() {
        for rect in stale_rects(prev, &next_range) {
            let range = rect_to_model_range(&rect)?;
            engine
                .clear_range(&dest_sheet_name, range, false)
                .map_err(|e| AppStateError::Engine(e.to_string()))?;
        }
    }

    let next_model_range = rect_to_model_range(&next_range)?;
    let engine_values: Vec<Vec<EngineValue>> = grid
        .iter()
        .map(|row| {
            row.iter()
                .map(|value| pivot_value_to_engine_value(value, pivot_date_system))
                .collect()
        })
        .collect();
    engine
        .set_range_values(&dest_sheet_name, next_model_range, &engine_values, false)
        .map_err(|e| AppStateError::Engine(e.to_string()))?;
    // Register the pivot metadata in the formula engine so `GETPIVOTDATA` can resolve pivots
    // without scanning the rendered output grid.
    let pivot_table = EnginePivotTable {
        id: pivot.id.clone(),
        name: pivot.name.clone(),
        config: pivot.config.clone(),
        cache,
    };
    engine
        .register_pivot_table(&dest_sheet_name, next_model_range, pivot_table)
        .map_err(|e| AppStateError::Pivot(e.to_string()))?;

    let mut updates = Vec::new();

    {
        let sheet = workbook
            .sheet_mut(&dest_sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(dest_sheet_id.clone()))?;

        // Clear stale cells from the previous rendered range that now fall outside the updated output.
        //
        // IMPORTANT: Only clear cells the pivot previously wrote. Using a union/bounding-box range can
        // accidentally clear unrelated cells when the output grows in one dimension while shrinking in
        // the other (e.g. taller-but-narrower), because the bounding box includes cells that were never
        // owned by the pivot output.
        if let Some(prev) = prev_range.as_ref() {
            for rect in stale_rects(prev, &next_range) {
                for row in rect.start_row..=rect.end_row {
                        for col in rect.start_col..=rect.end_col {
                            let desired_scalar = CellScalar::Empty;
                            let display_value = format_scalar_for_display_with_date_system(
                                &desired_scalar,
                                None,
                                workbook_fmt_date_system,
                            );

                            let existing = sheet.get_cell(row, col);
                            let changed =
                                existing.formula.is_some() || existing.computed_value != desired_scalar;
                            if !changed {
                            continue;
                        }

                        sheet.set_cell(row, col, Cell::empty());

                        updates.push(CellUpdateData {
                            sheet_id: dest_sheet_id.clone(),
                            row,
                            col,
                            value: desired_scalar,
                            formula: None,
                            display_value,
                        });
                    }
                }
            }
        }

        // Write the new pivot output values into the destination range.
        let value_field_count = pivot.config.value_fields.len();
        let row_label_width = match pivot.config.layout {
            formula_engine::pivot::Layout::Compact => 1,
            formula_engine::pivot::Layout::Outline | formula_engine::pivot::Layout::Tabular => {
                pivot.config.row_fields.len()
            }
        };

        for r in 0..row_count {
            for c in 0..col_count {
                let row = next_range.start_row + r;
                let col = next_range.start_col + c;

                let pv = &grid[r][c];
                let desired_number_format = if matches!(pv, PivotValue::Date(_)) {
                    row_field_number_formats
                        .get(c)
                        .cloned()
                        .unwrap_or(None)
                        .filter(|fmt| !fmt.trim().is_empty())
                        .or_else(|| Some("m/d/yyyy".to_string()))
                } else if r > 0 && value_field_count > 0 && c >= row_label_width {
                    let vf_idx = (c - row_label_width) % value_field_count;
                    let vf = &pivot.config.value_fields[vf_idx];
                    if let Some(fmt) = vf
                        .number_format
                        .as_deref()
                        .map(str::trim)
                        .filter(|s| !s.is_empty())
                    {
                        Some(fmt.to_string())
                    } else if matches!(
                        vf.show_as.unwrap_or(formula_engine::pivot::ShowAsType::Normal),
                        formula_engine::pivot::ShowAsType::PercentOfGrandTotal
                            | formula_engine::pivot::ShowAsType::PercentOfRowTotal
                            | formula_engine::pivot::ShowAsType::PercentOfColumnTotal
                            | formula_engine::pivot::ShowAsType::PercentOf
                            | formula_engine::pivot::ShowAsType::PercentDifferenceFrom
                    ) {
                        Some("0.00%".to_string())
                    } else {
                        None
                    }
                } else {
                    None
                };

                let desired_scalar = pivot_value_to_scalar(pv, pivot_date_system);
                let desired_opt = pivot_value_to_scalar_opt(pv, pivot_date_system);
                let display_value = format_scalar_for_display_with_date_system(
                    &desired_scalar,
                    desired_number_format.as_deref(),
                    workbook_fmt_date_system,
                );

                let existing = sheet.get_cell(row, col);
                let changed = existing.formula.is_some()
                    || existing.computed_value != desired_scalar
                    || (desired_number_format.is_some()
                        && existing.number_format.as_deref() != desired_number_format.as_deref());
                if !changed {
                    continue;
                }

                let mut new_cell = match desired_opt.as_ref() {
                    Some(s) => Cell::from_literal(Some(s.clone())),
                    None => Cell::empty(),
                };
                new_cell.number_format = desired_number_format.clone();
                sheet.set_cell(row, col, new_cell);

                updates.push(CellUpdateData {
                    sheet_id: dest_sheet_id.clone(),
                    row,
                    col,
                    value: desired_scalar,
                    formula: None,
                    display_value,
                });
            }
        }
    }

    // Track the *actual* rendered range so future refreshes only clear cells this pivot most
    // recently wrote. (If the output shrinks, cleared cells should be released for user edits.)
    pivot.last_output_range = Some(next_range);
    Ok(updates)
}

fn format_scalar_for_display_with_date_system(
    value: &CellScalar,
    number_format: Option<&str>,
    date_system: formula_format::DateSystem,
) -> String {
    match value {
        CellScalar::Number(n) => {
            let options = FormatOptions {
                date_system,
                ..FormatOptions::default()
            };
            let formatted = format_value(FormatValue::Number(*n), number_format, &options);
            formatted.text
        }
        other => other.display(),
    }
}

fn workbook_date_system(workbook: &Workbook) -> formula_format::DateSystem {
    match workbook.date_system {
        formula_model::DateSystem::Excel1900 => formula_format::DateSystem::Excel1900,
        formula_model::DateSystem::Excel1904 => formula_format::DateSystem::Excel1904,
    }
}
fn resolve_cell_ref(
    workbook: &Workbook,
    default_sheet_id: &str,
    default_sheet_name: &str,
    cell: &WhatIfCellRef,
) -> Result<(String, usize, usize), AppStateError> {
    let raw = cell.as_str().trim();
    let (sheet_name, addr) = match raw.split_once('!') {
        Some((sheet_raw, addr_raw)) => {
            let sheet_raw = sheet_raw.trim();
            let sheet_name = if let Some(inner) = sheet_raw
                .strip_prefix('\'')
                .and_then(|s| s.strip_suffix('\''))
            {
                inner.replace("''", "'")
            } else {
                sheet_raw.to_string()
            };
            (sheet_name, addr_raw.trim().to_string())
        }
        None => (default_sheet_name.to_string(), raw.to_string()),
    };

    let sheet_id = if sheet_name_eq_case_insensitive(&sheet_name, default_sheet_name) {
        default_sheet_id.to_string()
    } else {
        workbook
            .sheets
            .iter()
            .find(|s| sheet_name_eq_case_insensitive(&s.name, &sheet_name))
            .map(|s| s.id.clone())
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_name.clone()))?
    };

    let addr = parse_a1(&addr).map_err(|e| AppStateError::WhatIf(e.to_string()))?;
    Ok((sheet_id, addr.row as usize, addr.col as usize))
}

fn normalize_formula(formula: Option<String>) -> Option<String> {
    let formula = formula?;
    let display = formula_model::display_formula_text(&formula);
    if display.is_empty() {
        None
    } else {
        Some(display)
    }
}

fn coord_to_a1(row: usize, col: usize) -> String {
    // Prefer the shared A1 formatter for consistency with the rest of the codebase and to avoid
    // debug overflow panics when converting 0-based coordinates to 1-based A1 notation
    // (e.g. `row == u32::MAX` on 32-bit targets).
    match (u32::try_from(row), u32::try_from(col)) {
        (Ok(row), Ok(col)) => formula_model::CellRef::new(row, col).to_a1(),
        _ => {
            let row_1_based = u64::try_from(row)
                .unwrap_or(u64::MAX)
                .saturating_add(1);
            format!("{}{}", col_index_to_letters(col), row_1_based)
        }
    }
}

fn quote_sheet_name(name: &str) -> String {
    // Excel escapes single quotes inside a quoted sheet name by doubling them.
    let escaped = name.replace('\'', "''");
    format!("'{escaped}'")
}

fn col_index_to_letters(col: usize) -> String {
    // Excel columns are base-26 with A=1..Z=26.
    //
    // Do arithmetic in u64 so very large `usize` indices (e.g. u32::MAX on wasm32/32-bit targets)
    // don't overflow on the 0->1 conversion.
    let mut col = u64::try_from(col).unwrap_or(u64::MAX).saturating_add(1);
    let mut letters = Vec::new();
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        letters.push((b'A' + rem) as char);
        col = (col - 1) / 26;
    }
    letters.iter().rev().collect()
}

fn scalar_to_what_if_value(value: &CellScalar) -> WhatIfCellValue {
    match value {
        CellScalar::Empty => WhatIfCellValue::Blank,
        CellScalar::Number(n) => WhatIfCellValue::Number(*n),
        CellScalar::Text(s) => WhatIfCellValue::Text(s.clone()),
        CellScalar::Bool(b) => WhatIfCellValue::Bool(*b),
        CellScalar::Error(e) => WhatIfCellValue::Text(e.clone()),
    }
}

fn what_if_value_to_scalar(value: &WhatIfCellValue) -> Option<CellScalar> {
    match value {
        WhatIfCellValue::Blank => None,
        WhatIfCellValue::Number(n) => Some(CellScalar::Number(*n)),
        WhatIfCellValue::Text(s) => Some(CellScalar::Text(s.clone())),
        WhatIfCellValue::Bool(b) => Some(CellScalar::Bool(*b)),
    }
}

fn scalar_to_storage_value(value: &CellScalar) -> formula_model::CellValue {
    match value {
        CellScalar::Empty => formula_model::CellValue::Empty,
        CellScalar::Number(n) => formula_model::CellValue::Number(*n),
        CellScalar::Text(s) => formula_model::CellValue::String(s.clone()),
        CellScalar::Bool(b) => formula_model::CellValue::Boolean(*b),
        CellScalar::Error(e) => formula_model::CellValue::Error(
            e.parse::<formula_model::ErrorValue>()
                .unwrap_or(formula_model::ErrorValue::Unknown),
        ),
    }
}

fn scalar_to_engine_value(value: &CellScalar) -> EngineValue {
    match value {
        CellScalar::Empty => EngineValue::Blank,
        CellScalar::Number(n) => EngineValue::Number(*n),
        CellScalar::Text(s) => EngineValue::Text(s.clone()),
        CellScalar::Bool(b) => EngineValue::Bool(*b),
        CellScalar::Error(code) => match parse_error_kind(code) {
            Some(kind) => EngineValue::Error(kind),
            None => EngineValue::Text(code.clone()),
        },
    }
}

fn engine_value_to_scalar(value: EngineValue) -> CellScalar {
    match value {
        EngineValue::Blank => CellScalar::Empty,
        EngineValue::Number(n) => CellScalar::Number(n),
        EngineValue::Text(s) => CellScalar::Text(s),
        EngineValue::Entity(v) => CellScalar::Text(v.display),
        EngineValue::Record(v) => {
            if let Some(display_field) = v.display_field.as_deref() {
                if let Some(value) = v.get_field_case_insensitive(display_field) {
                    let text = value
                        .coerce_to_string()
                        .unwrap_or_else(|e| e.as_code().to_string());
                    return CellScalar::Text(text);
                }
            }
            CellScalar::Text(v.display)
        }
        EngineValue::Bool(b) => CellScalar::Bool(b),
        EngineValue::Error(e) => CellScalar::Error(e.as_code().to_string()),
        // Reference values are not meant to be stored directly in cells; when they
        // leak out as a final value we treat them like Excel does (a #VALUE! error).
        EngineValue::Reference(_) | EngineValue::ReferenceUnion(_) => {
            CellScalar::Error("#VALUE!".to_string())
        }
        EngineValue::Array(arr) => engine_value_to_scalar(arr.top_left()),
        EngineValue::Lambda(_) => CellScalar::Error("#CALC!".to_string()),
        EngineValue::Spill { .. } => CellScalar::Error("#SPILL!".to_string()),
    }
}

fn parse_error_kind(value: &str) -> Option<ErrorKind> {
    let trimmed = value.trim();
    match trimmed {
        "Null" => Some(ErrorKind::Null),
        "Div0" => Some(ErrorKind::Div0),
        "Value" => Some(ErrorKind::Value),
        "Ref" => Some(ErrorKind::Ref),
        "Name" => Some(ErrorKind::Name),
        "Num" => Some(ErrorKind::Num),
        "NA" => Some(ErrorKind::NA),
        "GettingData" => Some(ErrorKind::GettingData),
        "Spill" => Some(ErrorKind::Spill),
        "Calc" => Some(ErrorKind::Calc),
        "Field" => Some(ErrorKind::Field),
        "Connect" => Some(ErrorKind::Connect),
        "Blocked" => Some(ErrorKind::Blocked),
        "Unknown" => Some(ErrorKind::Unknown),
        other => ErrorKind::from_code(other),
    }
}

fn apply_snapshot_to_engine(
    engine: &mut FormulaEngine,
    sheet_name: &str,
    row: usize,
    col: usize,
    value: &Option<CellScalar>,
    formula: &Option<String>,
) {
    let addr = coord_to_a1(row, col);

    if let Some(formula) = formula.as_deref() {
        if engine.set_cell_formula(sheet_name, &addr, formula).is_err() {
            let _ = engine.set_cell_value(sheet_name, &addr, EngineValue::Error(ErrorKind::Name));
        }
        return;
    }

    let engine_value = value
        .as_ref()
        .map(scalar_to_engine_value)
        .unwrap_or(EngineValue::Blank);
    let _ = engine.set_cell_value(sheet_name, &addr, engine_value);
}

struct ColumnarExternalValueProvider {
    tables: HashMap<String, Arc<ColumnarTable>>,
}

impl ColumnarExternalValueProvider {
    fn from_workbook(workbook: &Workbook) -> Option<Self> {
        let mut tables = HashMap::new();
        for sheet in &workbook.sheets {
            if let Some(table) = sheet.columnar.clone() {
                tables.insert(sheet.name.clone(), table);
            }
        }
        if tables.is_empty() {
            None
        } else {
            Some(Self { tables })
        }
    }
}

impl ExternalValueProvider for ColumnarExternalValueProvider {
    fn get(&self, sheet: &str, addr: CellAddr) -> Option<EngineValue> {
        let table = self.tables.get(sheet)?;
        let row = usize::try_from(addr.row).ok()?;
        let col = usize::try_from(addr.col).ok()?;
        if row >= table.row_count() || col >= table.column_count() {
            return None;
        }

        let col_type = table
            .schema()
            .get(col)
            .map(|c| c.column_type)
            .unwrap_or(ColumnarType::String);
        Some(columnar_to_engine_value(table.get_cell(row, col), col_type))
    }
}

fn columnar_to_engine_value(value: ColumnarValue, column_type: ColumnarType) -> EngineValue {
    match value {
        ColumnarValue::Null => EngineValue::Blank,
        ColumnarValue::Number(v) => EngineValue::Number(v),
        ColumnarValue::Boolean(v) => EngineValue::Bool(v),
        ColumnarValue::String(v) => EngineValue::Text(v.as_ref().to_string()),
        ColumnarValue::DateTime(v) => EngineValue::Number(v as f64),
        ColumnarValue::Currency(v) => match column_type {
            ColumnarType::Currency { scale } => {
                let denom = 10f64.powi(scale as i32);
                EngineValue::Number(v as f64 / denom)
            }
            _ => EngineValue::Number(v as f64),
        },
        ColumnarValue::Percentage(v) => match column_type {
            ColumnarType::Percentage { scale } => {
                let denom = 10f64.powi(scale as i32);
                EngineValue::Number(v as f64 / denom)
            }
            _ => EngineValue::Number(v as f64),
        },
    }
}

fn default_sheet_print_settings(sheet_name: String) -> SheetPrintSettings {
    SheetPrintSettings {
        sheet_name,
        print_area: None,
        print_titles: None,
        page_setup: PageSetup::default(),
        manual_page_breaks: ManualPageBreaks::default(),
    }
}

fn ensure_sheet_print_settings<'a>(
    sheets: &'a mut Vec<SheetPrintSettings>,
    sheet_name: &str,
) -> &'a mut SheetPrintSettings {
    if let Some(idx) = sheets
        .iter()
        .position(|s| sheet_name_eq_case_insensitive(&s.sheet_name, sheet_name))
    {
        if sheets[idx].sheet_name != sheet_name {
            sheets[idx].sheet_name = sheet_name.to_string();
        }
        return &mut sheets[idx];
    }

    sheets.push(default_sheet_print_settings(sheet_name.to_string()));
    let idx = sheets.len().saturating_sub(1);
    &mut sheets[idx]
}

fn seed_sheet_formatting_metadata_from_model(
    storage: &formula_storage::Storage,
    sheet_metas: &[formula_storage::SheetMeta],
    model: &formula_model::Workbook,
) -> formula_storage::storage::Result<()> {
    for sheet_meta in sheet_metas {
        let mut cell_formats = Vec::new();
        if let Some(sheet) = model
            .sheets
            .iter()
            .find(|s| sheet_name_eq_case_insensitive(&s.name, &sheet_meta.name))
        {
            for (cell_ref, cell) in sheet.iter_cells() {
                if cell.style_id == 0 {
                    continue;
                }
                let Some(style) = model.styles.get(cell.style_id) else {
                    continue;
                };
                let format = serde_json::to_value(style)?;
                cell_formats.push(serde_json::json!({
                    "row": cell_ref.row,
                    "col": cell_ref.col,
                    "format": format,
                }));
            }
        }

        // Keep output deterministic regardless of hash map iteration order upstream.
        cell_formats.sort_by(|a, b| {
            let row_a = a.get("row").and_then(|v| v.as_u64()).unwrap_or(u64::MAX);
            let row_b = b.get("row").and_then(|v| v.as_u64()).unwrap_or(u64::MAX);
            if row_a == row_b {
                let col_a = a.get("col").and_then(|v| v.as_u64()).unwrap_or(u64::MAX);
                let col_b = b.get("col").and_then(|v| v.as_u64()).unwrap_or(u64::MAX);
                col_a.cmp(&col_b)
            } else {
                row_a.cmp(&row_b)
            }
        });

        let formatting = serde_json::json!({
            "schemaVersion": 1,
            "defaultFormat": null,
            "rowFormats": [],
            "colFormats": [],
            "cellFormats": cell_formats,
        });

        storage.update_sheet_metadata(sheet_meta.id, move |metadata| {
            let mut root = match metadata {
                Some(JsonValue::Object(map)) => map,
                _ => serde_json::Map::new(),
            };
            root.insert(FORMULA_UI_FORMATTING_METADATA_KEY.to_string(), formatting);
            Ok(Some(JsonValue::Object(root)))
        })?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::file_io::{read_xlsx_blocking, write_xlsx_blocking};
    use crate::resource_limits::{MAX_ORIGIN_XLSX_BYTES, MAX_RANGE_CELLS_PER_CALL, MAX_RANGE_DIM};
    use formula_engine::pivot::{
        AggregationType, GrandTotals, Layout, PivotConfig, PivotField, PivotFieldRef,
        SubtotalPosition, ValueField,
    };
    use formula_engine::what_if::monte_carlo::{Distribution, InputDistribution};
    use formula_model::import::{import_csv_to_columnar_table, CsvOptions};
    use formula_xlsx::drawingml::{
        PreservedChartSheet, PreservedDrawingParts, PreservedSheetControls, PreservedSheetDrawingHF,
        PreservedSheetDrawings, PreservedSheetOleObjects, PreservedSheetPicture,
        SheetDrawingRelationship, SheetRelationshipStub, SheetRelationshipStubWithType,
    };
    use formula_xlsx::pivots::preserve::PreservedSheetPivotTables;
    use formula_xlsx::{PreservedPivotParts, RelationshipStub};

    fn bytes_over_origin_xlsx_retention_limit() -> Arc<[u8]> {
        static BYTES: std::sync::OnceLock<Arc<[u8]>> = std::sync::OnceLock::new();
        BYTES
            .get_or_init(|| Arc::<[u8]>::from(vec![0u8; MAX_ORIGIN_XLSX_BYTES + 1]))
            .clone()
    }

    #[test]
    fn record_cells_use_display_field_when_degrading_to_text() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();
        let sheet_name = workbook.sheets[0].name.clone();

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut record = formula_engine::Record::new("Fallback").field(
            "Name",
            EngineValue::Text("Apple".to_string()),
        );
        record.display_field = Some("Name".to_string());

        state
            .engine
            .set_cell_value(&sheet_name, "A1", EngineValue::Record(record))
            .expect("set engine record value");

        let cell = state.get_cell(&sheet_id, 0, 0).expect("read A1");
        assert_eq!(cell.value, CellScalar::Text("Apple".to_string()));
        assert_eq!(cell.display_value, "Apple");
    }

    #[test]
    fn record_cells_use_display_field_error_code_when_coercion_fails() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();
        let sheet_name = workbook.sheets[0].name.clone();

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut record = formula_engine::Record::new("Fallback").field(
            "Name",
            EngineValue::Error(ErrorKind::Value),
        );
        record.display_field = Some("Name".to_string());

        state
            .engine
            .set_cell_value(&sheet_name, "A1", EngineValue::Record(record))
            .expect("set engine record value");

        let cell = state.get_cell(&sheet_id, 0, 0).expect("read A1");
        assert_eq!(cell.value, CellScalar::Text("#VALUE!".to_string()));
        assert_eq!(cell.display_value, "#VALUE!");
    }

    #[test]
    fn add_sheet_inserts_after_middle_sheet() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let inserted = state
            .add_sheet(
                "Inserted".to_string(),
                None,
                Some("sHeEt2".to_string()),
                None,
            )
            .expect("add sheet succeeds");
        let inserted_id = inserted.id;

        let workbook = state.get_workbook().expect("workbook loaded");
        let ids = workbook
            .sheets
            .iter()
            .map(|sheet| sheet.id.clone())
            .collect::<Vec<_>>();

        assert_eq!(
            ids,
            vec![
                "Sheet1".to_string(),
                "Sheet2".to_string(),
                inserted_id,
                "Sheet3".to_string()
            ]
        );
    }

    #[test]
    fn cell_width_reflects_imported_column_properties() {
        let mut model = formula_model::Workbook::new();
        let sheet_id = model.add_sheet("Sheet1").expect("add sheet");

        {
            let sheet = model.sheet_mut(sheet_id).expect("sheet exists");
            sheet.set_col_width(0, Some(20.0));
            sheet.set_col_hidden(1, true);
        }

        let mut cursor = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&model, &mut cursor).expect("write xlsx");
        let xlsx_bytes = cursor.into_inner();

        let mut tmp = tempfile::Builder::new()
            .suffix(".xlsx")
            .tempfile()
            .expect("temp xlsx file");
        use std::io::Write as _;
        tmp.as_file_mut()
            .write_all(&xlsx_bytes)
            .expect("write temp xlsx bytes");
        tmp.as_file_mut().flush().expect("flush temp xlsx bytes");

        let workbook = read_xlsx_blocking(tmp.path()).expect("read xlsx via desktop importer");
        let app_sheet_id = workbook
            .sheets
            .first()
            .map(|s| s.id.clone())
            .expect("imported workbook has a sheet");

        let mut state = AppState::new();
        state.load_workbook(workbook);

        // Column A width override is propagated into the engine.
        state
            .set_cell(
                &app_sheet_id,
                0,
                2,
                None,
                Some("=CELL(\"width\",A1)".to_string()),
        )
            .expect("set width formula");
        let c1 = state.get_cell(&app_sheet_id, 0, 2).expect("read C1");
        match c1.value {
            CellScalar::Number(v) => assert!(
                (v - 20.1).abs() < 0.2,
                "expected column width ~20 for A1, got {v}"
            ),
            other => panic!("expected numeric column width, got {other:?}"),
        }

        // Hidden columns report width=0.
        state
            .set_cell(
                &app_sheet_id,
                1,
                2,
                None,
                Some("=CELL(\"width\",B1)".to_string()),
            )
            .expect("set hidden width formula");
        let c2 = state.get_cell(&app_sheet_id, 1, 2).expect("read C2");
        assert_eq!(c2.value, CellScalar::Number(0.0));
    }

    #[test]
    fn add_sheet_unknown_after_sheet_id_appends() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let inserted = state
            .add_sheet(
                "Inserted".to_string(),
                None,
                Some("does-not-exist".to_string()),
                None,
            )
            .expect("add sheet succeeds");
        let inserted_id = inserted.id;

        let workbook = state.get_workbook().expect("workbook loaded");
        let ids = workbook
            .sheets
            .iter()
            .map(|sheet| sheet.id.clone())
            .collect::<Vec<_>>();

        assert_eq!(
            ids,
            vec![
                "Sheet1".to_string(),
                "Sheet2".to_string(),
                inserted_id
            ]
        );
    }

    #[test]
    fn add_sheet_after_last_sheet_appends() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let inserted = state
            .add_sheet(
                "Inserted".to_string(),
                None,
                Some("sheet2".to_string()),
                None,
            )
            .expect("add sheet succeeds");
        let inserted_id = inserted.id;

        let workbook = state.get_workbook().expect("workbook loaded");
        let ids = workbook
            .sheets
            .iter()
            .map(|sheet| sheet.id.clone())
            .collect::<Vec<_>>();

        assert_eq!(
            ids,
            vec![
                "Sheet1".to_string(),
                "Sheet2".to_string(),
                inserted_id
            ]
        );
    }

    #[test]
    fn cell_filename_includes_workbook_path_when_set() {
        let workbook_path = std::env::temp_dir().join("formula-cell-filename-test.xlsx");
        let mut workbook = Workbook::new_empty(Some(workbook_path.to_string_lossy().to_string()));
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();
        workbook
            .sheet_mut(&sheet_id)
            .expect("Sheet1 exists")
            .set_cell(
                0,
                0,
                Cell::from_formula("=CELL(\"filename\")".to_string()),
            );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let cell = state.get_cell(&sheet_id, 0, 0).expect("read Sheet1!A1");

        let parent = workbook_path.parent().expect("workbook path has parent");
        let mut dir = parent.to_string_lossy().to_string();
        if dir
            .chars()
            .last()
            .is_some_and(|c| c != std::path::MAIN_SEPARATOR)
        {
            dir.push(std::path::MAIN_SEPARATOR);
        }
        let filename = workbook_path
            .file_name()
            .expect("workbook path has filename")
            .to_string_lossy();
        let expected = format!("{dir}[{filename}]Sheet1");
        assert_eq!(cell.value, CellScalar::Text(expected));
    }

    #[test]
    fn add_sheet_rebuilds_engine_order_for_3d_references() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        let sheet1_id = workbook.sheets[0].id.clone();
        let sheet2_id = workbook.sheets[1].id.clone();
        let sheet3_id = workbook.sheets[2].id.clone();

        workbook
            .sheet_mut(&sheet1_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(1.0))));
        workbook
            .sheet_mut(&sheet2_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(2.0))));
        workbook
            .sheet_mut(&sheet3_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(3.0))));

        // A 3D reference should include any sheets inserted between the two endpoints.
        workbook
            .sheet_mut(&sheet1_id)
            .unwrap()
            .set_cell(
                0,
                1,
                Cell::from_formula("=SUM(Sheet1:Sheet3!A1)".to_string()),
            );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let b1_before = state.get_cell(&sheet1_id, 0, 1).expect("read Sheet1!B1");
        assert_eq!(b1_before.value, CellScalar::Number(6.0));

        let inserted = state
            .add_sheet(
                "Inserted".to_string(),
                None,
                Some(sheet1_id.clone()),
                None,
            )
            .expect("add sheet succeeds");
        let inserted_id = inserted.id;

        state
            .set_cell(&inserted_id, 0, 0, Some(JsonValue::from(10)), None)
            .expect("set inserted A1");

        let b1_after = state.get_cell(&sheet1_id, 0, 1).expect("read Sheet1!B1 after insert");
        assert_eq!(b1_after.value, CellScalar::Number(16.0));
    }

    #[test]
    fn add_sheet_with_id_rebuilds_engine_order_for_3d_references() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        let sheet1_id = workbook.sheets[0].id.clone();
        let sheet2_id = workbook.sheets[1].id.clone();
        let sheet3_id = workbook.sheets[2].id.clone();

        workbook
            .sheet_mut(&sheet1_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(1.0))));
        workbook
            .sheet_mut(&sheet2_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(2.0))));
        workbook
            .sheet_mut(&sheet3_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(3.0))));

        // A 3D reference should include any sheets inserted between the two endpoints.
        workbook
            .sheet_mut(&sheet1_id)
            .unwrap()
            .set_cell(
                0,
                1,
                Cell::from_formula("=SUM(Sheet1:Sheet3!A1)".to_string()),
            );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let b1_before = state.get_cell(&sheet1_id, 0, 1).expect("read Sheet1!B1");
        assert_eq!(b1_before.value, CellScalar::Number(6.0));

        let inserted_id = "InsertedSheet".to_string();
        state
            .add_sheet_with_id(
                inserted_id.clone(),
                "Inserted".to_string(),
                Some(sheet1_id.clone()),
                None,
            )
            .expect("add sheet succeeds");

        state
            .set_cell(&inserted_id, 0, 0, Some(JsonValue::from(10)), None)
            .expect("set inserted A1");

        let b1_after = state.get_cell(&sheet1_id, 0, 1).expect("read Sheet1!B1 after insert");
        assert_eq!(b1_after.value, CellScalar::Number(16.0));
    }

    #[test]
    fn add_sheet_inserting_before_existing_sheets_shifts_macro_runtime_context() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .set_macro_runtime_context(MacroRuntimeContext {
                active_sheet: 2,
                active_cell: (1, 1),
                selection: Some(formula_vba_runtime::VbaRangeRef {
                    sheet: 1,
                    start_row: 1,
                    start_col: 1,
                    end_row: 1,
                    end_col: 1,
                }),
            })
            .expect("set macro runtime context");

        state
            .add_sheet("Inserted".to_string(), None, None, Some(0))
            .expect("add sheet succeeds");

        let ctx = state.macro_runtime_context();
        assert_eq!(ctx.active_sheet, 3);
        assert_eq!(ctx.selection.expect("selection").sheet, 2);
    }

    #[test]
    fn add_sheet_with_id_inserting_before_existing_sheets_shifts_macro_runtime_context() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .set_macro_runtime_context(MacroRuntimeContext {
                active_sheet: 2,
                active_cell: (1, 1),
                selection: Some(formula_vba_runtime::VbaRangeRef {
                    sheet: 1,
                    start_row: 1,
                    start_col: 1,
                    end_row: 1,
                    end_col: 1,
                }),
            })
            .expect("set macro runtime context");

        state
            .add_sheet_with_id(
                "InsertedSheet".to_string(),
                "Inserted".to_string(),
                None,
                Some(0),
            )
            .expect("add sheet succeeds");

        let ctx = state.macro_runtime_context();
        assert_eq!(ctx.active_sheet, 3);
        assert_eq!(ctx.selection.expect("selection").sheet, 2);
    }

    #[test]
    fn add_sheet_shifts_macro_runtime_context_before_macro_host_has_synced_vba_hash() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());
        workbook.vba_project_bin = Some(vec![1, 2, 3, 4]);

        let mut state = AppState::new();
        state.load_workbook(workbook);

        // Simulate a macro UI context that was set before the macro host has a chance to compute
        // the workbook's VBA project hash (i.e. before the first `sync_with_workbook`).
        state.macro_host.set_runtime_context(MacroRuntimeContext {
            active_sheet: 2,
            active_cell: (1, 1),
            selection: Some(formula_vba_runtime::VbaRangeRef {
                sheet: 1,
                start_row: 1,
                start_col: 1,
                end_row: 1,
                end_col: 1,
            }),
        });

        state
            .add_sheet("Inserted".to_string(), None, None, Some(0))
            .expect("add sheet succeeds");

        // `add_sheet` should keep the adjusted context stable even after the macro host later
        // synchronizes with the workbook's VBA project hash.
        {
            let workbook = state.workbook.as_ref().expect("workbook loaded");
            let macro_host = &mut state.macro_host;
            macro_host.sync_with_workbook(workbook);
        }

        let ctx = state.macro_runtime_context();
        assert_eq!(ctx.active_sheet, 3);
        assert_eq!(ctx.selection.expect("selection").sheet, 2);
    }

    #[test]
    fn add_sheet_with_id_shifts_macro_runtime_context_before_macro_host_has_synced_vba_hash() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());
        workbook.vba_project_bin = Some(vec![1, 2, 3, 4]);

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state.macro_host.set_runtime_context(MacroRuntimeContext {
            active_sheet: 2,
            active_cell: (1, 1),
            selection: Some(formula_vba_runtime::VbaRangeRef {
                sheet: 1,
                start_row: 1,
                start_col: 1,
                end_row: 1,
                end_col: 1,
            }),
        });

        state
            .add_sheet_with_id(
                "InsertedSheet".to_string(),
                "Inserted".to_string(),
                None,
                Some(0),
            )
            .expect("add sheet succeeds");

        {
            let workbook = state.workbook.as_ref().expect("workbook loaded");
            let macro_host = &mut state.macro_host;
            macro_host.sync_with_workbook(workbook);
        }

        let ctx = state.macro_runtime_context();
        assert_eq!(ctx.active_sheet, 3);
        assert_eq!(ctx.selection.expect("selection").sheet, 2);
    }

    #[test]
    fn add_sheet_allows_reusing_stable_id_after_delete_with_new_name() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .rename_sheet("Sheet1", "Budget".to_string())
            .expect("rename Sheet1");
        state.delete_sheet("Sheet1").expect("delete Sheet1");

        let restored = state
            .add_sheet(
                "Budget".to_string(),
                Some("Sheet1".to_string()),
                None,
                Some(0),
            )
            .expect("re-add sheet with stable id");
        assert_eq!(restored.id, "Sheet1");
        assert_eq!(restored.name, "Budget");

        let info = state.workbook_info().expect("workbook info");
        let summary = info
            .sheets
            .iter()
            .map(|s| (s.id.clone(), s.name.clone()))
            .collect::<Vec<_>>();
        assert_eq!(
            summary,
            vec![
                ("Sheet1".to_string(), "Budget".to_string()),
                ("Sheet2".to_string(), "Sheet2".to_string())
            ]
        );
    }

    #[test]
    fn rename_sheet_updates_preserved_sheet_keys() {
        use std::collections::BTreeMap;

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());

        let pivot_tables = BTreeMap::from([
            (
                "Sheet1".to_string(),
                PreservedSheetPivotTables {
                    sheet_index: 0,
                    sheet_id: Some(1),
                    pivot_tables_xml: b"<pivotTables/>".to_vec(),
                    pivot_table_rels: vec![RelationshipStub {
                        rel_id: "rId1".to_string(),
                        target: "xl/pivotTables/pivotTable1.xml".to_string(),
                    }],
                },
            ),
            (
                "Sheet2".to_string(),
                PreservedSheetPivotTables {
                    sheet_index: 1,
                    sheet_id: Some(2),
                    pivot_tables_xml: b"<pivotTables/>".to_vec(),
                    pivot_table_rels: vec![RelationshipStub {
                        rel_id: "rId2".to_string(),
                        target: "xl/pivotTables/pivotTable2.xml".to_string(),
                    }],
                },
            ),
        ]);

        workbook.preserved_pivot_parts = Some(PreservedPivotParts {
            content_types_xml: b"<Types/>".to_vec(),
            parts: BTreeMap::new(),
            workbook_sheets: vec![
                formula_xlsx::pivots::preserve::PreservedWorkbookSheet {
                    name: "Sheet1".to_string(),
                    index: 0,
                },
                formula_xlsx::pivots::preserve::PreservedWorkbookSheet {
                    name: "Sheet2".to_string(),
                    index: 1,
                },
            ],
            workbook_pivot_caches: None,
            workbook_pivot_cache_rels: Vec::new(),
            workbook_slicer_caches: None,
            workbook_slicer_cache_rels: Vec::new(),
            workbook_timeline_caches: None,
            workbook_timeline_cache_rels: Vec::new(),
            sheet_pivot_tables: pivot_tables,
        });

        workbook.preserved_drawing_parts = Some(PreservedDrawingParts {
            content_types_xml: b"<Types/>".to_vec(),
            parts: BTreeMap::new(),
            sheet_drawings: BTreeMap::from([
                (
                    "Sheet1".to_string(),
                    PreservedSheetDrawings {
                        sheet_index: 0,
                        sheet_id: Some(1),
                        drawings: vec![SheetDrawingRelationship {
                            rel_id: "rIdDrawing1".to_string(),
                            target: "xl/drawings/drawing1.xml".to_string(),
                        }],
                    },
                ),
                (
                    "Sheet2".to_string(),
                    PreservedSheetDrawings {
                        sheet_index: 1,
                        sheet_id: Some(2),
                        drawings: vec![SheetDrawingRelationship {
                            rel_id: "rIdDrawing2".to_string(),
                            target: "xl/drawings/drawing2.xml".to_string(),
                        }],
                    },
                ),
            ]),
            sheet_pictures: BTreeMap::from([(
                "Sheet1".to_string(),
                PreservedSheetPicture {
                    sheet_index: 0,
                    sheet_id: Some(1),
                    picture_xml: b"<picture/>".to_vec(),
                    picture_rel: SheetRelationshipStub {
                        rel_id: "rIdPic1".to_string(),
                        target: "xl/media/image1.png".to_string(),
                    },
                },
            )]),
            sheet_ole_objects: BTreeMap::from([(
                "Sheet1".to_string(),
                PreservedSheetOleObjects {
                    sheet_index: 0,
                    sheet_id: Some(1),
                    ole_objects_xml: b"<oleObjects/>".to_vec(),
                    ole_object_rels: vec![SheetRelationshipStub {
                        rel_id: "rIdOle1".to_string(),
                        target: "xl/embeddings/oleObject1.bin".to_string(),
                    }],
                },
            )]),
            sheet_controls: BTreeMap::from([(
                "Sheet1".to_string(),
                PreservedSheetControls {
                    sheet_index: 0,
                    sheet_id: Some(1),
                    controls_xml: b"<controls/>".to_vec(),
                    control_rels: vec![SheetRelationshipStubWithType {
                        rel_id: "rIdCtl1".to_string(),
                        type_: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control".to_string(),
                        target: "xl/controls/control1.xml".to_string(),
                    }],
                },
            )]),
            sheet_drawing_hfs: BTreeMap::from([(
                "Sheet1".to_string(),
                PreservedSheetDrawingHF {
                    sheet_index: 0,
                    sheet_id: Some(1),
                    drawing_hf_xml: b"<drawingHF/>".to_vec(),
                    drawing_hf_rels: vec![SheetRelationshipStubWithType {
                        rel_id: "rIdHeader1".to_string(),
                        type_: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image".to_string(),
                        target: "xl/media/header1.png".to_string(),
                    }],
                },
            )]),
            chart_sheets: BTreeMap::from([(
                "Chart1".to_string(),
                PreservedChartSheet {
                    sheet_index: 2,
                    sheet_id: Some(3),
                    rel_id: "rIdChart1".to_string(),
                    rel_target: "chartsheets/sheet1.xml".to_string(),
                    state: None,
                    part_name: "xl/chartsheets/sheet1.xml".to_string(),
                },
            )]),
        });

        let mut state = AppState::new();
        state.load_workbook(workbook);
        state
            .rename_sheet("Sheet1", "Budget".to_string())
            .expect("rename sheet succeeds");

        let workbook = state.get_workbook().expect("workbook loaded");

        let pivots = workbook
            .preserved_pivot_parts
            .as_ref()
            .expect("pivot parts preserved");
        assert!(!pivots.sheet_pivot_tables.contains_key("Sheet1"));
        assert!(pivots.sheet_pivot_tables.contains_key("Budget"));
        assert!(pivots.sheet_pivot_tables.contains_key("Sheet2"));
        let sheet1 = pivots
            .workbook_sheets
            .iter()
            .find(|s| s.name == "Sheet1")
            .expect("original sheet list should keep old names for rename mapping");
        assert_eq!(sheet1.index, 0);
        assert!(
            !pivots.workbook_sheets.iter().any(|s| s.name == "Budget"),
            "workbook_sheets should retain the original sheet names"
        );

        let drawings = workbook
            .preserved_drawing_parts
            .as_ref()
            .expect("drawing parts preserved");
        assert!(!drawings.sheet_drawings.contains_key("Sheet1"));
        assert!(drawings.sheet_drawings.contains_key("Budget"));
        assert!(!drawings.sheet_pictures.contains_key("Sheet1"));
        assert!(drawings.sheet_pictures.contains_key("Budget"));
        assert!(!drawings.sheet_ole_objects.contains_key("Sheet1"));
        assert!(drawings.sheet_ole_objects.contains_key("Budget"));
        assert!(!drawings.sheet_controls.contains_key("Sheet1"));
        assert!(drawings.sheet_controls.contains_key("Budget"));
        assert!(!drawings.sheet_drawing_hfs.contains_key("Sheet1"));
        assert!(drawings.sheet_drawing_hfs.contains_key("Budget"));
        assert!(drawings.chart_sheets.contains_key("Chart1"));
    }

    #[test]
    fn delete_sheet_removes_preserved_parts_for_deleted_sheet() {
        use std::collections::BTreeMap;

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());

        workbook.preserved_pivot_parts = Some(PreservedPivotParts {
            content_types_xml: b"<Types/>".to_vec(),
            parts: BTreeMap::new(),
            workbook_sheets: vec![
                formula_xlsx::pivots::preserve::PreservedWorkbookSheet {
                    name: "Sheet1".to_string(),
                    index: 0,
                },
                formula_xlsx::pivots::preserve::PreservedWorkbookSheet {
                    name: "Sheet2".to_string(),
                    index: 1,
                },
            ],
            workbook_pivot_caches: None,
            workbook_pivot_cache_rels: Vec::new(),
            workbook_slicer_caches: None,
            workbook_slicer_cache_rels: Vec::new(),
            workbook_timeline_caches: None,
            workbook_timeline_cache_rels: Vec::new(),
            sheet_pivot_tables: BTreeMap::from([
                (
                    "Sheet1".to_string(),
                    PreservedSheetPivotTables {
                        sheet_index: 0,
                        sheet_id: Some(1),
                        pivot_tables_xml: b"<pivotTables/>".to_vec(),
                        pivot_table_rels: vec![RelationshipStub {
                            rel_id: "rId1".to_string(),
                            target: "xl/pivotTables/pivotTable1.xml".to_string(),
                        }],
                    },
                ),
                (
                    "Sheet2".to_string(),
                    PreservedSheetPivotTables {
                        sheet_index: 1,
                        sheet_id: Some(2),
                        pivot_tables_xml: b"<pivotTables/>".to_vec(),
                        pivot_table_rels: vec![RelationshipStub {
                            rel_id: "rId2".to_string(),
                            target: "xl/pivotTables/pivotTable2.xml".to_string(),
                        }],
                    },
                ),
            ]),
        });

        workbook.preserved_drawing_parts = Some(PreservedDrawingParts {
            content_types_xml: b"<Types/>".to_vec(),
            parts: BTreeMap::new(),
            sheet_drawings: BTreeMap::from([
                (
                    "Sheet1".to_string(),
                    PreservedSheetDrawings {
                        sheet_index: 0,
                        sheet_id: Some(1),
                        drawings: vec![SheetDrawingRelationship {
                            rel_id: "rIdDrawing1".to_string(),
                            target: "xl/drawings/drawing1.xml".to_string(),
                        }],
                    },
                ),
                (
                    "Sheet2".to_string(),
                    PreservedSheetDrawings {
                        sheet_index: 1,
                        sheet_id: Some(2),
                        drawings: vec![SheetDrawingRelationship {
                            rel_id: "rIdDrawing2".to_string(),
                            target: "xl/drawings/drawing2.xml".to_string(),
                        }],
                    },
                ),
            ]),
            sheet_pictures: BTreeMap::from([(
                "Sheet1".to_string(),
                PreservedSheetPicture {
                    sheet_index: 0,
                    sheet_id: Some(1),
                    picture_xml: b"<picture/>".to_vec(),
                    picture_rel: SheetRelationshipStub {
                        rel_id: "rIdPic1".to_string(),
                        target: "xl/media/image1.png".to_string(),
                    },
                },
            )]),
            sheet_ole_objects: BTreeMap::from([(
                "Sheet1".to_string(),
                PreservedSheetOleObjects {
                    sheet_index: 0,
                    sheet_id: Some(1),
                    ole_objects_xml: b"<oleObjects/>".to_vec(),
                    ole_object_rels: vec![SheetRelationshipStub {
                        rel_id: "rIdOle1".to_string(),
                        target: "xl/embeddings/oleObject1.bin".to_string(),
                    }],
                },
            )]),
            sheet_controls: BTreeMap::from([(
                "Sheet1".to_string(),
                PreservedSheetControls {
                    sheet_index: 0,
                    sheet_id: Some(1),
                    controls_xml: b"<controls/>".to_vec(),
                    control_rels: vec![SheetRelationshipStubWithType {
                        rel_id: "rIdCtl1".to_string(),
                        type_: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control".to_string(),
                        target: "xl/controls/control1.xml".to_string(),
                    }],
                },
            )]),
            sheet_drawing_hfs: BTreeMap::from([(
                "Sheet1".to_string(),
                PreservedSheetDrawingHF {
                    sheet_index: 0,
                    sheet_id: Some(1),
                    drawing_hf_xml: b"<drawingHF/>".to_vec(),
                    drawing_hf_rels: vec![SheetRelationshipStubWithType {
                        rel_id: "rIdHeader1".to_string(),
                        type_: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image".to_string(),
                        target: "xl/media/header1.png".to_string(),
                    }],
                },
            )]),
            chart_sheets: BTreeMap::from([(
                "Chart1".to_string(),
                PreservedChartSheet {
                    sheet_index: 2,
                    sheet_id: Some(3),
                    rel_id: "rIdChart1".to_string(),
                    rel_target: "chartsheets/sheet1.xml".to_string(),
                    state: None,
                    part_name: "xl/chartsheets/sheet1.xml".to_string(),
                },
            )]),
        });

        let mut state = AppState::new();
        state.load_workbook(workbook);
        state.delete_sheet("Sheet1").expect("delete sheet succeeds");

        let workbook = state.get_workbook().expect("workbook loaded");

        let pivots = workbook
            .preserved_pivot_parts
            .as_ref()
            .expect("pivot parts preserved");
        assert!(!pivots.sheet_pivot_tables.contains_key("Sheet1"));
        assert!(pivots.sheet_pivot_tables.contains_key("Sheet2"));
        assert!(
            !pivots.workbook_sheets.iter().any(|s| s.name == "Sheet1"),
            "deleted sheet should be removed from workbook_sheets"
        );
        let sheet2 = pivots
            .workbook_sheets
            .iter()
            .find(|s| s.name == "Sheet2")
            .expect("Sheet2 should remain in workbook_sheets");
        assert_eq!(sheet2.index, 0, "Sheet2 index should shift after delete");

        let drawings = workbook
            .preserved_drawing_parts
            .as_ref()
            .expect("drawing parts preserved");
        assert!(!drawings.sheet_drawings.contains_key("Sheet1"));
        assert!(drawings.sheet_drawings.contains_key("Sheet2"));
        assert!(!drawings.sheet_pictures.contains_key("Sheet1"));
        assert!(!drawings.sheet_ole_objects.contains_key("Sheet1"));
        assert!(!drawings.sheet_controls.contains_key("Sheet1"));
        assert!(!drawings.sheet_drawing_hfs.contains_key("Sheet1"));
        // Chart sheets are not deleted by the worksheet delete operation.
        assert!(drawings.chart_sheets.contains_key("Chart1"));
    }

    #[test]
    fn add_sheet_shifts_preserved_pivot_workbook_sheet_indices() {
        use std::collections::BTreeMap;

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());

        workbook.preserved_pivot_parts = Some(PreservedPivotParts {
            content_types_xml: b"<Types/>".to_vec(),
            parts: BTreeMap::new(),
            workbook_sheets: vec![
                formula_xlsx::pivots::preserve::PreservedWorkbookSheet {
                    name: "Sheet1".to_string(),
                    index: 0,
                },
                formula_xlsx::pivots::preserve::PreservedWorkbookSheet {
                    name: "Sheet2".to_string(),
                    index: 1,
                },
            ],
            workbook_pivot_caches: None,
            workbook_pivot_cache_rels: Vec::new(),
            workbook_slicer_caches: None,
            workbook_slicer_cache_rels: Vec::new(),
            workbook_timeline_caches: None,
            workbook_timeline_cache_rels: Vec::new(),
            sheet_pivot_tables: BTreeMap::new(),
        });

        let mut state = AppState::new();
        state.load_workbook(workbook);
        state
            .add_sheet("Inserted".to_string(), None, None, Some(0))
            .expect("insert sheet succeeds");

        let workbook = state.get_workbook().expect("workbook loaded");
        let pivots = workbook
            .preserved_pivot_parts
            .as_ref()
            .expect("pivot parts preserved");
        assert_eq!(pivots.workbook_sheets.len(), 2);

        let sheet1 = pivots
            .workbook_sheets
            .iter()
            .find(|s| s.name == "Sheet1")
            .expect("Sheet1 present");
        let sheet2 = pivots
            .workbook_sheets
            .iter()
            .find(|s| s.name == "Sheet2")
            .expect("Sheet2 present");
        assert_eq!(sheet1.index, 1);
        assert_eq!(sheet2.index, 2);
    }

    #[test]
    fn reorder_sheets_shifts_preserved_pivot_workbook_sheet_indices() {
        use std::collections::BTreeMap;

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        workbook.preserved_pivot_parts = Some(PreservedPivotParts {
            content_types_xml: b"<Types/>".to_vec(),
            parts: BTreeMap::new(),
            workbook_sheets: vec![
                formula_xlsx::pivots::preserve::PreservedWorkbookSheet {
                    name: "Sheet1".to_string(),
                    index: 0,
                },
                formula_xlsx::pivots::preserve::PreservedWorkbookSheet {
                    name: "Sheet2".to_string(),
                    index: 1,
                },
                formula_xlsx::pivots::preserve::PreservedWorkbookSheet {
                    name: "Sheet3".to_string(),
                    index: 2,
                },
            ],
            workbook_pivot_caches: None,
            workbook_pivot_cache_rels: Vec::new(),
            workbook_slicer_caches: None,
            workbook_slicer_cache_rels: Vec::new(),
            workbook_timeline_caches: None,
            workbook_timeline_cache_rels: Vec::new(),
            sheet_pivot_tables: BTreeMap::new(),
        });

        let mut state = AppState::new();
        state.load_workbook(workbook);
        state
            .reorder_sheets(vec![
                "Sheet3".to_string(),
                "Sheet1".to_string(),
                "Sheet2".to_string(),
            ])
            .expect("reorder sheets succeeds");

        let workbook = state.get_workbook().expect("workbook loaded");
        let pivots = workbook
            .preserved_pivot_parts
            .as_ref()
            .expect("pivot parts preserved");

        let sheet1 = pivots
            .workbook_sheets
            .iter()
            .find(|s| s.name == "Sheet1")
            .expect("Sheet1 present");
        let sheet2 = pivots
            .workbook_sheets
            .iter()
            .find(|s| s.name == "Sheet2")
            .expect("Sheet2 present");
        let sheet3 = pivots
            .workbook_sheets
            .iter()
            .find(|s| s.name == "Sheet3")
            .expect("Sheet3 present");
        assert_eq!(sheet1.index, 1);
        assert_eq!(sheet2.index, 2);
        assert_eq!(sheet3.index, 0);
    }

    #[test]
    fn move_sheet_shifts_preserved_pivot_workbook_sheet_indices() {
        use std::collections::BTreeMap;

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        workbook.preserved_pivot_parts = Some(PreservedPivotParts {
            content_types_xml: b"<Types/>".to_vec(),
            parts: BTreeMap::new(),
            workbook_sheets: vec![
                formula_xlsx::pivots::preserve::PreservedWorkbookSheet {
                    name: "Sheet1".to_string(),
                    index: 0,
                },
                formula_xlsx::pivots::preserve::PreservedWorkbookSheet {
                    name: "Sheet2".to_string(),
                    index: 1,
                },
                formula_xlsx::pivots::preserve::PreservedWorkbookSheet {
                    name: "Sheet3".to_string(),
                    index: 2,
                },
            ],
            workbook_pivot_caches: None,
            workbook_pivot_cache_rels: Vec::new(),
            workbook_slicer_caches: None,
            workbook_slicer_cache_rels: Vec::new(),
            workbook_timeline_caches: None,
            workbook_timeline_cache_rels: Vec::new(),
            sheet_pivot_tables: BTreeMap::new(),
        });

        let mut state = AppState::new();
        state.load_workbook(workbook);
        state
            .move_sheet("Sheet1", 2)
            .expect("move sheet succeeds");

        let workbook = state.get_workbook().expect("workbook loaded");
        let pivots = workbook
            .preserved_pivot_parts
            .as_ref()
            .expect("pivot parts preserved");

        let sheet1 = pivots
            .workbook_sheets
            .iter()
            .find(|s| s.name == "Sheet1")
            .expect("Sheet1 present");
        let sheet2 = pivots
            .workbook_sheets
            .iter()
            .find(|s| s.name == "Sheet2")
            .expect("Sheet2 present");
        let sheet3 = pivots
            .workbook_sheets
            .iter()
            .find(|s| s.name == "Sheet3")
            .expect("Sheet3 present");

        assert_eq!(sheet1.index, 2);
        assert_eq!(sheet2.index, 0);
        assert_eq!(sheet3.index, 1);
    }

    #[test]
    fn add_sheet_with_explicit_id_rejects_empty_id() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let err = state
            .add_sheet(
                "Inserted".to_string(),
                Some("   ".to_string()),
                None,
                None,
            )
            .expect_err("expected empty sheet id to be rejected");
        match err {
            AppStateError::WhatIf(message) => {
                assert!(message.contains("sheet id"), "unexpected error: {message}");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn add_sheet_with_explicit_id_rejects_duplicate_id_case_insensitive() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let err = state
            .add_sheet(
                "Inserted".to_string(),
                Some("sHeEt1".to_string()),
                None,
                None,
            )
            .expect_err("expected duplicate sheet id to be rejected");
        match err {
            AppStateError::WhatIf(message) => {
                assert!(message.contains("duplicate sheet id"), "unexpected error: {message}");
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn add_sheet_with_explicit_id_rejects_duplicate_name_case_insensitive() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let err = state
            .add_sheet(
                "sHeEt1".to_string(),
                Some("UniqueId".to_string()),
                None,
                None,
            )
            .expect_err("expected duplicate sheet name to be rejected");
        match err {
            AppStateError::WhatIf(message) => {
                assert!(
                    message.contains("sheet name already exists"),
                    "unexpected error: {message}"
                );
            }
            other => panic!("unexpected error: {other:?}"),
        }
    }

    #[test]
    fn set_sheet_visibility_rejects_hiding_last_visible_sheet_and_preserves_state() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());

        // Ensure there is exactly one visible sheet.
        workbook.sheets[1].visibility = SheetVisibility::Hidden;

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let err = state
            .set_sheet_visibility("Sheet1", SheetVisibility::Hidden)
            .expect_err("expected hiding the last visible sheet to fail");

        match err {
            AppStateError::WhatIf(msg) => assert!(
                msg.contains("cannot hide the last visible sheet"),
                "unexpected error message: {msg}"
            ),
            other => panic!("expected AppStateError::WhatIf, got {other:?}"),
        }

        // Ensure the workbook was not modified after a failed update.
        let info = state.workbook_info().expect("workbook info");
        let sheet1 = info
            .sheets
            .iter()
            .find(|s| s.id == "Sheet1")
            .expect("Sheet1 exists");
        assert_eq!(sheet1.visibility, SheetVisibility::Visible);
    }

    #[test]
    fn set_sheet_visibility_allows_very_hidden_and_persists_in_storage() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());

        let mut state = AppState::new();
        state
            .load_workbook_persistent(workbook, WorkbookPersistenceLocation::InMemory)
            .expect("load persistent workbook");

        state
            .set_sheet_visibility("Sheet2", SheetVisibility::VeryHidden)
            .expect("set Sheet2 visibility");

        let info = state.workbook_info().expect("workbook info");
        let sheet2 = info
            .sheets
            .iter()
            .find(|s| s.id == "Sheet2")
            .expect("Sheet2 exists");
        assert_eq!(sheet2.visibility, SheetVisibility::VeryHidden);

        let storage = state.persistent_storage().expect("storage");
        let workbook_id = state.persistent_workbook_id().expect("workbook id");
        let sheets = storage.list_sheets(workbook_id).expect("list sheets");
        let sheet2 = sheets.iter().find(|s| s.name == "Sheet2").expect("Sheet2");
        assert_eq!(
            sheet2.visibility,
            formula_storage::SheetVisibility::VeryHidden
        );
    }

    #[test]
    fn set_sheet_tab_color_accepts_theme_updates_and_persists() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        let mut state = AppState::new();
        state
            .load_workbook_persistent(workbook, WorkbookPersistenceLocation::InMemory)
            .expect("load persistent workbook");

        state
            .set_sheet_tab_color(
                "Sheet1",
                Some(TabColor {
                    theme: Some(1),
                    tint: Some(0.5),
                    ..Default::default()
                }),
            )
            .expect("set theme tab color");

        // Ensure update mutates in-memory state.
        let info = state.workbook_info().expect("workbook info");
        let sheet1 = info
            .sheets
            .iter()
            .find(|s| s.id == "Sheet1")
            .expect("Sheet1 exists");
        assert_eq!(
            sheet1.tab_color,
            Some(TabColor {
                theme: Some(1),
                tint: Some(0.5),
                ..Default::default()
            })
        );

        // Ensure update writes to persistent storage.
        let storage = state.persistent_storage().expect("storage");
        let workbook_id = state.persistent_workbook_id().expect("workbook id");
        let model = storage
            .export_model_workbook(workbook_id)
            .expect("export model");
        let sheet1 = model.sheet_by_name("Sheet1").expect("Sheet1 exists");
        assert_eq!(
            sheet1.tab_color,
            Some(TabColor {
                theme: Some(1),
                tint: Some(0.5),
                ..Default::default()
            })
        );
    }

    #[test]
    fn sheet_visibility_and_tab_color_survive_persistent_recovery() {
        let mut model = formula_model::Workbook::new();
        model.add_sheet("Sheet1").expect("add sheet");
        model.add_sheet("Sheet2").expect("add sheet");

        let mut buf = std::io::Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&model, &mut buf).expect("write xlsx bytes");

        let tmp = tempfile::tempdir().expect("temp dir");
        let xlsx_path = tmp.path().join("sheet-meta.xlsx");
        std::fs::write(&xlsx_path, buf.into_inner()).expect("write xlsx file");

        let db_path = tmp.path().join("autosave.sqlite");
        let location = WorkbookPersistenceLocation::OnDisk(db_path);

        {
            let workbook = read_xlsx_blocking(&xlsx_path).expect("read xlsx workbook");
            let mut state = AppState::new();
            state
                .load_workbook_persistent(workbook, location.clone())
                .expect("load persistent workbook");

            state
                .set_sheet_visibility("Sheet2", SheetVisibility::Hidden)
                .expect("hide Sheet2");
            state
                .set_sheet_tab_color(
                    "Sheet1",
                    Some(TabColor {
                        rgb: Some("FFFF0000".to_string()),
                        ..Default::default()
                    }),
                )
                .expect("set Sheet1 tab color");

            let info = state.workbook_info().expect("workbook info");
            let sheet2 = info
                .sheets
                .iter()
                .find(|s| s.id == "Sheet2")
                .expect("Sheet2 exists");
            assert_eq!(sheet2.visibility, SheetVisibility::Hidden);
            let sheet1 = info
                .sheets
                .iter()
                .find(|s| s.id == "Sheet1")
                .expect("Sheet1 exists");
            assert_eq!(
                sheet1.tab_color.as_ref().and_then(|c| c.rgb.as_deref()),
                Some("FFFF0000")
            );
        }

        let workbook = read_xlsx_blocking(&xlsx_path).expect("re-read xlsx workbook");
        let mut state = AppState::new();
        state
            .load_workbook_persistent(workbook, location)
            .expect("recover workbook from autosave storage");

        let info = state.workbook_info().expect("workbook info");
        let sheet2 = info
            .sheets
            .iter()
            .find(|s| s.id == "Sheet2")
            .expect("Sheet2 exists");
        assert_eq!(sheet2.visibility, SheetVisibility::Hidden);
        let sheet1 = info
            .sheets
            .iter()
            .find(|s| s.id == "Sheet1")
            .expect("Sheet1 exists");
        assert_eq!(
            sheet1.tab_color.as_ref().and_then(|c| c.rgb.as_deref()),
            Some("FFFF0000")
        );
    }

    #[test]
    fn xlsb_provenance_survives_persistent_recovery() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsb/tests/fixtures/simple.xlsb"
        ));
 
        let tmp = tempfile::tempdir().expect("temp dir");
        let db_path = tmp.path().join("autosave.sqlite");
        let location = WorkbookPersistenceLocation::OnDisk(db_path);
 
        {
            let workbook = read_xlsx_blocking(fixture_path).expect("read xlsb workbook");
            let mut state = AppState::new();
            state
                .load_workbook_persistent(workbook, location.clone())
                .expect("load persistent workbook");
        }
 
        let workbook = read_xlsx_blocking(fixture_path).expect("re-read xlsb workbook");
        let mut state = AppState::new();
        state
            .load_workbook_persistent(workbook, location)
            .expect("recover workbook from autosave storage");
 
        let recovered = state.get_workbook().expect("workbook");
        assert_eq!(
            recovered.origin_xlsb_path.as_deref(),
            Some(fixture_path.to_string_lossy().as_ref())
        );
        assert_eq!(recovered.sheets[0].origin_ordinal, Some(0));
    }

    #[test]
    fn xlsx_style_only_number_formats_survive_persistent_recovery() {
        use formula_model::{Cell as ModelCell, CellRef, CellValue as ModelCellValue, Style};

        let mut model = formula_model::Workbook::new();
        let sheet_id = model.add_sheet("Sheet1").expect("add sheet");
        let style_id = model.intern_style(Style {
            number_format: Some("m/d/yyyy".to_string()),
            ..Default::default()
        });
        let sheet = model.sheet_mut(sheet_id).expect("sheet exists");
        let mut model_cell = ModelCell::new(ModelCellValue::Empty);
        model_cell.style_id = style_id;
        sheet.set_cell(CellRef::new(0, 0), model_cell);

        let mut buf = std::io::Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&model, &mut buf).expect("write xlsx bytes");

        let tmp = tempfile::tempdir().expect("temp dir");
        let xlsx_path = tmp.path().join("style-only.xlsx");
        std::fs::write(&xlsx_path, buf.into_inner()).expect("write xlsx file");

        let db_path = tmp.path().join("autosave.sqlite");
        let location = WorkbookPersistenceLocation::OnDisk(db_path);

        {
            let workbook = read_xlsx_blocking(&xlsx_path).expect("read xlsx workbook");
            let mut state = AppState::new();
            state
                .load_workbook_persistent(workbook, location.clone())
                .expect("load persistent workbook");
        }

        let workbook = read_xlsx_blocking(&xlsx_path).expect("re-read xlsx workbook");
        let mut state = AppState::new();
        state
            .load_workbook_persistent(workbook, location)
            .expect("recover workbook from autosave storage");

        let recovered = state.get_workbook().expect("workbook");
        let cell = recovered.sheets[0].get_cell(0, 0);
        assert_eq!(cell.computed_value, CellScalar::Empty);
        assert_eq!(cell.number_format.as_deref(), Some("m/d/yyyy"));
    }

    #[test]
    fn xlsx_style_only_cells_seed_sheet_formatting_metadata_on_first_import() {
        use formula_model::{Cell as ModelCell, CellRef, CellValue as ModelCellValue, Font, Style};
        use std::io::Cursor;

        let mut model = formula_model::Workbook::new();
        let sheet_id = model.add_sheet("Sheet1").expect("add sheet");
        let style_id = model.intern_style(Style {
            font: Some(Font {
                bold: true,
                ..Default::default()
            }),
            ..Default::default()
        });

        let sheet = model.sheet_mut(sheet_id).expect("sheet exists");
        let mut model_cell = ModelCell::new(ModelCellValue::Empty);
        model_cell.style_id = style_id;
        sheet.set_cell(CellRef::new(0, 0), model_cell);

        let mut buf = Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&model, &mut buf).expect("write xlsx bytes");
        let bytes = buf.into_inner();

        let tmp = tempfile::tempdir().expect("temp dir");
        let xlsx_path = tmp.path().join("formatting-seed.xlsx");
        std::fs::write(&xlsx_path, &bytes).expect("write xlsx file");

        let workbook = read_xlsx_blocking(&xlsx_path).expect("read xlsx workbook");
        let mut state = AppState::new();
        state
            .load_workbook_persistent(workbook, WorkbookPersistenceLocation::InMemory)
            .expect("load persistent workbook");

        let persistent = state.persistent.as_ref().expect("persistent workbook");
        let sheets = persistent
            .storage
            .list_sheets(persistent.workbook_id)
            .expect("list sheets");
        assert_eq!(sheets.len(), 1, "expected one sheet in persisted storage");

        let sheet_meta = persistent
            .storage
            .get_sheet_meta(sheets[0].id)
            .expect("get sheet meta");
        let metadata = sheet_meta.metadata.expect("expected sheet metadata to be seeded");
        let formatting = metadata
            .get(FORMULA_UI_FORMATTING_METADATA_KEY)
            .expect("expected formatting metadata key");

        let cell_formats = formatting
            .get("cellFormats")
            .and_then(|v| v.as_array())
            .expect("cellFormats array");
        assert_eq!(cell_formats.len(), 1);

        let entry = &cell_formats[0];
        assert_eq!(entry.get("row").and_then(|v| v.as_u64()), Some(0));
        assert_eq!(entry.get("col").and_then(|v| v.as_u64()), Some(0));

        // Compare against the style payload as read from the XLSX bytes (so the test is resilient
        // to XLSX round-trip normalization).
        let reread_model =
            formula_xlsx::read_workbook_from_reader(Cursor::new(bytes.as_slice())).expect("read model");
        let reread_sheet = reread_model
            .sheets
            .iter()
            .find(|s| s.name == "Sheet1")
            .expect("Sheet1 exists");
        let reread_cell = reread_sheet
            .cell(CellRef::new(0, 0))
            .expect("A1 exists in reread model");
        let expected_style = reread_model
            .styles
            .get(reread_cell.style_id)
            .expect("style exists");
        let expected_format =
            serde_json::to_value(expected_style).expect("serialize expected style");

        assert_eq!(
            entry.get("format"),
            Some(&expected_format),
            "expected seeded cell format to match the XLSX style payload"
        );
    }

    #[test]
    fn sheet_formatting_deltas_update_engine_and_survive_persistent_reload_for_cell_protect() {
        use crate::commands::{
            apply_sheet_formatting_deltas_inner, ApplySheetFormattingDeltasRequest,
            LimitedSheetCellFormatDeltas, SheetCellFormatDelta,
        };
        use formula_model::Protection;

        let tmp = tempfile::tempdir().expect("temp dir");
        let db_path = tmp.path().join("autosave.sqlite");
        let location = WorkbookPersistenceLocation::OnDisk(db_path);

        {
            let mut workbook = Workbook::new_empty(None);
            workbook.add_sheet("Sheet1".to_string());

            let mut state = AppState::new();
            state
                .load_workbook_persistent(workbook, location.clone())
                .expect("load persistent workbook");

            // B1 = CELL("protect", A1) should report the default locked state.
            state
                .set_cell("Sheet1", 0, 1, None, Some("=CELL(\"protect\",A1)".to_string()))
                .expect("set formula");
            let before = state.get_cell("Sheet1", 0, 1).expect("get B1");
            assert_eq!(before.value, CellScalar::Number(1.0));

            // Apply a formatting delta that unlocks A1 (protection.locked = false).
            let format = serde_json::to_value(formula_model::Style {
                protection: Some(Protection {
                    locked: false,
                    hidden: false,
                }),
                ..Default::default()
            })
            .expect("serialize style");
            apply_sheet_formatting_deltas_inner(
                &mut state,
                ApplySheetFormattingDeltasRequest {
                    sheet_id: "Sheet1".to_string(),
                    default_format: None,
                    row_formats: None,
                    col_formats: None,
                    format_runs_by_col: None,
                    cell_formats: Some(LimitedSheetCellFormatDeltas(vec![SheetCellFormatDelta {
                        row: 0,
                        col: 0,
                        format,
                    }])),
                },
            )
            .expect("apply formatting delta");

            // Recalculate the dirty set; B1 should update without a full engine rebuild.
            state.engine.recalculate_with_value_changes_multi_threaded();
            let after = state.get_cell("Sheet1", 0, 1).expect("get B1");
            assert_eq!(after.value, CellScalar::Number(0.0));
        }

        // Re-open from persistent storage; the unlocked format should be loaded into the engine.
        let mut recovered_state = AppState::new();
        recovered_state
            .load_workbook_persistent(Workbook::new_empty(None), location)
            .expect("recover workbook from storage");
        let recovered = recovered_state
            .get_cell("Sheet1", 0, 1)
            .expect("get B1 after reload");
        assert_eq!(recovered.value, CellScalar::Number(0.0));
    }

    #[test]
    fn sheet_default_formatting_survives_engine_rebuild_and_persistent_reload_for_cell_protect() {
        use crate::commands::{apply_sheet_formatting_deltas_inner, ApplySheetFormattingDeltasRequest};

        let tmp = tempfile::tempdir().expect("temp dir");
        let db_path = tmp.path().join("autosave.sqlite");
        let location = WorkbookPersistenceLocation::OnDisk(db_path);

        {
            let mut workbook = Workbook::new_empty(None);
            workbook.add_sheet("Sheet1".to_string());

            let mut state = AppState::new();
            state
                .load_workbook_persistent(workbook, location.clone())
                .expect("load persistent workbook");

            // B1 = CELL("protect", A1) should report the default locked state.
            state
                .set_cell("Sheet1", 0, 1, None, Some("=CELL(\"protect\",A1)".to_string()))
                .expect("set formula");
            let before = state.get_cell("Sheet1", 0, 1).expect("get B1");
            assert_eq!(before.value, CellScalar::Number(1.0));

            // Set the sheet default format to unlocked (protection.locked=false).
            let default_format = serde_json::json!({ "protection": { "locked": false } });
            apply_sheet_formatting_deltas_inner(
                &mut state,
                ApplySheetFormattingDeltasRequest {
                    sheet_id: "Sheet1".to_string(),
                    default_format: Some(Some(default_format)),
                    row_formats: None,
                    col_formats: None,
                    format_runs_by_col: None,
                    cell_formats: None,
                },
            )
            .expect("apply formatting delta");

            state.engine.recalculate_with_value_changes_multi_threaded();
            let after = state.get_cell("Sheet1", 0, 1).expect("get B1");
            assert_eq!(after.value, CellScalar::Number(0.0));

            // Rebuild the engine and ensure the default format is re-applied from persistence
            // metadata.
            state
                .rebuild_engine_from_workbook()
                .expect("rebuild engine from workbook");
            state.engine.recalculate_with_value_changes_multi_threaded();
            let after_rebuild = state.get_cell("Sheet1", 0, 1).expect("get B1 after rebuild");
            assert_eq!(after_rebuild.value, CellScalar::Number(0.0));
        }

        // Re-open from persistent storage; the unlocked default format should be loaded into the engine.
        let mut recovered_state = AppState::new();
        recovered_state
            .load_workbook_persistent(Workbook::new_empty(None), location)
            .expect("recover workbook from storage");
        let recovered = recovered_state
            .get_cell("Sheet1", 0, 1)
            .expect("get B1 after reload");
        assert_eq!(recovered.value, CellScalar::Number(0.0));
    }

    #[test]
    fn sheet_format_runs_by_col_survive_engine_rebuild_and_persistent_reload_for_cell_protect() {
        use crate::commands::{
            apply_sheet_formatting_deltas_inner, ApplySheetFormattingDeltasRequest,
            LimitedSheetFormatRunDeltas, LimitedSheetFormatRunsByColDeltas, SheetFormatRunDelta,
            SheetFormatRunsByColDelta,
        };

        let tmp = tempfile::tempdir().expect("temp dir");
        let db_path = tmp.path().join("autosave.sqlite");
        let location = WorkbookPersistenceLocation::OnDisk(db_path);

        {
            let mut workbook = Workbook::new_empty(None);
            workbook.add_sheet("Sheet1".to_string());

            let mut state = AppState::new();
            state
                .load_workbook_persistent(workbook, location.clone())
                .expect("load persistent workbook");

            state
                .set_cell("Sheet1", 0, 1, None, Some("=CELL(\"protect\",A1)".to_string()))
                .expect("set formula");
            let before = state.get_cell("Sheet1", 0, 1).expect("get B1");
            assert_eq!(before.value, CellScalar::Number(1.0));

            // Apply a run format that unlocks A1 (column A rows [0,1)).
            let run_format = serde_json::json!({ "protection": { "locked": false } });
            apply_sheet_formatting_deltas_inner(
                &mut state,
                ApplySheetFormattingDeltasRequest {
                    sheet_id: "Sheet1".to_string(),
                    default_format: None,
                    row_formats: None,
                    col_formats: None,
                    format_runs_by_col: Some(LimitedSheetFormatRunsByColDeltas(vec![
                        SheetFormatRunsByColDelta {
                            col: 0,
                            runs: LimitedSheetFormatRunDeltas(vec![SheetFormatRunDelta {
                                start_row: 0,
                                end_row_exclusive: 1,
                                format: run_format,
                            }]),
                        },
                    ])),
                    cell_formats: None,
                },
            )
            .expect("apply formatting delta");

            state.engine.recalculate_with_value_changes_multi_threaded();
            let after = state.get_cell("Sheet1", 0, 1).expect("get B1");
            assert_eq!(after.value, CellScalar::Number(0.0));

            // Rebuild and ensure format runs are re-applied from persistence metadata.
            state
                .rebuild_engine_from_workbook()
                .expect("rebuild engine from workbook");
            state.engine.recalculate_with_value_changes_multi_threaded();
            let after_rebuild = state.get_cell("Sheet1", 0, 1).expect("get B1 after rebuild");
            assert_eq!(after_rebuild.value, CellScalar::Number(0.0));
        }

        let mut recovered_state = AppState::new();
        recovered_state
            .load_workbook_persistent(Workbook::new_empty(None), location)
            .expect("recover workbook from storage");
        let recovered = recovered_state
            .get_cell("Sheet1", 0, 1)
            .expect("get B1 after reload");
        assert_eq!(recovered.value, CellScalar::Number(0.0));
    }

    #[test]
    fn ui_style_number_format_clear_semantics_survive_engine_rebuild_and_persistent_reload() {
        use crate::commands::{
            apply_sheet_formatting_deltas_inner, ApplySheetFormattingDeltasRequest,
            LimitedSheetCellFormatDeltas, SheetCellFormatDelta,
        };

        let tmp = tempfile::tempdir().expect("temp dir");
        let db_path = tmp.path().join("autosave.sqlite");
        let location = WorkbookPersistenceLocation::OnDisk(db_path);

        {
            let mut workbook = Workbook::new_empty(None);
            workbook.add_sheet("Sheet1".to_string());

            let mut state = AppState::new();
            state
                .load_workbook_persistent(workbook, location.clone())
                .expect("load persistent workbook");

            // B1 = CELL("format", A1) should report General by default.
            state
                .set_cell("Sheet1", 0, 1, None, Some("=CELL(\"format\",A1)".to_string()))
                .expect("set formula");
            let before = state.get_cell("Sheet1", 0, 1).expect("get B1");
            assert_eq!(before.value, CellScalar::Text("G".to_string()));

            // Apply a format payload that contains both:
            // - an imported `number_format` string, and
            // - a UI "clear" override (`numberFormat: null`).
            //
            // The UI semantics should treat this as clearing the number format back to General.
            let format = serde_json::json!({
                "number_format": "0.00",
                "numberFormat": null,
            });
            apply_sheet_formatting_deltas_inner(
                &mut state,
                ApplySheetFormattingDeltasRequest {
                    sheet_id: "Sheet1".to_string(),
                    default_format: None,
                    row_formats: None,
                    col_formats: None,
                    format_runs_by_col: None,
                    cell_formats: Some(LimitedSheetCellFormatDeltas(vec![SheetCellFormatDelta {
                        row: 0,
                        col: 0,
                        format,
                    }])),
                },
            )
            .expect("apply formatting delta");

            state.engine.recalculate_with_value_changes_multi_threaded();
            let after = state.get_cell("Sheet1", 0, 1).expect("get B1");
            assert_eq!(after.value, CellScalar::Text("G".to_string()));

            // Rebuild and ensure the clear semantics are re-applied from persistence metadata.
            state
                .rebuild_engine_from_workbook()
                .expect("rebuild engine from workbook");
            state.engine.recalculate_with_value_changes_multi_threaded();
            let after_rebuild = state.get_cell("Sheet1", 0, 1).expect("get B1 after rebuild");
            assert_eq!(after_rebuild.value, CellScalar::Text("G".to_string()));
        }

        let mut recovered_state = AppState::new();
        recovered_state
            .load_workbook_persistent(Workbook::new_empty(None), location)
            .expect("recover workbook from storage");
        let recovered = recovered_state
            .get_cell("Sheet1", 0, 1)
            .expect("get B1 after reload");
        assert_eq!(recovered.value, CellScalar::Text("G".to_string()));
    }

    #[test]
    fn xls_number_formats_survive_persistent_recovery() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xls/tests/fixtures/dates.xls"
        ));

        let tmp = tempfile::tempdir().expect("temp dir");
        let db_path = tmp.path().join("autosave.sqlite");
        let location = WorkbookPersistenceLocation::OnDisk(db_path);

        {
            let workbook = read_xlsx_blocking(fixture_path).expect("read xls workbook");
            let mut state = AppState::new();
            state
                .load_workbook_persistent(workbook, location.clone())
                .expect("load persistent workbook");
        }

        let workbook = read_xlsx_blocking(fixture_path).expect("re-read xls workbook");
        let mut state = AppState::new();
        state
            .load_workbook_persistent(workbook, location)
            .expect("recover workbook from autosave storage");

        let recovered = state.get_workbook().expect("workbook");
        let dates_sheet = recovered
            .sheets
            .iter()
            .find(|s| sheet_name_eq_case_insensitive(&s.name, "Dates"))
            .expect("Dates sheet exists");
        let cell = dates_sheet.get_cell(0, 0);
        assert_eq!(cell.number_format.as_deref(), Some("m/d/yy"));
    }

    #[test]
    fn xlsx_sheet_visibility_and_tab_color_surface_in_workbook_info() {
        use formula_model::{SheetVisibility, TabColor};

        let mut model = formula_model::Workbook::new();
        let visible_id = model.add_sheet("Visible").expect("add sheet");
        let hidden_id = model.add_sheet("Hidden").expect("add sheet");
        let very_hidden_id = model.add_sheet("VeryHidden").expect("add sheet");

        model.sheet_mut(visible_id).expect("sheet").visibility = SheetVisibility::Visible;
        model.sheet_mut(hidden_id).expect("sheet").visibility = SheetVisibility::Hidden;
        model.sheet_mut(very_hidden_id).expect("sheet").visibility = SheetVisibility::VeryHidden;
        model.sheet_mut(hidden_id).expect("sheet").tab_color = Some(TabColor::rgb("FFFF0000"));

        let mut buf = std::io::Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&model, &mut buf).expect("write xlsx bytes");

        let tmp = tempfile::tempdir().expect("temp dir");
        let xlsx_path = tmp.path().join("sheet-metadata.xlsx");
        std::fs::write(&xlsx_path, buf.into_inner()).expect("write xlsx file");

        let workbook = read_xlsx_blocking(&xlsx_path).expect("read xlsx workbook");
        let mut state = AppState::new();
        let info = state.load_workbook(workbook);

        let visible = info
            .sheets
            .iter()
            .find(|s| s.name == "Visible")
            .expect("Visible sheet");
        assert_eq!(visible.visibility, SheetVisibility::Visible);
        assert_eq!(visible.tab_color, None);

        let hidden = info
            .sheets
            .iter()
            .find(|s| s.name == "Hidden")
            .expect("Hidden sheet");
        assert_eq!(hidden.visibility, SheetVisibility::Hidden);
        assert_eq!(
            hidden.tab_color.as_ref().and_then(|c| c.rgb.as_deref()),
            Some("FFFF0000")
        );

        let very_hidden = info
            .sheets
            .iter()
            .find(|s| s.name == "VeryHidden")
            .expect("VeryHidden sheet");
        assert_eq!(very_hidden.visibility, SheetVisibility::VeryHidden);
        assert_eq!(very_hidden.tab_color, None);
    }

    #[test]
    fn sheet_metadata_round_trips_through_patch_based_save() {
        use formula_model::{SheetVisibility, TabColor};
        use formula_xlsx::XlsxPackage;

        let mut model = formula_model::Workbook::new();
        let _ = model.add_sheet("Sheet1").expect("add sheet");
        let _ = model.add_sheet("Sheet2").expect("add sheet");
        let sheet3_id = model.add_sheet("Sheet3").expect("add sheet");
        model.set_sheet_visibility(sheet3_id, SheetVisibility::VeryHidden);

        let mut buf = std::io::Cursor::new(Vec::new());
        formula_xlsx::write_workbook_to_writer(&model, &mut buf).expect("write xlsx bytes");
        let mut base_bytes = buf.into_inner();

        // Add an unknown part to the XLSX container. Patch-based saves should preserve this part,
        // while the full regeneration path would drop it.
        const CUSTOM_PART: &str = "xl/formula-custom-part.xml";
        let custom_payload = b"<custom>Hello</custom>".to_vec();
        let mut pkg = XlsxPackage::from_bytes(&base_bytes).expect("parse xlsx package");
        pkg.set_part(CUSTOM_PART, custom_payload.clone());
        base_bytes = pkg.write_to_bytes().expect("repack xlsx");

        let tmp = tempfile::tempdir().expect("temp dir");
        let xlsx_path = tmp.path().join("base.xlsx");
        std::fs::write(&xlsx_path, base_bytes).expect("write xlsx file");

        let workbook = read_xlsx_blocking(&xlsx_path).expect("read xlsx workbook");
        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .set_sheet_visibility("Sheet2", SheetVisibility::Hidden)
            .expect("hide Sheet2");
        state
            .set_sheet_tab_color("Sheet2", Some(TabColor::rgb("FF00FF00")))
            .expect("set tab color");

        // Sheet metadata edits should not clear `origin_xlsx_bytes` (otherwise we'd lose unknown
        // parts and fall back to the slower/less-faithful regeneration save path).
        assert!(
            state
                .get_workbook()
                .expect("workbook")
                .origin_xlsx_bytes
                .is_some()
        );

        let out_path = tmp.path().join("saved.xlsx");
        let workbook = state.get_workbook().expect("workbook").clone();
        assert!(
            workbook.origin_xlsx_bytes.is_some(),
            "expected metadata edits to preserve origin_xlsx_bytes so save uses patch-based path"
        );
        let saved_bytes = write_xlsx_blocking(&out_path, &workbook).expect("save xlsx");

        let preserved_custom = formula_xlsx::read_part_from_reader(
            std::io::Cursor::new(saved_bytes.as_ref()),
            CUSTOM_PART,
        )
        .expect("read custom part")
        .expect("expected custom part to be preserved");
        assert_eq!(preserved_custom, custom_payload);

        let reparsed = formula_xlsx::read_workbook_from_reader(std::io::Cursor::new(
            saved_bytes.as_ref(),
        ))
        .expect("re-read saved workbook");

        let sheet2 = reparsed
            .sheets
            .iter()
            .find(|s| s.name == "Sheet2")
            .expect("Sheet2");
        assert_eq!(sheet2.visibility, SheetVisibility::Hidden);
        assert_eq!(
            sheet2.tab_color.as_ref().and_then(|c| c.rgb.as_deref()),
            Some("FF00FF00")
        );

        let sheet3 = reparsed
            .sheets
            .iter()
            .find(|s| s.name == "Sheet3")
            .expect("Sheet3");
        assert_eq!(sheet3.visibility, SheetVisibility::VeryHidden);
    }

    #[test]
    fn set_sheet_visibility_rejects_hiding_last_visible_sheet() {
        use formula_model::SheetVisibility;

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .set_sheet_visibility("Sheet2", SheetVisibility::Hidden)
            .expect("hide Sheet2");
        let err = state
            .set_sheet_visibility("Sheet1", SheetVisibility::Hidden)
            .expect_err("expected hiding last visible sheet to fail");
        assert!(
            err.to_string().contains("last visible"),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn mark_saved_treats_xltm_as_xlsx_family_for_origin_bookkeeping() {
        let old_bytes = Arc::<[u8]>::from(vec![1u8, 2, 3]);
        let new_bytes = Arc::<[u8]>::from(vec![9u8, 8, 7]);

        let mut workbook = Workbook::new_empty(Some("/tmp/original.xlsb".to_string()));
        workbook.add_sheet("Sheet1".to_string());
        workbook.origin_xlsx_bytes = Some(old_bytes);
        workbook.origin_xlsb_path = Some("/tmp/original.xlsb".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .mark_saved(Some("/tmp/foo.xltm".to_string()), Some(new_bytes.clone()))
            .expect("mark_saved succeeds");

        let workbook = state.get_workbook().expect("workbook loaded");
        assert_eq!(workbook.origin_xlsb_path, None);
        assert_eq!(workbook.origin_xlsx_bytes.as_deref(), Some(new_bytes.as_ref()));
    }

    #[test]
    fn mark_saved_retains_origin_xlsx_bytes_within_limit() {
        let bytes = Arc::<[u8]>::from(vec![1u8, 2, 3]);

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .mark_saved(Some("/tmp/foo.xlsx".to_string()), Some(bytes.clone()))
            .expect("mark_saved succeeds");

        let workbook = state.get_workbook().expect("workbook loaded");
        assert_eq!(workbook.origin_xlsx_bytes.as_deref(), Some(bytes.as_ref()));
    }

    #[test]
    fn mark_saved_drops_origin_xlsx_bytes_over_limit() {
        let bytes = bytes_over_origin_xlsx_retention_limit();

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .mark_saved(Some("/tmp/foo.xlsx".to_string()), Some(bytes))
            .expect("mark_saved succeeds");

        let workbook = state.get_workbook().expect("workbook loaded");
        assert!(
            workbook.origin_xlsx_bytes.is_none(),
            "expected oversized XLSX baseline to be dropped"
        );
    }

    #[test]
    fn mark_saved_over_limit_clears_existing_origin_xlsx_bytes_baseline() {
        let bytes = bytes_over_origin_xlsx_retention_limit();

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let small_bytes = Arc::<[u8]>::from(vec![9u8, 8, 7]);
        state
            .mark_saved(Some("/tmp/foo.xlsx".to_string()), Some(small_bytes))
            .expect("mark_saved succeeds");

        state
            .mark_saved(Some("/tmp/foo.xlsx".to_string()), Some(bytes))
            .expect("mark_saved succeeds");

        let workbook = state.get_workbook().expect("workbook loaded");
        assert!(
            workbook.origin_xlsx_bytes.is_none(),
            "expected oversized save snapshot to clear the previous baseline so future patch-based saves remain correct"
        );
    }

    #[test]
    fn mark_saved_clears_macros_when_saving_to_xltx() {
        let new_bytes = Arc::<[u8]>::from(vec![9u8, 8, 7]);

        let mut workbook = Workbook::new_empty(Some("/tmp/original.xlsm".to_string()));
        workbook.add_sheet("Sheet1".to_string());
        workbook.vba_project_bin = Some(vec![1u8, 2, 3]);
        workbook.vba_project_signature_bin = Some(vec![4u8, 5, 6]);
        workbook.macro_fingerprint = Some("abc".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .mark_saved(Some("/tmp/foo.xltx".to_string()), Some(new_bytes))
            .expect("mark_saved succeeds");

        let workbook = state.get_workbook().expect("workbook loaded");
        assert_eq!(workbook.vba_project_bin, None);
        assert_eq!(workbook.vba_project_signature_bin, None);
        assert_eq!(workbook.macro_fingerprint, None);
    }

    #[test]
    fn mark_saved_updates_engine_workbook_file_metadata_for_cell_filename() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_formula("=CELL(\"filename\")".to_string()));

        let mut state = AppState::new();
        state.load_workbook(workbook);

        assert_eq!(
            state.engine.get_cell_value("Sheet1", "A1"),
            EngineValue::Text(String::new())
        );

        let saved_path = std::env::temp_dir().join("foo.xlsx");
        let saved_path = saved_path.to_string_lossy().to_string();
        state
            .mark_saved(Some(saved_path.clone()), None)
            .expect("mark_saved succeeds");

        let mut expected_dir = std::path::Path::new(&saved_path)
            .parent()
            .unwrap_or_else(|| std::path::Path::new(""))
            .to_string_lossy()
            .to_string();
        if !expected_dir.ends_with(std::path::MAIN_SEPARATOR) {
            expected_dir.push(std::path::MAIN_SEPARATOR);
        }
        let expected = format!("{expected_dir}[foo.xlsx]Sheet1");

        assert_eq!(
            state.engine.get_cell_value("Sheet1", "A1"),
            EngineValue::Text(expected.clone())
        );

        let workbook = state.get_workbook().expect("workbook loaded");
        let cell = workbook
            .sheet(&sheet_id)
            .expect("sheet exists")
            .get_cell(0, 0);
        assert_eq!(cell.computed_value, CellScalar::Text(expected));
    }

    #[test]
    fn mark_saved_updates_engine_workbook_file_metadata_for_info_directory() {
        fn workbook_dir_for_excel(dir: &str) -> String {
            if dir.is_empty() {
                return String::new();
            }
            if dir.ends_with('/') || dir.ends_with('\\') {
                return dir.to_string();
            }

            // Excel returns directory strings with a trailing path separator. We don't want to probe
            // the OS, so infer the separator from the host-supplied directory string.
            let last_slash = dir.rfind('/');
            let last_backslash = dir.rfind('\\');
            let sep = match (last_slash, last_backslash) {
                (Some(i), Some(j)) => {
                    if i > j {
                        '/'
                    } else {
                        '\\'
                    }
                }
                (Some(_), None) => '/',
                (None, Some(_)) => '\\',
                (None, None) => '/',
            };

            let mut out = String::with_capacity(dir.len() + 1);
            out.push_str(dir);
            out.push(sep);
            out
        }

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_formula("=INFO(\"directory\")".to_string()));

        let mut state = AppState::new();
        state.load_workbook(workbook);

        assert_eq!(
            state.engine.get_cell_value("Sheet1", "A1"),
            EngineValue::Error(ErrorKind::NA)
        );

        let saved_path = std::env::temp_dir().join("foo.xlsx");
        let saved_path = saved_path.to_string_lossy().to_string();

        state
            .mark_saved(Some(saved_path.clone()), None)
            .expect("mark_saved succeeds");

        let dir = std::path::Path::new(&saved_path)
            .parent()
            .unwrap_or_else(|| std::path::Path::new(""))
            .to_string_lossy()
            .to_string();
        let expected_dir = workbook_dir_for_excel(&dir);

        assert_eq!(
            state.engine.get_cell_value("Sheet1", "A1"),
            EngineValue::Text(expected_dir.clone())
        );

        let workbook = state.get_workbook().expect("workbook loaded");
        let cell = workbook
            .sheet(&sheet_id)
            .expect("sheet exists")
            .get_cell(0, 0);
        assert_eq!(cell.computed_value, CellScalar::Text(expected_dir));
    }

    #[test]
    fn mark_saved_sets_engine_filename_metadata_without_directory() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_formula("=CELL(\"filename\")".to_string()));

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .mark_saved(Some("foo.xlsx".to_string()), None)
            .expect("mark_saved succeeds");

        assert_eq!(
            state.engine.get_cell_value("Sheet1", "A1"),
            EngineValue::Text("[foo.xlsx]Sheet1".to_string())
        );

        let workbook = state.get_workbook().expect("workbook loaded");
        let cell = workbook
            .sheet(&sheet_id)
            .expect("sheet exists")
            .get_cell(0, 0);
        assert_eq!(
            cell.computed_value,
            CellScalar::Text("[foo.xlsx]Sheet1".to_string())
        );
    }

    #[test]
    fn mark_saved_updates_engine_workbook_file_metadata_on_save_as() {
        let initial_path = std::env::temp_dir().join("orig.xlsx");
        let initial_path = initial_path.to_string_lossy().to_string();

        let mut workbook = Workbook::new_empty(Some(initial_path.clone()));
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_formula("=CELL(\"filename\")".to_string()));

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut initial_dir = std::path::Path::new(&initial_path)
            .parent()
            .unwrap_or_else(|| std::path::Path::new(""))
            .to_string_lossy()
            .to_string();
        if !initial_dir.ends_with(std::path::MAIN_SEPARATOR) {
            initial_dir.push(std::path::MAIN_SEPARATOR);
        }
        let expected_initial = format!("{initial_dir}[orig.xlsx]Sheet1");
        assert_eq!(
            state.engine.get_cell_value("Sheet1", "A1"),
            EngineValue::Text(expected_initial.clone())
        );

        let new_path = std::env::temp_dir().join("new.xlsx");
        let new_path = new_path.to_string_lossy().to_string();
        state
            .mark_saved(Some(new_path.clone()), None)
            .expect("mark_saved succeeds");

        let mut new_dir = std::path::Path::new(&new_path)
            .parent()
            .unwrap_or_else(|| std::path::Path::new(""))
            .to_string_lossy()
            .to_string();
        if !new_dir.ends_with(std::path::MAIN_SEPARATOR) {
            new_dir.push(std::path::MAIN_SEPARATOR);
        }
        let expected_new = format!("{new_dir}[new.xlsx]Sheet1");

        assert_eq!(
            state.engine.get_cell_value("Sheet1", "A1"),
            EngineValue::Text(expected_new.clone())
        );

        let workbook = state.get_workbook().expect("workbook loaded");
        let cell = workbook
            .sheet(&sheet_id)
            .expect("sheet exists")
            .get_cell(0, 0);
        assert_eq!(cell.computed_value, CellScalar::Text(expected_new));
    }

    #[test]
    fn normalize_formula_matches_formula_model_display_semantics() {
        assert_eq!(normalize_formula(None), None);
        assert_eq!(normalize_formula(Some("".to_string())), None);
        assert_eq!(normalize_formula(Some("   ".to_string())), None);
        assert_eq!(normalize_formula(Some("=".to_string())), None);
        assert_eq!(normalize_formula(Some("   =   ".to_string())), None);

        assert_eq!(
            normalize_formula(Some("=1+1".to_string())),
            Some("=1+1".to_string())
        );
        assert_eq!(
            normalize_formula(Some("1+1".to_string())),
            Some("=1+1".to_string())
        );
        assert_eq!(
            normalize_formula(Some("  =  SUM(A1:A3)  ".to_string())),
            Some("=SUM(A1:A3)".to_string())
        );
        assert_eq!(
            normalize_formula(Some("==1+1".to_string())),
            Some("==1+1".to_string())
        );
    }

    #[test]
    fn move_sheet_reorders_workbook_info() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let before = state.workbook_info().expect("workbook_info before move");
        let before_order = before
            .sheets
            .iter()
            .map(|s| s.id.as_str())
            .collect::<Vec<_>>();
        assert_eq!(before_order, vec!["Sheet1", "Sheet2", "Sheet3"]);

        state
            .move_sheet("Sheet3", 0)
            .expect("move Sheet3 to front");

        let after = state.workbook_info().expect("workbook_info after move");
        let after_order = after
            .sheets
            .iter()
            .map(|s| s.id.as_str())
            .collect::<Vec<_>>();
        assert_eq!(after_order, vec!["Sheet3", "Sheet1", "Sheet2"]);
        assert!(state.has_unsaved_changes());
    }

    #[test]
    fn move_sheet_reorders_workbook_and_storage() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        let sheet1_id = workbook.sheets[0].id.clone();
        let sheet2_id = workbook.sheets[1].id.clone();
        let sheet3_id = workbook.sheets[2].id.clone();

        let mut state = AppState::new();
        state
            .load_workbook_persistent(workbook, WorkbookPersistenceLocation::InMemory)
            .expect("load persistent workbook");

        state.move_sheet(&sheet3_id, 0).expect("move sheet");

        let info = state.workbook_info().expect("workbook info");
        let ids: Vec<String> = info.sheets.iter().map(|s| s.id.clone()).collect();
        assert_eq!(ids, vec![sheet3_id.clone(), sheet1_id.clone(), sheet2_id.clone()]);

        let storage = state.persistent_storage().expect("storage available");
        let workbook_id = state.persistent_workbook_id().expect("workbook id");
        let metas = storage.list_sheets(workbook_id).expect("list sheets");
        let names: Vec<String> = metas.iter().map(|m| m.name.clone()).collect();
        assert_eq!(
            names,
            vec!["Sheet3".to_string(), "Sheet1".to_string(), "Sheet2".to_string()]
        );
    }

    #[test]
    fn reorder_sheets_reorders_workbook_info() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        let sheet1_id = workbook.sheets[0].id.clone();
        let sheet2_id = workbook.sheets[1].id.clone();
        let sheet3_id = workbook.sheets[2].id.clone();

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let before = state.workbook_info().expect("workbook_info before reorder");
        let before_order = before
            .sheets
            .iter()
            .map(|s| s.id.as_str())
            .collect::<Vec<_>>();
        assert_eq!(before_order, vec![sheet1_id.as_str(), sheet2_id.as_str(), sheet3_id.as_str()]);

        state
            .reorder_sheets(vec![sheet3_id.clone(), sheet1_id.clone(), sheet2_id.clone()])
            .expect("reorder sheets");

        let after = state.workbook_info().expect("workbook_info after reorder");
        let after_order = after
            .sheets
            .iter()
            .map(|s| s.id.as_str())
            .collect::<Vec<_>>();
        assert_eq!(after_order, vec![sheet3_id.as_str(), sheet1_id.as_str(), sheet2_id.as_str()]);
        assert!(state.has_unsaved_changes());
    }

    #[test]
    fn reorder_sheets_reorders_workbook_and_storage() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        let sheet1_id = workbook.sheets[0].id.clone();
        let sheet2_id = workbook.sheets[1].id.clone();
        let sheet3_id = workbook.sheets[2].id.clone();

        let mut state = AppState::new();
        state
            .load_workbook_persistent(workbook, WorkbookPersistenceLocation::InMemory)
            .expect("load persistent workbook");

        state
            .reorder_sheets(vec![sheet3_id.clone(), sheet1_id.clone(), sheet2_id.clone()])
            .expect("reorder sheets");

        let info = state.workbook_info().expect("workbook info");
        let ids: Vec<String> = info.sheets.iter().map(|s| s.id.clone()).collect();
        assert_eq!(ids, vec![sheet3_id.clone(), sheet1_id.clone(), sheet2_id.clone()]);

        let storage = state.persistent_storage().expect("storage available");
        let workbook_id = state.persistent_workbook_id().expect("workbook id");
        let metas = storage.list_sheets(workbook_id).expect("list sheets");
        let names: Vec<String> = metas.iter().map(|m| m.name.clone()).collect();
        assert_eq!(
            names,
            vec!["Sheet3".to_string(), "Sheet1".to_string(), "Sheet2".to_string()]
        );
    }

    #[test]
    fn move_sheet_validates_sheet_and_index() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        let sheet1_id = workbook.sheets[0].id.clone();

        let mut state = AppState::new();
        state.load_workbook(workbook);

        assert!(matches!(
            state.move_sheet("missing", 0),
            Err(AppStateError::UnknownSheet(_))
        ));

        assert!(matches!(
            state.move_sheet(&sheet1_id, 2),
            Err(AppStateError::InvalidSheetIndex { .. })
        ));
    }

    #[test]
    fn add_sheet_inserts_at_index_and_persists_order() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let db_path = tmp.path().join("autosave.sqlite");
        let location = WorkbookPersistenceLocation::OnDisk(db_path);

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        let mut state = AppState::new();
        state
            .load_workbook_persistent(workbook, location)
            .expect("load persistent workbook");

        state
            .add_sheet("Inserted".to_string(), None, None, Some(1))
            .expect("add sheet");

        let info = state.workbook_info().expect("workbook info");
        let sheet_names = info
            .sheets
            .iter()
            .map(|s| s.name.clone())
            .collect::<Vec<_>>();
        assert_eq!(
            sheet_names,
            vec![
                "Sheet1".to_string(),
                "Inserted".to_string(),
                "Sheet2".to_string(),
                "Sheet3".to_string()
            ]
        );

        let storage = state.persistent_storage().expect("storage");
        let workbook_id = state.persistent_workbook_id().expect("workbook id");
        let meta_names = storage
            .list_sheets(workbook_id)
            .expect("list sheets")
            .into_iter()
            .map(|s| s.name)
            .collect::<Vec<_>>();
        assert_eq!(
            meta_names,
            vec![
                "Sheet1".to_string(),
                "Inserted".to_string(),
                "Sheet2".to_string(),
                "Sheet3".to_string()
            ]
        );
    }

    #[test]
    fn add_sheet_inserts_after_sheet_id_and_persists_order() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let db_path = tmp.path().join("autosave.sqlite");
        let location = WorkbookPersistenceLocation::OnDisk(db_path);

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        workbook.add_sheet("Sheet3".to_string());

        let mut state = AppState::new();
        state
            .load_workbook_persistent(workbook, location)
            .expect("load persistent workbook");

        state
            .add_sheet(
                "Inserted".to_string(),
                None,
                Some("sHeEt2".to_string()),
                None,
            )
            .expect("add sheet");

        let info = state.workbook_info().expect("workbook info");
        let sheet_names = info
            .sheets
            .iter()
            .map(|s| s.name.clone())
            .collect::<Vec<_>>();
        assert_eq!(
            sheet_names,
            vec![
                "Sheet1".to_string(),
                "Sheet2".to_string(),
                "Inserted".to_string(),
                "Sheet3".to_string()
            ]
        );

        let storage = state.persistent_storage().expect("storage");
        let workbook_id = state.persistent_workbook_id().expect("workbook id");
        let meta_names = storage
            .list_sheets(workbook_id)
            .expect("list sheets")
            .into_iter()
            .map(|s| s.name)
            .collect::<Vec<_>>();
        assert_eq!(
            meta_names,
            vec![
                "Sheet1".to_string(),
                "Sheet2".to_string(),
                "Inserted".to_string(),
                "Sheet3".to_string()
            ]
        );
    }

    #[test]
    fn set_cell_recalculates_dependents() {
        let mut workbook = Workbook::new_empty(Some("fixture.xlsx".to_string()));
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(1.0))),
        );
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_formula("=A1+1".to_string()),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let b1 = state.get_cell(&sheet_id, 0, 1).unwrap();
        assert_eq!(b1.value, CellScalar::Number(2.0));

        let updates = state
            .set_cell(&sheet_id, 0, 0, Some(JsonValue::from(10)), None)
            .unwrap();
        assert!(updates.iter().any(|u| u.row == 0 && u.col == 1));

        let b1_after = state.get_cell(&sheet_id, 0, 1).unwrap();
        assert_eq!(b1_after.value, CellScalar::Number(11.0));
        assert!(state.has_unsaved_changes());
    }

    #[test]
    fn rebuild_engine_loads_workbook_defined_names() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        workbook.defined_names.push(crate::file_io::DefinedName {
            name: "MyName".to_string(),
            refers_to: "42".to_string(),
            sheet_id: None,
            hidden: false,
        });

        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_formula("=MyName+1".to_string()),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let a1 = state.get_cell(&sheet_id, 0, 0).unwrap();
        assert_eq!(a1.value, CellScalar::Number(43.0));
    }

    #[test]
    fn delete_sheet_rewrites_cross_sheet_formulas_to_ref() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("DeletedSheet".to_string());
        let sheet1_id = workbook.sheets[0].id.clone();
        let deleted_id = workbook.sheets[1].id.clone();

        workbook
            .sheet_mut(&sheet1_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_formula("=DeletedSheet!A1".to_string()));

        // 3D reference with the deleted sheet as a boundary should be shifted toward the other boundary.
        workbook
            .sheet_mut(&sheet1_id)
            .unwrap()
            .set_cell(0, 1, Cell::from_formula("=SUM(Sheet1:DeletedSheet!A1)".to_string()));

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state.delete_sheet(&deleted_id).expect("delete sheet");

        let workbook = state.get_workbook().expect("workbook");
        assert_eq!(workbook.cell_formula(&sheet1_id, 0, 0).as_deref(), Some("=#REF!"));
        assert_eq!(
            workbook.cell_formula(&sheet1_id, 0, 1).as_deref(),
            Some("=SUM(Sheet1!A1)")
        );
    }

    #[test]
    fn get_range_and_updates_preserve_typed_values_when_number_format_present() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        let mut cell = Cell::from_literal(Some(CellScalar::Number(1.25)));
        cell.number_format = Some("0.00".to_string());
        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 0, cell);

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let a1 = state.get_cell(&sheet_id, 0, 0).expect("get A1");
        assert_eq!(a1.value, CellScalar::Number(1.25));
        assert_eq!(a1.display_value, "1.25");

        let range = state.get_range(&sheet_id, 0, 0, 0, 0).expect("get range");
        assert_eq!(range[0][0].value, CellScalar::Number(1.25));
        assert_eq!(range[0][0].display_value, "1.25");

        let updates = state
            .set_cell(&sheet_id, 0, 0, Some(JsonValue::from(2.5)), None)
            .expect("set cell");
        let update = updates
            .iter()
            .find(|u| u.row == 0 && u.col == 0)
            .expect("expected A1 update");
        assert_eq!(update.value, CellScalar::Number(2.5));
        assert_eq!(update.display_value, "2.50");
    }

    #[test]
    fn get_range_returns_values_and_preserves_formulas() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        let mut a1 = Cell::from_literal(Some(CellScalar::Number(1.25)));
        a1.number_format = Some("0.00".to_string());

        let sheet = workbook.sheet_mut(&sheet_id).unwrap();
        sheet.set_cell(0, 0, a1);
        sheet.set_cell(0, 1, Cell::from_formula("=A1+1".to_string()));
        sheet.set_cell(1, 0, Cell::from_literal(Some(CellScalar::Text("hello".to_string()))));
        sheet.set_cell(1, 1, Cell::from_formula("=B1*2".to_string()));

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let range = state
            .get_range(&sheet_id, 0, 0, 1, 1)
            .expect("get range");

        assert_eq!(range.len(), 2);
        assert_eq!(range[0].len(), 2);

        assert_eq!(range[0][0].value, CellScalar::Number(1.25));
        assert_eq!(range[0][0].formula, None);
        assert_eq!(range[0][0].display_value, "1.25");

        assert_eq!(range[0][1].value, CellScalar::Number(2.25));
        assert_eq!(range[0][1].formula.as_deref(), Some("=A1+1"));

        assert_eq!(range[1][0].value, CellScalar::Text("hello".to_string()));
        assert_eq!(range[1][0].formula, None);
        assert_eq!(range[1][0].display_value, "hello");

        assert_eq!(range[1][1].value, CellScalar::Number(4.5));
        assert_eq!(range[1][1].formula.as_deref(), Some("=B1*2"));
        assert_eq!(range[1][1].display_value, "4.5");
    }

    #[test]
    fn get_range_rejects_oversized_ranges() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        let mut state = AppState::new();
        state.load_workbook(workbook);

        assert!(
            (MAX_RANGE_DIM as u128) * (MAX_RANGE_DIM as u128) > MAX_RANGE_CELLS_PER_CALL as u128,
            "test expects MAX_RANGE_DIM^2 > MAX_RANGE_CELLS_PER_CALL"
        );

        let end_row = MAX_RANGE_DIM - 1;
        let end_col = MAX_RANGE_DIM - 1;
        let err = state
            .get_range(&sheet_id, 0, 0, end_row, end_col)
            .expect_err("expected oversized get_range to fail");
        assert!(matches!(
            err,
            AppStateError::RangeTooLarge {
                rows: MAX_RANGE_DIM,
                cols: MAX_RANGE_DIM,
                limit: MAX_RANGE_CELLS_PER_CALL
            }
        ));
    }

    #[test]
    fn set_range_rejects_oversized_ranges() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let end_row = MAX_RANGE_DIM - 1;
        let end_col = MAX_RANGE_DIM - 1;
        let err = state
            .set_range(&sheet_id, 0, 0, end_row, end_col, Vec::new())
            .expect_err("expected oversized set_range to fail");
        assert!(matches!(
            err,
            AppStateError::RangeTooLarge {
                rows: MAX_RANGE_DIM,
                cols: MAX_RANGE_DIM,
                limit: MAX_RANGE_CELLS_PER_CALL
            }
        ));
    }

    #[test]
    fn set_range_allows_typical_ui_ranges() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let values = vec![
            vec![(Some(JsonValue::from(1)), None), (Some(JsonValue::from(2)), None)],
            vec![(Some(JsonValue::from(3)), None), (Some(JsonValue::from(4)), None)],
        ];

        let updates = state
            .set_range(&sheet_id, 0, 0, 1, 1, values)
            .expect("set_range should succeed");
        assert!(
            updates.iter().any(|u| u.row == 0 && u.col == 0),
            "expected updates to include A1"
        );

        let a1 = state.get_cell(&sheet_id, 0, 0).expect("get A1");
        assert_eq!(a1.value, CellScalar::Number(1.0));
        let b2 = state.get_cell(&sheet_id, 1, 1).expect("get B2");
        assert_eq!(b2.value, CellScalar::Number(4.0));
    }

    #[test]
    fn set_range_number_format_rejects_oversized_ranges() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let end_row = MAX_RANGE_DIM - 1;
        let end_col = MAX_RANGE_DIM - 1;
        let err = state
            .set_range_number_format(&sheet_id, 0, 0, end_row, end_col, Some("0.00".to_string()))
            .expect_err("expected oversized set_range_number_format to fail");
        assert!(matches!(
            err,
            AppStateError::RangeTooLarge {
                rows: MAX_RANGE_DIM,
                cols: MAX_RANGE_DIM,
                limit: MAX_RANGE_CELLS_PER_CALL
            }
        ));
    }

    #[test]
    fn cell_format_reflects_workbook_number_format_after_engine_rebuild() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .set_range_number_format(&sheet_id, 0, 0, 0, 0, Some("0.00".to_string()))
            .expect("set_range_number_format");
        state
            .set_cell(&sheet_id, 0, 1, None, Some("=CELL(\"format\",A1)".to_string()))
            .expect("set formula");

        state
            .rebuild_engine_from_workbook()
            .expect("rebuild engine");
        state.engine.recalculate_single_threaded();
        state.refresh_computed_values().expect("refresh computed values");

        let b1 = state.get_cell(&sheet_id, 0, 1).expect("get B1");
        assert_eq!(b1.value, CellScalar::Text("F2".to_string()));
    }

    #[test]
    fn create_pivot_table_rejects_oversized_source_range() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let source_range = CellRect {
            start_row: 0,
            start_col: 0,
            end_row: MAX_RANGE_DIM - 1,
            end_col: MAX_RANGE_DIM - 1,
        };

        let err = state
            .create_pivot_table(
                "Pivot".to_string(),
                sheet_id.clone(),
                source_range,
                PivotDestination {
                    sheet_id: sheet_id.clone(),
                    row: 0,
                    col: 0,
                },
                PivotConfig::default(),
            )
            .expect_err("expected oversized pivot source range to fail");

        assert!(matches!(
            err,
            AppStateError::RangeTooLarge {
                rows: MAX_RANGE_DIM,
                cols: MAX_RANGE_DIM,
                limit: MAX_RANGE_CELLS_PER_CALL
            }
        ));
    }

    #[test]
    fn clear_range_clears_values_but_preserves_number_format() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        let mut cell = Cell::from_literal(Some(CellScalar::Number(1.25)));
        cell.number_format = Some("0.00".to_string());
        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 0, cell);

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .clear_range(&sheet_id, 0, 0, 0, 0)
            .expect("clear range");

        let a1 = state.get_cell(&sheet_id, 0, 0).expect("get A1");
        assert_eq!(a1.value, CellScalar::Empty);

        let workbook = state.get_workbook().expect("workbook");
        let sheet = workbook.sheet(&sheet_id).expect("sheet");
        let a1_cell = sheet.get_cell(0, 0);
        assert_eq!(a1_cell.number_format.as_deref(), Some("0.00"));
    }

    #[test]
    fn display_value_uses_general_number_formatting_when_no_explicit_format_is_present() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();
        let date_system = workbook_date_system(&workbook);

        let n = 123_456_789_012.0;
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(n))),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let cell = state.get_cell(&sheet_id, 0, 0).expect("get cell");
        assert_eq!(cell.value, CellScalar::Number(n));

        let expected = format_value(
            FormatValue::Number(n),
            None,
            &FormatOptions {
                date_system,
                ..FormatOptions::default()
            },
        )
        .text;
        assert_eq!(cell.display_value, expected);
    }

    #[test]
    fn display_value_normalizes_negative_zero_for_general_formatting() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(-0.0))),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let cell = state.get_cell(&sheet_id, 0, 0).expect("get cell");
        assert_eq!(cell.display_value, "0");
    }

    #[test]
    fn display_value_respects_workbook_date_system_for_date_formats() {
        let mut workbook = Workbook::new_empty(None);
        workbook.date_system = formula_model::DateSystem::Excel1904;
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        let mut cell = Cell::from_literal(Some(CellScalar::Number(0.0)));
        cell.number_format = Some("m/d/yyyy".to_string());
        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 0, cell);

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let cell = state.get_cell(&sheet_id, 0, 0).expect("get cell");
        assert_eq!(cell.value, CellScalar::Number(0.0));

        let expected = format_value(
            FormatValue::Number(0.0),
            Some("m/d/yyyy"),
            &FormatOptions {
                date_system: formula_format::DateSystem::Excel1904,
                ..FormatOptions::default()
            },
        )
        .text;
        assert_eq!(cell.display_value, expected);
    }

    #[test]
    fn xlsx_1904_date_system_propagates_to_engine_and_roundtrips() {
        use formula_xlsx::XlsxPackage;

        let fixture_path = std::path::Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/metadata/date-system-1904.xlsx"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read 1904 date system xlsx");
        assert_eq!(workbook.date_system, formula_model::DateSystem::Excel1904);
        assert_eq!(workbook.sheets.len(), 1);
        let sheet_id = workbook.sheets[0].id.clone();

        let mut state = AppState::new();
        state.load_workbook(workbook);
        assert_eq!(
            state.engine.date_system(),
            formula_engine::date::ExcelDateSystem::Excel1904
        );

        state
            .set_cell(&sheet_id, 0, 0, None, Some("=DATE(1904,1,1)".to_string()))
            .expect("set formula cell");
        let cell = state.get_cell(&sheet_id, 0, 0).expect("get cell");
        assert_eq!(cell.value, CellScalar::Number(0.0));

        let tmp_dir = tempfile::tempdir().expect("temp dir");
        let xlsx_path = tmp_dir.path().join("roundtrip.xlsx");
        let workbook_to_save = state.get_workbook().expect("workbook").clone();
        let bytes = write_xlsx_blocking(&xlsx_path, &workbook_to_save).expect("write xlsx");

        let pkg = XlsxPackage::from_bytes(bytes.as_ref()).expect("parse saved xlsx");
        let workbook_xml = std::str::from_utf8(
            pkg.part("xl/workbook.xml")
                .expect("saved workbook should contain xl/workbook.xml"),
        )
        .expect("workbook.xml should be utf-8");
        assert!(
            workbook_xml.contains("date1904=\"1\""),
            "expected date1904=\"1\" in workbook.xml, got:\n{workbook_xml}"
        );
    }

    #[test]
    fn goal_seek_commits_solution_and_can_be_undone() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(1.0))),
        );
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_formula("=A1*A1".to_string()),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);
        assert!(!state.has_unsaved_changes());

        let mut params = GoalSeekParams::new("B1", 9.0, "A1");
        params.tolerance = 1e-9;
        let (result, _updates) = state.goal_seek(&sheet_id, params).unwrap();
        assert!(result.success(), "{result:?}");

        let a1 = state.get_cell(&sheet_id, 0, 0).unwrap().value;
        let b1 = state.get_cell(&sheet_id, 0, 1).unwrap().value;
        match a1 {
            CellScalar::Number(v) => assert!((v - 3.0).abs() < 1e-6, "A1 = {v}"),
            other => panic!("expected numeric A1, got {other:?}"),
        }
        match b1 {
            CellScalar::Number(v) => assert!((v - 9.0).abs() < 1e-6, "B1 = {v}"),
            other => panic!("expected numeric B1, got {other:?}"),
        }
        assert!(state.has_unsaved_changes());

        state.undo().unwrap();
        assert_eq!(
            state.get_cell(&sheet_id, 0, 0).unwrap().value,
            CellScalar::Number(1.0)
        );
        assert_eq!(
            state.get_cell(&sheet_id, 0, 1).unwrap().value,
            CellScalar::Number(1.0)
        );
    }

    #[test]
    fn monte_carlo_does_not_mutate_workbook_inputs_or_dirty_flag() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(0.0))),
        );
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_formula("=A1+1".to_string()),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);
        assert!(!state.has_unsaved_changes());

        let before_a1 = state.get_cell(&sheet_id, 0, 0).unwrap().value;
        let before_b1 = state.get_cell(&sheet_id, 0, 1).unwrap().value;

        let mut config = SimulationConfig::new(2_000);
        config.seed = 123;
        config.input_distributions = vec![InputDistribution {
            cell: WhatIfCellRef::from("A1"),
            distribution: Distribution::Normal {
                mean: 0.0,
                std_dev: 1.0,
            },
        }];
        config.output_cells = vec![WhatIfCellRef::from("B1")];

        let result = state.run_monte_carlo(&sheet_id, config).unwrap();
        let stats = result.output_stats.get(&WhatIfCellRef::from("B1")).unwrap();
        assert!((stats.mean - 1.0).abs() < 0.05, "mean = {}", stats.mean);

        assert_eq!(state.get_cell(&sheet_id, 0, 0).unwrap().value, before_a1);
        assert_eq!(state.get_cell(&sheet_id, 0, 1).unwrap().value, before_b1);
        assert!(!state.has_unsaved_changes());
    }

    #[test]
    fn scenario_manager_apply_restore_and_summary_report() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(2.0))),
        );
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_formula("=A1*2".to_string()),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        // Create scenario "Low" capturing A1=3.
        state
            .set_cell(&sheet_id, 0, 0, Some(JsonValue::from(3)), None)
            .unwrap();
        let low = state
            .create_scenario(
                &sheet_id,
                "Low".to_string(),
                vec![WhatIfCellRef::from("A1")],
                "tester".to_string(),
                None,
            )
            .unwrap();

        // Create scenario "High" capturing A1=8.
        state
            .set_cell(&sheet_id, 0, 0, Some(JsonValue::from(8)), None)
            .unwrap();
        let high = state
            .create_scenario(
                &sheet_id,
                "High".to_string(),
                vec![WhatIfCellRef::from("A1")],
                "tester".to_string(),
                None,
            )
            .unwrap();

        // Return to base (A1=2) before applying scenarios.
        state
            .set_cell(&sheet_id, 0, 0, Some(JsonValue::from(2)), None)
            .unwrap();
        assert_eq!(
            state.get_cell(&sheet_id, 0, 1).unwrap().value,
            CellScalar::Number(4.0)
        );

        state.apply_scenario(&sheet_id, low.id).unwrap();
        assert_eq!(
            state.get_cell(&sheet_id, 0, 1).unwrap().value,
            CellScalar::Number(6.0)
        );

        state.restore_base_scenario(&sheet_id).unwrap();
        assert_eq!(
            state.get_cell(&sheet_id, 0, 1).unwrap().value,
            CellScalar::Number(4.0)
        );

        let report = state
            .generate_summary_report(
                &sheet_id,
                vec![WhatIfCellRef::from("B1")],
                vec![low.id, high.id],
            )
            .unwrap();

        let base = report.results.get("Base").unwrap();
        assert_eq!(
            base.get(&WhatIfCellRef::from("B1"))
                .unwrap()
                .as_number()
                .unwrap(),
            4.0
        );

        let low_row = report.results.get("Low").unwrap();
        assert_eq!(
            low_row
                .get(&WhatIfCellRef::from("B1"))
                .unwrap()
                .as_number()
                .unwrap(),
            6.0
        );

        let high_row = report.results.get("High").unwrap();
        assert_eq!(
            high_row
                .get(&WhatIfCellRef::from("B1"))
                .unwrap()
                .as_number()
                .unwrap(),
            16.0
        );
    }

    #[test]
    fn undo_redo_round_trip() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(1.0))),
        );
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_formula("=A1+1".to_string()),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state
            .set_cell(&sheet_id, 0, 0, Some(JsonValue::from(10)), None)
            .unwrap();
        assert_eq!(
            state.get_cell(&sheet_id, 0, 1).unwrap().value,
            CellScalar::Number(11.0)
        );

        state.undo().unwrap();
        assert_eq!(
            state.get_cell(&sheet_id, 0, 1).unwrap().value,
            CellScalar::Number(2.0)
        );

        state.redo().unwrap();
        assert_eq!(
            state.get_cell(&sheet_id, 0, 1).unwrap().value,
            CellScalar::Number(11.0)
        );
    }

    #[test]
    fn xlsx_round_trip_through_file_io() {
        use formula_xlsx::print::{CellRange, Orientation, PageSetup, PaperSize, Scaling};

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(3.0))),
        );
        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_formula("=A1+4".to_string()),
        );

        // Set a print area + a non-default page setup and ensure it survives the
        // write/read round-trip.
        workbook
            .print_settings
            .sheets
            .push(formula_xlsx::print::SheetPrintSettings {
                sheet_name: "sheet1".to_string(),
                print_area: Some(vec![CellRange {
                    start_row: 1,
                    end_row: 10,
                    start_col: 1,
                    end_col: 4,
                }]),
                print_titles: None,
                page_setup: PageSetup {
                    orientation: Orientation::Landscape,
                    paper_size: PaperSize { code: 9 },
                    margins: formula_xlsx::print::PageMargins::default(),
                    scaling: Scaling::Percent(90),
                },
                manual_page_breaks: formula_xlsx::print::ManualPageBreaks::default(),
            });

        let tmp_dir = tempfile::tempdir().expect("temp dir");
        let path = tmp_dir.path().join("roundtrip.xlsx");
        write_xlsx_blocking(&path, &workbook).expect("write");

        let loaded = read_xlsx_blocking(&path).expect("read");
        let mut state = AppState::new();
        let info = state.load_workbook(loaded);
        assert_eq!(info.sheets.len(), 1);

        let sheet_id = info.sheets[0].id.clone();
        let b1 = state.get_cell(&sheet_id, 0, 1).unwrap();
        assert_eq!(b1.value, CellScalar::Number(7.0));

        let print_settings = state
            .sheet_print_settings(&sheet_id)
            .expect("sheet print settings");
        assert_eq!(
            print_settings.print_area.as_deref(),
            Some(
                &[CellRange {
                    start_row: 1,
                    end_row: 10,
                    start_col: 1,
                    end_col: 4,
                }][..]
            )
        );
        assert_eq!(
            print_settings.page_setup.orientation,
            Orientation::Landscape
        );
        assert_eq!(print_settings.page_setup.paper_size.code, 9);
        assert_eq!(print_settings.page_setup.scaling, Scaling::Percent(90));
    }

    #[test]
    fn save_as_xlsx_drops_macros_and_prevents_resurrection() {
        use formula_xlsx::XlsxPackage;

        let fixture_path = std::path::Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../fixtures/xlsx/macros/basic.xlsm"
        ));
        let workbook = read_xlsx_blocking(fixture_path).expect("read xlsm fixture");
        assert!(
            workbook.vba_project_bin.is_some(),
            "fixture should contain vbaProject.bin"
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let tmp_dir = tempfile::tempdir().expect("temp dir");
        let xlsx_path = tmp_dir.path().join("converted.xlsx");
        let workbook_to_save = state.get_workbook().expect("workbook").clone();
        write_xlsx_blocking(&xlsx_path, &workbook_to_save).expect("write as xlsx");

        let saved_bytes = Arc::<[u8]>::from(std::fs::read(&xlsx_path).expect("read saved xlsx"));
        state
            .mark_saved(
                Some(xlsx_path.to_string_lossy().to_string()),
                Some(saved_bytes),
            )
            .expect("mark saved");

        assert!(
            state
                .get_workbook()
                .expect("workbook after save")
                .vba_project_bin
                .is_none(),
            "expected in-memory macro payload to be cleared after saving as .xlsx"
        );

        // Ensure macros can't come back if we later save as `.xlsm` (matches Excel's behavior).
        let xlsm_path = tmp_dir.path().join("converted.xlsm");
        let workbook_to_save = state.get_workbook().expect("workbook").clone();
        write_xlsx_blocking(&xlsm_path, &workbook_to_save).expect("write as xlsm");

        let written_bytes = std::fs::read(&xlsm_path).expect("read saved xlsm");
        let written_pkg = XlsxPackage::from_bytes(&written_bytes).expect("parse saved xlsm");
        assert!(
            written_pkg.vba_project_bin().is_none(),
            "expected vbaProject.bin to remain absent after converting to .xlsx"
        );
    }

    #[test]
    fn sheet_print_settings_lookup_is_case_insensitive() {
        let mut sheets = vec![default_sheet_print_settings("sheet1".to_string())];
        ensure_sheet_print_settings(&mut sheets, "Sheet1");
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].sheet_name, "Sheet1");
    }

    #[test]
    fn engine_can_evaluate_sum_function() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();
        let sheet = workbook.sheet_mut(&sheet_id).unwrap();
        sheet.set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(1.0))));
        sheet.set_cell(1, 0, Cell::from_literal(Some(CellScalar::Number(2.0))));
        sheet.set_cell(2, 0, Cell::from_literal(Some(CellScalar::Number(3.0))));
        sheet.set_cell(0, 1, Cell::from_formula("=SUM(A1:A3)".to_string()));

        let mut state = AppState::new();
        state.load_workbook(workbook);
        let b1 = state.get_cell(&sheet_id, 0, 1).unwrap();
        assert_eq!(b1.value, CellScalar::Number(6.0));
    }

    #[test]
    fn engine_can_evaluate_formulas_against_columnar_sheet_data() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        let table = import_csv_to_columnar_table(
            std::io::Cursor::new("value\n1\n2\n3\n"),
            CsvOptions::default(),
        )
        .expect("import csv");
        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_columnar_table(Arc::new(table));

        workbook.sheet_mut(&sheet_id).unwrap().set_cell(
            0,
            1,
            Cell::from_formula("=SUM(A1:A3)".to_string()),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let b1 = state.get_cell(&sheet_id, 0, 1).unwrap();
        assert_eq!(b1.value, CellScalar::Number(6.0));
    }

    #[test]
    fn csv_imported_columnar_values_are_visible_through_get_cell_and_get_range() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();

        let table = import_csv_to_columnar_table(
            std::io::Cursor::new("value\n1\n2\n3\n"),
            CsvOptions::default(),
        )
        .expect("import csv");
        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_columnar_table(Arc::new(table));

        let mut state = AppState::new();
        state.load_workbook(workbook);

        assert_eq!(
            state.get_cell(&sheet_id, 0, 0).unwrap().value,
            CellScalar::Number(1.0)
        );

        let range = state.get_range(&sheet_id, 0, 0, 2, 0).unwrap();
        assert_eq!(range.len(), 3);
        assert_eq!(range[0].len(), 1);
        assert_eq!(range[0][0].value, CellScalar::Number(1.0));
        assert_eq!(range[1][0].value, CellScalar::Number(2.0));
        assert_eq!(range[2][0].value, CellScalar::Number(3.0));
    }

    #[test]
    fn cross_sheet_references_recalculate() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        let sheet1_id = workbook.sheets[0].id.clone();
        let sheet2_id = workbook.sheets[1].id.clone();

        workbook.sheet_mut(&sheet1_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(1.0))),
        );
        workbook.sheet_mut(&sheet2_id).unwrap().set_cell(
            0,
            1,
            Cell::from_formula("=Sheet1!A1+1".to_string()),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let b1 = state.get_cell(&sheet2_id, 0, 1).unwrap();
        assert_eq!(b1.value, CellScalar::Number(2.0));
    }

    #[test]
    fn delete_sheet_rejects_last_sheet() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let sheet_id = workbook.sheets[0].id.clone();
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let err = state.delete_sheet(&sheet_id).expect_err("expected last-sheet guard");
        assert!(matches!(err, AppStateError::CannotDeleteLastSheet));
    }

    #[test]
    fn delete_sheet_rewrites_formulas_to_ref_error() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        let sheet1_id = workbook.sheets[0].id.clone();
        let sheet2_id = workbook.sheets[1].id.clone();

        workbook
            .sheet_mut(&sheet1_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_formula("=Sheet2!A1".to_string()));
        workbook
            .sheet_mut(&sheet2_id)
            .unwrap()
            .set_cell(0, 0, Cell::from_literal(Some(CellScalar::Number(10.0))));

        // Seed print settings + defined names that reference the deleted sheet.
        workbook
            .print_settings
            .sheets
            .push(default_sheet_print_settings("Sheet2".to_string()));
        workbook.defined_names.push(crate::file_io::DefinedName {
            name: "GlobalName".to_string(),
            refers_to: "Sheet2!$A$1".to_string(),
            sheet_id: None,
            hidden: false,
        });
        workbook.defined_names.push(crate::file_io::DefinedName {
            name: "LocalName".to_string(),
            refers_to: "$A$1".to_string(),
            sheet_id: Some(sheet2_id.clone()),
            hidden: false,
        });

        let mut state = AppState::new();
        state.load_workbook(workbook);

        state.delete_sheet(&sheet2_id).expect("delete sheet");

        let wb = state.get_workbook().expect("workbook exists");
        assert_eq!(wb.sheets.len(), 1);

        let sheet1 = wb.sheet(&sheet1_id).expect("Sheet1 exists");
        let a1 = sheet1.cells.get(&(0, 0)).expect("A1 exists");
        assert_eq!(a1.formula.as_deref(), Some("=#REF!"));

        let cell = state.get_cell(&sheet1_id, 0, 0).unwrap();
        assert_eq!(cell.value, CellScalar::Error("#REF!".to_string()));

        // Print settings for the deleted sheet should be removed.
        assert!(
            !wb.print_settings
                .sheets
                .iter()
                .any(|s| s.sheet_name.eq_ignore_ascii_case("Sheet2")),
            "expected print settings for deleted sheet to be removed"
        );

        // Defined name scoped to the deleted sheet should be removed; the workbook-scoped name
        // should be rewritten to #REF!.
        assert!(
            !wb.defined_names
                .iter()
                .any(|n| n.name.eq_ignore_ascii_case("LocalName")),
            "expected sheet-scoped name to be removed"
        );
        let global = wb
            .defined_names
            .iter()
            .find(|n| n.name.eq_ignore_ascii_case("GlobalName"))
            .expect("GlobalName exists");
        assert_eq!(global.refers_to, "#REF!");
    }

    #[test]
    fn cross_sheet_references_resolve_even_if_target_sheet_loaded_later() {
        // Ensure the engine can compile a formula that references a sheet that exists
        // in the workbook but whose cells haven't been loaded into the engine yet.
        //
        // This used to "stick" as #REF! because sheet ids were assigned lazily during
        // compilation; now we pre-create sheets before setting formulas.
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        let sheet1_id = workbook.sheets[0].id.clone();
        let sheet2_id = workbook.sheets[1].id.clone();

        // Sheet1 contains a formula referencing Sheet2, but Sheet2's value is set later.
        workbook.sheet_mut(&sheet1_id).unwrap().set_cell(
            0,
            1,
            Cell::from_formula("=Sheet2!A1+1".to_string()),
        );
        workbook.sheet_mut(&sheet2_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(41.0))),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let b1 = state.get_cell(&sheet1_id, 0, 1).unwrap();
        assert_eq!(b1.value, CellScalar::Number(42.0));
    }

    #[test]
    fn what_if_sheet_name_resolution_is_case_insensitive() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("My Sheet".to_string());
        let sheet1_id = workbook.sheets[0].id.clone();
        let my_sheet_id = workbook.sheets[1].id.clone();

        let (resolved, row, col) = resolve_cell_ref(
            &workbook,
            &sheet1_id,
            "Sheet1",
            &WhatIfCellRef::from("sheet1!A1"),
        )
        .unwrap();
        assert_eq!(resolved, sheet1_id);
        assert_eq!(row, 0);
        assert_eq!(col, 0);

        let (resolved, row, col) = resolve_cell_ref(
            &workbook,
            &sheet1_id,
            "Sheet1",
            &WhatIfCellRef::from("'my sheet'!B2"),
        )
        .unwrap();
        assert_eq!(resolved, my_sheet_id);
        assert_eq!(row, 1);
        assert_eq!(col, 1);
    }

    #[test]
    fn quoted_sheet_references_recalculate() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("My Sheet".to_string());
        workbook.add_sheet("Sheet1".to_string());
        let my_sheet_id = workbook.sheets[0].id.clone();
        let sheet1_id = workbook.sheets[1].id.clone();

        workbook.sheet_mut(&my_sheet_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(10.0))),
        );
        workbook.sheet_mut(&sheet1_id).unwrap().set_cell(
            0,
            0,
            Cell::from_formula("='My Sheet'!A1+1".to_string()),
        );

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let a1 = state.get_cell(&sheet1_id, 0, 0).unwrap();
        assert_eq!(a1.value, CellScalar::Number(11.0));
    }

    #[test]
    fn rename_sheet_rewrites_cross_sheet_formulas_and_metadata() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        let sheet1_id = workbook.sheets[0].id.clone();
        let sheet2_id = workbook.sheets[1].id.clone();

        workbook.sheet_mut(&sheet2_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(41.0))),
        );
        workbook.sheet_mut(&sheet1_id).unwrap().set_cell(
            0,
            1,
            Cell::from_formula("=Sheet2!A1+1".to_string()),
        );

        workbook.defined_names.push(crate::file_io::DefinedName {
            name: "MyRange".to_string(),
            refers_to: "Sheet2!A1".to_string(),
            sheet_id: None,
            hidden: false,
        });
        workbook.defined_names.push(crate::file_io::DefinedName {
            name: "LocalRange".to_string(),
            refers_to: "Sheet2!A1".to_string(),
            sheet_id: Some("Sheet2".to_string()),
            hidden: false,
        });

        workbook.tables.push(crate::file_io::Table {
            name: "Table1".to_string(),
            sheet_id: "Sheet2".to_string(),
            start_row: 0,
            start_col: 0,
            end_row: 0,
            end_col: 0,
            columns: vec!["Col1".to_string()],
        });

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let initial = state.get_cell(&sheet1_id, 0, 1).unwrap();
        assert_eq!(initial.formula.as_deref(), Some("=Sheet2!A1+1"));
        assert_eq!(initial.value, CellScalar::Number(42.0));

        // Plain rename: unquoted -> unquoted.
        state
            .rename_sheet(&sheet2_id, "Budget".to_string())
            .expect("rename Sheet2 -> Budget");
        let renamed = state.get_cell(&sheet1_id, 0, 1).unwrap();
        assert_eq!(renamed.formula.as_deref(), Some("=Budget!A1+1"));
        assert_eq!(renamed.value, CellScalar::Number(42.0));

        let wb = state.get_workbook().unwrap();
        assert!(wb.defined_names.iter().any(|n| n.name == "MyRange" && n.refers_to == "Budget!A1"));
        assert!(wb
            .defined_names
            .iter()
            .any(|n| n.name == "LocalRange" && n.refers_to == "Budget!A1" && n.sheet_id.as_deref() == Some("Budget")));
        assert!(wb.tables.iter().any(|t| t.name == "Table1" && t.sheet_id == "Budget"));

        // Tricky sheet names should be quoted/escaped like Excel.
        let cases = [
            ("My Sheet", "='My Sheet'!A1+1", "'My Sheet'!A1"),
            ("O'Brien", "='O''Brien'!A1+1", "'O''Brien'!A1"),
            ("预算", "='预算'!A1+1", "'预算'!A1"),
        ];

        for (new_name, expected_cell_formula, expected_defined_ref) in cases {
            state
                .rename_sheet(&sheet2_id, new_name.to_string())
                .unwrap_or_else(|e| panic!("rename Budget -> {new_name} failed: {e}"));
            let cell = state.get_cell(&sheet1_id, 0, 1).unwrap();
            assert_eq!(cell.formula.as_deref(), Some(expected_cell_formula));
            assert_eq!(cell.value, CellScalar::Number(42.0));

            let wb = state.get_workbook().unwrap();
            assert!(wb
                .defined_names
                .iter()
                .any(|n| n.name == "MyRange" && n.refers_to == expected_defined_ref));
            assert!(wb.defined_names.iter().any(|n| {
                n.name == "LocalRange"
                    && n.refers_to == expected_defined_ref
                    && n.sheet_id.as_deref() == Some(new_name)
            }));
            assert!(wb.tables.iter().any(|t| t.name == "Table1" && t.sheet_id == new_name));
        }
    }

    #[test]
    fn rename_sheet_updates_persistent_storage_formulas() {
        let tmp = tempfile::tempdir().expect("temp dir");
        let db_path = tmp.path().join("workbook.sqlite");

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        let sheet1_id = workbook.sheets[0].id.clone();
        let sheet2_id = workbook.sheets[1].id.clone();

        workbook.sheet_mut(&sheet2_id).unwrap().set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Number(1.0))),
        );
        workbook.sheet_mut(&sheet1_id).unwrap().set_cell(
            0,
            0,
            Cell::from_formula("=Sheet2!A1+1".to_string()),
        );

        let mut state = AppState::new();
        state
            .load_workbook_persistent(workbook, WorkbookPersistenceLocation::OnDisk(db_path))
            .expect("load persistent workbook");

        let before = state.get_cell(&sheet1_id, 0, 0).unwrap();
        assert_eq!(before.formula.as_deref(), Some("=Sheet2!A1+1"));
        assert_eq!(before.value, CellScalar::Number(2.0));

        state
            .rename_sheet(&sheet2_id, "Budget".to_string())
            .expect("rename Sheet2 -> Budget");

        // Ensure the viewport/cache path returns the rewritten formula text (not the stale value
        // from the pre-rename in-memory pages).
        let after = state.get_cell(&sheet1_id, 0, 0).unwrap();
        assert_eq!(after.formula.as_deref(), Some("=Budget!A1+1"));
        assert_eq!(after.value, CellScalar::Number(2.0));
    }

    #[test]
    fn col_index_to_letters_matches_excel_a1() {
        assert_eq!(col_index_to_letters(0), "A");
        assert_eq!(col_index_to_letters(25), "Z");
        assert_eq!(col_index_to_letters(26), "AA");
        assert_eq!(col_index_to_letters(27), "AB");
        assert_eq!(col_index_to_letters(701), "ZZ");
        assert_eq!(col_index_to_letters(702), "AAA");
    }

    #[test]
    fn coord_to_a1_formats_u32_max_row_without_overflow() {
        assert_eq!(coord_to_a1(u32::MAX as usize, 0), "A4294967296");
    }

    fn simple_pivot_config() -> PivotConfig {
        PivotConfig {
            row_fields: vec![PivotField::new("Region")],
            column_fields: Vec::new(),
            value_fields: vec![ValueField {
                source_field: "Sales".into(),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: Vec::new(),
            calculated_fields: Vec::new(),
            calculated_items: Vec::new(),
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        }
    }

    #[test]
    fn creating_pivot_table_writes_expected_output() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Data".to_string());
        workbook.add_sheet("Pivot".to_string());
        let data_sheet_id = workbook.sheets[0].id.clone();
        let pivot_sheet_id = workbook.sheets[1].id.clone();

        let sheet = workbook.sheet_mut(&data_sheet_id).unwrap();
        sheet.set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Text("Region".to_string()))),
        );
        sheet.set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Text("Sales".to_string()))),
        );
        sheet.set_cell(
            1,
            0,
            Cell::from_literal(Some(CellScalar::Text("East".to_string()))),
        );
        sheet.set_cell(1, 1, Cell::from_literal(Some(CellScalar::Number(100.0))));
        sheet.set_cell(
            2,
            0,
            Cell::from_literal(Some(CellScalar::Text("West".to_string()))),
        );
        sheet.set_cell(2, 1, Cell::from_literal(Some(CellScalar::Number(200.0))));
        sheet.set_cell(
            3,
            0,
            Cell::from_literal(Some(CellScalar::Text("East".to_string()))),
        );
        sheet.set_cell(3, 1, Cell::from_literal(Some(CellScalar::Number(50.0))));

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let (_pivot_id, _updates) = state
            .create_pivot_table(
                "Sales by Region".to_string(),
                data_sheet_id.clone(),
                CellRect {
                    start_row: 0,
                    start_col: 0,
                    end_row: 3,
                    end_col: 1,
                },
                PivotDestination {
                    sheet_id: pivot_sheet_id.clone(),
                    row: 0,
                    col: 0,
                },
                simple_pivot_config(),
            )
            .unwrap();

        assert_eq!(
            state.get_cell(&pivot_sheet_id, 0, 0).unwrap().value,
            CellScalar::Text("Region".to_string())
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 0, 1).unwrap().value,
            CellScalar::Text("Sum of Sales".to_string())
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 1, 0).unwrap().value,
            CellScalar::Text("East".to_string())
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 1, 1).unwrap().value,
            CellScalar::Number(150.0)
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 2, 0).unwrap().value,
            CellScalar::Text("West".to_string())
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 2, 1).unwrap().value,
            CellScalar::Number(200.0)
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 3, 0).unwrap().value,
            CellScalar::Text("Grand Total".to_string())
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 3, 1).unwrap().value,
            CellScalar::Number(350.0)
        );
    }

    #[test]
    fn pivot_source_records_use_display_field_when_degrading_to_text() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Data".to_string());
        workbook.add_sheet("Pivot".to_string());
        let data_sheet_id = workbook.sheets[0].id.clone();
        let pivot_sheet_id = workbook.sheets[1].id.clone();

        let sheet = workbook.sheet_mut(&data_sheet_id).unwrap();
        sheet.set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Text("Item".to_string()))),
        );
        sheet.set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Text("Sales".to_string()))),
        );
        sheet.set_cell(1, 1, Cell::from_literal(Some(CellScalar::Number(10.0))));
        sheet.set_cell(2, 1, Cell::from_literal(Some(CellScalar::Number(20.0))));

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let data_sheet_name = state
            .get_workbook()
            .unwrap()
            .sheet(&data_sheet_id)
            .unwrap()
            .name
            .clone();

        // Use the same fallback display string for both records so the pivot will incorrectly
        // group them together if `display_field` is ignored.
        let mut record_1 = formula_engine::Record::new("Fallback").field(
            "Name",
            EngineValue::Text("Apple".to_string()),
        );
        record_1.display_field = Some("Name".to_string());
        let mut record_2 = formula_engine::Record::new("Fallback").field(
            "Name",
            EngineValue::Text("Banana".to_string()),
        );
        record_2.display_field = Some("Name".to_string());

        state
            .engine
            .set_cell_value(&data_sheet_name, "A2", EngineValue::Record(record_1))
            .expect("set A2 record");
        state
            .engine
            .set_cell_value(&data_sheet_name, "A3", EngineValue::Record(record_2))
            .expect("set A3 record");

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Item")],
            column_fields: Vec::new(),
            value_fields: vec![ValueField {
                source_field: PivotFieldRef::CacheFieldName("Sales".to_string()),
                name: "Sum of Sales".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: Vec::new(),
            calculated_fields: Vec::new(),
            calculated_items: Vec::new(),
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        state
            .create_pivot_table(
                "By Item".to_string(),
                data_sheet_id,
                CellRect {
                    start_row: 0,
                    start_col: 0,
                    end_row: 2,
                    end_col: 1,
                },
                PivotDestination {
                    sheet_id: pivot_sheet_id.clone(),
                    row: 0,
                    col: 0,
                },
                cfg,
            )
            .unwrap();

        let label_1 = state.get_cell(&pivot_sheet_id, 1, 0).unwrap().display_value;
        let label_2 = state.get_cell(&pivot_sheet_id, 2, 0).unwrap().display_value;
        let mut labels = vec![label_1, label_2];
        labels.sort();
        assert_eq!(labels, vec!["Apple".to_string(), "Banana".to_string()]);

        let value_1 = state.get_cell(&pivot_sheet_id, 1, 1).unwrap().value;
        let value_2 = state.get_cell(&pivot_sheet_id, 2, 1).unwrap().value;
        let mut values = vec![];
        for v in [value_1, value_2] {
            match v {
                CellScalar::Number(n) => values.push(n),
                other => panic!("expected numeric pivot value, got {other:?}"),
            }
        }
        values.sort_by(|a, b| a.total_cmp(b));
        assert_eq!(values, vec![10.0, 20.0]);

        assert_eq!(
            state.get_cell(&pivot_sheet_id, 3, 1).unwrap().value,
            CellScalar::Number(30.0)
        );
    }

    #[test]
    fn pivot_source_date_number_formats_are_inferred_as_dates() {
        use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Data".to_string());
        workbook.add_sheet("Pivot".to_string());
        let data_sheet_id = workbook.sheets[0].id.clone();
        let pivot_sheet_id = workbook.sheets[1].id.clone();

        let date1_serial =
            ymd_to_serial(ExcelDate::new(2024, 1, 15), ExcelDateSystem::EXCEL_1900).unwrap() as f64;
        let date2_serial =
            ymd_to_serial(ExcelDate::new(2024, 2, 1), ExcelDateSystem::EXCEL_1900).unwrap() as f64;

        let sheet = workbook.sheet_mut(&data_sheet_id).unwrap();
        sheet.set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Text("Date".to_string()))),
        );
        sheet.set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Text("Amount".to_string()))),
        );

        let mut date_cell_1 = Cell::from_literal(Some(CellScalar::Number(date1_serial)));
        date_cell_1.number_format = Some("m/d/yyyy".to_string());
        sheet.set_cell(1, 0, date_cell_1);
        sheet.set_cell(1, 1, Cell::from_literal(Some(CellScalar::Number(10.0))));

        let mut date_cell_2 = Cell::from_literal(Some(CellScalar::Number(date2_serial)));
        date_cell_2.number_format = Some("m/d/yyyy".to_string());
        sheet.set_cell(2, 0, date_cell_2);
        sheet.set_cell(2, 1, Cell::from_literal(Some(CellScalar::Number(20.0))));

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Date")],
            column_fields: Vec::new(),
            value_fields: vec![ValueField {
                source_field: "Amount".into(),
                name: "Sum of Amount".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: Vec::new(),
            calculated_fields: Vec::new(),
            calculated_items: Vec::new(),
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        state
            .create_pivot_table(
                "By Date".to_string(),
                data_sheet_id,
                CellRect {
                    start_row: 0,
                    start_col: 0,
                    end_row: 2,
                    end_col: 1,
                },
                PivotDestination {
                    sheet_id: pivot_sheet_id.clone(),
                    row: 0,
                    col: 0,
                },
                cfg,
            )
            .unwrap();

        // Row label values should be stored as date serial numbers, with a date number format for
        // display (Excel stores dates as numbers + formatting).
        let label_1 = state.get_cell(&pivot_sheet_id, 1, 0).unwrap();
        assert_eq!(label_1.value, CellScalar::Number(date1_serial));
        assert_eq!(label_1.display_value, "1/15/2024".to_string());
        let label_2 = state.get_cell(&pivot_sheet_id, 2, 0).unwrap();
        assert_eq!(label_2.value, CellScalar::Number(date2_serial));
        assert_eq!(label_2.display_value, "2/1/2024".to_string());

        let workbook = state.workbook.as_ref().unwrap();
        let pivot_sheet = workbook.sheet(&pivot_sheet_id).unwrap();
        assert_eq!(pivot_sheet.get_cell(1, 0).number_format.as_deref(), Some("m/d/yyyy"));
        assert_eq!(pivot_sheet.get_cell(2, 0).number_format.as_deref(), Some("m/d/yyyy"));

        // Numeric measure column should remain numeric (no date inference).
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 1, 1).unwrap().value,
            CellScalar::Number(10.0)
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 2, 1).unwrap().value,
            CellScalar::Number(20.0)
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 3, 1).unwrap().value,
            CellScalar::Number(30.0)
        );
    }

    #[test]
    fn pivot_date_inference_respects_engine_column_number_formats() {
        use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};
        use formula_model::Style;

        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Data".to_string());
        workbook.add_sheet("Pivot".to_string());
        let data_sheet_id = workbook.sheets[0].id.clone();
        let data_sheet_name = workbook.sheets[0].name.clone();
        let pivot_sheet_id = workbook.sheets[1].id.clone();

        let date_serial =
            ymd_to_serial(ExcelDate::new(2024, 1, 15), ExcelDateSystem::EXCEL_1900).unwrap() as f64;

        let sheet = workbook.sheet_mut(&data_sheet_id).unwrap();
        sheet.set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Text("Date".to_string()))),
        );
        sheet.set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Text("Amount".to_string()))),
        );
        sheet.set_cell(1, 0, Cell::from_literal(Some(CellScalar::Number(date_serial))));
        sheet.set_cell(1, 1, Cell::from_literal(Some(CellScalar::Number(10.0))));

        // Intentionally omit per-cell number formats: the engine column style should drive date
        // typing for pivot cache building.
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let date_style = state.engine.intern_style(Style {
            number_format: Some("m/d/yyyy".to_string()),
            ..Style::default()
        });
        state
            .engine
            .set_col_style_id(&data_sheet_name, 0, Some(date_style));

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Date")],
            column_fields: Vec::new(),
            value_fields: vec![ValueField {
                source_field: "Amount".into(),
                name: "Sum of Amount".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: Vec::new(),
            calculated_fields: Vec::new(),
            calculated_items: Vec::new(),
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        state
            .create_pivot_table(
                "By Date".to_string(),
                data_sheet_id,
                CellRect {
                    start_row: 0,
                    start_col: 0,
                    end_row: 1,
                    end_col: 1,
                },
                PivotDestination {
                    sheet_id: pivot_sheet_id.clone(),
                    row: 0,
                    col: 0,
                },
                cfg,
            )
            .unwrap();

        let label = state.get_cell(&pivot_sheet_id, 1, 0).unwrap();
        assert_eq!(label.value, CellScalar::Number(date_serial));
        assert_eq!(label.display_value, "1/15/2024".to_string());

        let workbook = state.workbook.as_ref().unwrap();
        let pivot_sheet = workbook.sheet(&pivot_sheet_id).unwrap();
        assert_eq!(pivot_sheet.get_cell(1, 0).number_format.as_deref(), Some("m/d/yyyy"));
    }

    #[test]
    fn pivot_date_labels_respect_excel_1904_date_system() {
        use formula_engine::date::{ymd_to_serial, ExcelDate, ExcelDateSystem};

        let mut workbook = Workbook::new_empty(None);
        workbook.date_system = formula_model::DateSystem::Excel1904;
        workbook.add_sheet("Data".to_string());
        workbook.add_sheet("Pivot".to_string());
        let data_sheet_id = workbook.sheets[0].id.clone();
        let pivot_sheet_id = workbook.sheets[1].id.clone();

        let date1_serial =
            ymd_to_serial(ExcelDate::new(2024, 1, 15), ExcelDateSystem::Excel1904).unwrap() as f64;
        let date2_serial =
            ymd_to_serial(ExcelDate::new(2024, 2, 1), ExcelDateSystem::Excel1904).unwrap() as f64;

        let sheet = workbook.sheet_mut(&data_sheet_id).unwrap();
        sheet.set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Text("Date".to_string()))),
        );
        sheet.set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Text("Amount".to_string()))),
        );

        let mut date_cell_1 = Cell::from_literal(Some(CellScalar::Number(date1_serial)));
        date_cell_1.number_format = Some("m/d/yyyy".to_string());
        sheet.set_cell(1, 0, date_cell_1);
        sheet.set_cell(1, 1, Cell::from_literal(Some(CellScalar::Number(10.0))));

        let mut date_cell_2 = Cell::from_literal(Some(CellScalar::Number(date2_serial)));
        date_cell_2.number_format = Some("m/d/yyyy".to_string());
        sheet.set_cell(2, 0, date_cell_2);
        sheet.set_cell(2, 1, Cell::from_literal(Some(CellScalar::Number(20.0))));

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let cfg = PivotConfig {
            row_fields: vec![PivotField::new("Date")],
            column_fields: Vec::new(),
            value_fields: vec![ValueField {
                source_field: PivotFieldRef::CacheFieldName("Amount".to_string()),
                name: "Sum of Amount".to_string(),
                aggregation: AggregationType::Sum,
                number_format: None,
                show_as: None,
                base_field: None,
                base_item: None,
            }],
            filter_fields: Vec::new(),
            calculated_fields: Vec::new(),
            calculated_items: Vec::new(),
            layout: Layout::Tabular,
            subtotals: SubtotalPosition::None,
            grand_totals: GrandTotals {
                rows: true,
                columns: false,
            },
        };

        state
            .create_pivot_table(
                "By Date".to_string(),
                data_sheet_id,
                CellRect {
                    start_row: 0,
                    start_col: 0,
                    end_row: 2,
                    end_col: 1,
                },
                PivotDestination {
                    sheet_id: pivot_sheet_id.clone(),
                    row: 0,
                    col: 0,
                },
                cfg,
            )
            .unwrap();

        // Same behavior as the Excel 1900 system: labels are numeric serials + formatting.
        let label_1 = state.get_cell(&pivot_sheet_id, 1, 0).unwrap();
        assert_eq!(label_1.value, CellScalar::Number(date1_serial));
        assert_eq!(label_1.display_value, "1/15/2024".to_string());
        let label_2 = state.get_cell(&pivot_sheet_id, 2, 0).unwrap();
        assert_eq!(label_2.value, CellScalar::Number(date2_serial));
        assert_eq!(label_2.display_value, "2/1/2024".to_string());

        let workbook = state.workbook.as_ref().unwrap();
        let pivot_sheet = workbook.sheet(&pivot_sheet_id).unwrap();
        assert_eq!(pivot_sheet.get_cell(1, 0).number_format.as_deref(), Some("m/d/yyyy"));
        assert_eq!(pivot_sheet.get_cell(2, 0).number_format.as_deref(), Some("m/d/yyyy"));
    }

    #[test]
    fn pivot_auto_refreshes_when_source_cell_changes() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Data".to_string());
        workbook.add_sheet("Pivot".to_string());
        let data_sheet_id = workbook.sheets[0].id.clone();
        let pivot_sheet_id = workbook.sheets[1].id.clone();

        let sheet = workbook.sheet_mut(&data_sheet_id).unwrap();
        sheet.set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Text("Region".to_string()))),
        );
        sheet.set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Text("Sales".to_string()))),
        );
        sheet.set_cell(
            1,
            0,
            Cell::from_literal(Some(CellScalar::Text("East".to_string()))),
        );
        sheet.set_cell(1, 1, Cell::from_literal(Some(CellScalar::Number(100.0))));
        sheet.set_cell(
            2,
            0,
            Cell::from_literal(Some(CellScalar::Text("West".to_string()))),
        );
        sheet.set_cell(2, 1, Cell::from_literal(Some(CellScalar::Number(200.0))));

        let mut state = AppState::new();
        state.load_workbook(workbook);
        state
            .create_pivot_table(
                "Sales by Region".to_string(),
                data_sheet_id.clone(),
                CellRect {
                    start_row: 0,
                    start_col: 0,
                    end_row: 2,
                    end_col: 1,
                },
                PivotDestination {
                    sheet_id: pivot_sheet_id.clone(),
                    row: 0,
                    col: 0,
                },
                simple_pivot_config(),
            )
            .unwrap();

        // Change East sales from 100 -> 120; pivot should refresh.
        state
            .set_cell(&data_sheet_id, 1, 1, Some(JsonValue::from(120)), None)
            .unwrap();

        assert_eq!(
            state.get_cell(&pivot_sheet_id, 1, 1).unwrap().value,
            CellScalar::Number(120.0)
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 3, 1).unwrap().value,
            CellScalar::Number(320.0)
        );
    }

    #[test]
    fn pivot_refreshes_when_formula_in_source_range_updates() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Data".to_string());
        workbook.add_sheet("Pivot".to_string());
        let data_sheet_id = workbook.sheets[0].id.clone();
        let pivot_sheet_id = workbook.sheets[1].id.clone();

        let sheet = workbook.sheet_mut(&data_sheet_id).unwrap();
        sheet.set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Text("Region".to_string()))),
        );
        sheet.set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Text("Sales".to_string()))),
        );

        // Sales values are formulas that depend on C2/C3 (outside the pivot source range).
        sheet.set_cell(
            1,
            0,
            Cell::from_literal(Some(CellScalar::Text("East".to_string()))),
        );
        sheet.set_cell(1, 1, Cell::from_formula("=C2".to_string()));
        sheet.set_cell(1, 2, Cell::from_literal(Some(CellScalar::Number(100.0))));

        sheet.set_cell(
            2,
            0,
            Cell::from_literal(Some(CellScalar::Text("West".to_string()))),
        );
        sheet.set_cell(2, 1, Cell::from_formula("=C3".to_string()));
        sheet.set_cell(2, 2, Cell::from_literal(Some(CellScalar::Number(200.0))));

        let mut state = AppState::new();
        state.load_workbook(workbook);
        state
            .create_pivot_table(
                "Sales by Region".to_string(),
                data_sheet_id.clone(),
                CellRect {
                    start_row: 0,
                    start_col: 0,
                    end_row: 2,
                    end_col: 1,
                },
                PivotDestination {
                    sheet_id: pivot_sheet_id.clone(),
                    row: 0,
                    col: 0,
                },
                simple_pivot_config(),
            )
            .unwrap();

        // Update C2 (outside source range). B2 should update (formula), and that should trigger pivot refresh.
        state
            .set_cell(&data_sheet_id, 1, 2, Some(JsonValue::from(150)), None)
            .unwrap();

        assert_eq!(
            state.get_cell(&pivot_sheet_id, 1, 1).unwrap().value,
            CellScalar::Number(150.0)
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 3, 1).unwrap().value,
            CellScalar::Number(350.0)
        );
    }

    #[test]
    fn pivot_shrinking_clears_stale_cells() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Data".to_string());
        workbook.add_sheet("Pivot".to_string());
        let data_sheet_id = workbook.sheets[0].id.clone();
        let pivot_sheet_id = workbook.sheets[1].id.clone();

        let sheet = workbook.sheet_mut(&data_sheet_id).unwrap();
        sheet.set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Text("Region".to_string()))),
        );
        sheet.set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Text("Sales".to_string()))),
        );
        sheet.set_cell(
            1,
            0,
            Cell::from_literal(Some(CellScalar::Text("East".to_string()))),
        );
        sheet.set_cell(1, 1, Cell::from_literal(Some(CellScalar::Number(100.0))));
        sheet.set_cell(
            2,
            0,
            Cell::from_literal(Some(CellScalar::Text("West".to_string()))),
        );
        sheet.set_cell(2, 1, Cell::from_literal(Some(CellScalar::Number(200.0))));

        let mut state = AppState::new();
        state.load_workbook(workbook);
        state
            .create_pivot_table(
                "Sales by Region".to_string(),
                data_sheet_id.clone(),
                CellRect {
                    start_row: 0,
                    start_col: 0,
                    end_row: 2,
                    end_col: 1,
                },
                PivotDestination {
                    sheet_id: pivot_sheet_id.clone(),
                    row: 0,
                    col: 0,
                },
                simple_pivot_config(),
            )
            .unwrap();

        // Ensure the "West" row exists before shrinking.
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 2, 0).unwrap().value,
            CellScalar::Text("West".to_string())
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 3, 0).unwrap().value,
            CellScalar::Text("Grand Total".to_string())
        );

        // Collapse West into East; pivot should now have only 3 rows (header, East, Grand Total).
        state
            .set_cell(&data_sheet_id, 2, 0, Some(JsonValue::from("East")), None)
            .unwrap();

        assert_eq!(
            state.get_cell(&pivot_sheet_id, 2, 0).unwrap().value,
            CellScalar::Text("Grand Total".to_string())
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 3, 0).unwrap().value,
            CellScalar::Empty
        );
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 3, 1).unwrap().value,
            CellScalar::Empty
        );
    }

    #[test]
    fn pivot_refresh_after_shrink_does_not_clear_cells_outside_latest_output() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Data".to_string());
        workbook.add_sheet("Pivot".to_string());
        let data_sheet_id = workbook.sheets[0].id.clone();
        let pivot_sheet_id = workbook.sheets[1].id.clone();

        let sheet = workbook.sheet_mut(&data_sheet_id).unwrap();
        sheet.set_cell(
            0,
            0,
            Cell::from_literal(Some(CellScalar::Text("Region".to_string()))),
        );
        sheet.set_cell(
            0,
            1,
            Cell::from_literal(Some(CellScalar::Text("Sales".to_string()))),
        );
        sheet.set_cell(
            1,
            0,
            Cell::from_literal(Some(CellScalar::Text("East".to_string()))),
        );
        sheet.set_cell(1, 1, Cell::from_literal(Some(CellScalar::Number(100.0))));
        sheet.set_cell(
            2,
            0,
            Cell::from_literal(Some(CellScalar::Text("West".to_string()))),
        );
        sheet.set_cell(2, 1, Cell::from_literal(Some(CellScalar::Number(200.0))));

        let mut state = AppState::new();
        state.load_workbook(workbook);
        state
            .create_pivot_table(
                "Sales by Region".to_string(),
                data_sheet_id.clone(),
                CellRect {
                    start_row: 0,
                    start_col: 0,
                    end_row: 2,
                    end_col: 1,
                },
                PivotDestination {
                    sheet_id: pivot_sheet_id.clone(),
                    row: 0,
                    col: 0,
                },
                simple_pivot_config(),
            )
            .unwrap();

        // Shrink the pivot output by collapsing West into East.
        state
            .set_cell(&data_sheet_id, 2, 0, Some(JsonValue::from("East")), None)
            .unwrap();
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 3, 0).unwrap().value,
            CellScalar::Empty
        );

        // User edits a cell that used to be part of the pivot output, but is now outside the
        // shrunken range. Subsequent refreshes should not clear it unless the pivot grows back.
        state
            .set_cell(&pivot_sheet_id, 3, 0, Some(JsonValue::from("Note")), None)
            .unwrap();
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 3, 0).unwrap().value,
            CellScalar::Text("Note".to_string())
        );

        // Trigger another pivot refresh by changing a source value.
        state
            .set_cell(&data_sheet_id, 1, 1, Some(JsonValue::from(110)), None)
            .unwrap();

        // The note should survive because it's outside the most recently rendered pivot range.
        assert_eq!(
            state.get_cell(&pivot_sheet_id, 3, 0).unwrap().value,
            CellScalar::Text("Note".to_string())
        );
    }
}
