use crate::file_io::Workbook;
use crate::macros::{
    execute_invocation, MacroExecutionOptions, MacroExecutionOutcome, MacroHost, MacroHostError,
    MacroInfo, MacroInvocation,
};
use formula_columnar::{ColumnType as ColumnarType, ColumnarTable, Value as ColumnarValue};
use formula_engine::eval::{parse_a1, CellAddr};
use formula_engine::what_if::goal_seek::{GoalSeek, GoalSeekParams, GoalSeekResult};
use formula_engine::what_if::monte_carlo::{MonteCarloEngine, SimulationConfig, SimulationResult};
use formula_engine::what_if::scenario_manager::{
    Scenario, ScenarioId, ScenarioManager, SummaryReport,
};
use formula_engine::what_if::{
    CellRef as WhatIfCellRef, CellValue as WhatIfCellValue, EngineWhatIfModel, WhatIfModel,
};
use formula_engine::{
    Engine as FormulaEngine, ErrorKind, ExternalValueProvider, RecalcMode, Value as EngineValue,
};
use formula_xlsx::print::{
    CellRange as PrintCellRange, ManualPageBreaks, PageSetup, SheetPrintSettings,
};
use serde_json::Value as JsonValue;
use std::collections::HashMap;
use std::sync::Arc;

#[derive(Debug, thiserror::Error)]
pub enum AppStateError {
    #[error("no workbook loaded")]
    NoWorkbookLoaded,
    #[error("no undo history")]
    NoUndoHistory,
    #[error("no redo history")]
    NoRedoHistory,
    #[error("unknown sheet id: {0}")]
    UnknownSheet(String),
    #[error("invalid range: start ({start_row},{start_col}) end ({end_row},{end_col})")]
    InvalidRange {
        start_row: usize,
        start_col: usize,
        end_row: usize,
        end_col: usize,
    },
    #[error("what-if analysis failed: {0}")]
    WhatIf(String),
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
            CellScalar::Number(n) => {
                if (n.fract() - 0.0).abs() < f64::EPSILON {
                    format!("{:.0}", n)
                } else {
                    n.to_string()
                }
            }
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
}

impl Cell {
    pub(crate) fn empty() -> Self {
        Self {
            input_value: None,
            formula: None,
            computed_value: CellScalar::Empty,
        }
    }

    pub(crate) fn from_literal(value: Option<CellScalar>) -> Self {
        let computed_value = value.clone().unwrap_or(CellScalar::Empty);
        Self {
            input_value: value,
            formula: None,
            computed_value,
        }
    }

    pub(crate) fn from_formula(formula: String) -> Self {
        Self {
            input_value: None,
            formula: Some(formula),
            computed_value: CellScalar::Empty,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct CellData {
    pub value: CellScalar,
    pub formula: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct CellUpdateData {
    pub sheet_id: String,
    pub row: usize,
    pub col: usize,
    pub value: CellScalar,
    pub formula: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct SheetInfoData {
    pub id: String,
    pub name: String,
}

#[derive(Clone, Debug, PartialEq)]
pub struct WorkbookInfoData {
    pub path: Option<String>,
    pub origin_path: Option<String>,
    pub sheets: Vec<SheetInfoData>,
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

pub struct AppState {
    workbook: Option<Workbook>,
    engine: FormulaEngine,
    dirty: bool,
    undo_stack: Vec<UndoEntry>,
    redo_stack: Vec<UndoEntry>,
    scenario_manager: ScenarioManager,
    macro_host: MacroHost,
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
            engine: FormulaEngine::new(),
            dirty: false,
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
            scenario_manager: ScenarioManager::new(),
            macro_host: MacroHost::default(),
        }
    }

    pub fn has_unsaved_changes(&self) -> bool {
        self.dirty
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
                })
                .collect(),
        })
    }

    pub fn load_workbook(&mut self, mut workbook: Workbook) -> WorkbookInfoData {
        workbook.ensure_sheet_ids();
        self.workbook = Some(workbook);
        self.engine = FormulaEngine::new();
        self.scenario_manager = ScenarioManager::new();
        self.macro_host.invalidate();

        // Best effort: rebuild and calculate. Unsupported formulas become #NAME? via the engine.
        let _ = self.rebuild_engine_from_workbook();
        self.engine.recalculate();
        let _ = self.refresh_computed_values();

        self.dirty = false;
        self.undo_stack.clear();
        self.redo_stack.clear();
        self.workbook_info()
            .expect("workbook_info should succeed right after load")
    }

    pub fn mark_saved(
        &mut self,
        new_path: Option<String>,
        new_origin_xlsx_bytes: Option<Arc<[u8]>>,
    ) -> Result<(), AppStateError> {
        let workbook = self
            .workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        if let Some(path) = new_path {
            workbook.path = Some(path);
        }
        if let Some(bytes) = new_origin_xlsx_bytes {
            workbook.origin_xlsx_bytes = Some(bytes);
        }
        // Saving establishes a new baseline for "net" changes. Clear the per-cell baseline so
        // subsequent edits are tracked against this saved state (not the previously opened or
        // previously saved workbook bytes).
        workbook.cell_input_baseline.clear();
        workbook.original_print_settings = workbook.print_settings.clone();
        for sheet in &mut workbook.sheets {
            sheet.clear_dirty_cells();
        }

        // If the saved file is `.xlsx`, macros are not preserved; clear any in-memory macro
        // payloads so the UI doesn't continue to treat the workbook as macro-enabled.
        if workbook
            .path
            .as_deref()
            .and_then(|p| std::path::Path::new(p).extension().and_then(|s| s.to_str()))
            .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx"))
        {
            workbook.vba_project_bin = None;
            workbook.macro_fingerprint = None;
        }
        self.dirty = false;
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
            .find(|s| s.sheet_name.eq_ignore_ascii_case(&sheet.name));

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
        let sheet = workbook
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
        let cell = sheet.get_cell(row, col);
        let addr = coord_to_a1(row, col);
        let value = engine_value_to_scalar(self.engine.get_cell_value(&sheet.name, &addr));
        Ok(CellData {
            value,
            formula: cell.formula,
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

        let workbook = self
            .workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        let sheet = workbook
            .sheet(sheet_id)
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()))?;
        let mut rows = Vec::with_capacity(end_row - start_row + 1);
        for r in start_row..=end_row {
            let mut row_out = Vec::with_capacity(end_col - start_col + 1);
            for c in start_col..=end_col {
                let cell = sheet.get_cell(r, c);
                let addr = coord_to_a1(r, c);
                let value = engine_value_to_scalar(self.engine.get_cell_value(&sheet.name, &addr));
                row_out.push(CellData {
                    value,
                    formula: cell.formula,
                });
            }
            rows.push(row_out);
        }
        Ok(rows)
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
            if workbook.origin_xlsx_bytes.is_some() {
                workbook
                    .cell_input_baseline
                    .entry((before.sheet_id.clone(), before.row, before.col))
                    .or_insert_with(|| (before.value.clone(), before.formula.clone()));
            }
        }

        self.apply_snapshots(&[after_cell.clone()])?;
        self.engine.recalculate();
        let mut updates = self.refresh_computed_values()?;

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

        let expected_rows = end_row - start_row + 1;
        let expected_cols = end_col - start_col + 1;
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
            if workbook.origin_xlsx_bytes.is_some() {
                for snap in &before {
                    workbook
                        .cell_input_baseline
                        .entry((snap.sheet_id.clone(), snap.row, snap.col))
                        .or_insert_with(|| (snap.value.clone(), snap.formula.clone()));
                }
            }
        }

        self.apply_snapshots(&after)?;
        self.engine.recalculate();
        let mut updates = self.refresh_computed_values()?;

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
                });
            }
        }

        self.dirty = true;
        self.redo_stack.clear();
        self.undo_stack.push(UndoEntry { before, after });
        Ok(updates)
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

        self.engine.recalculate();
        self.refresh_computed_values()
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
            if workbook.origin_xlsx_bytes.is_some() {
                for snap in &before {
                    workbook
                        .cell_input_baseline
                        .entry((snap.sheet_id.clone(), snap.row, snap.col))
                        .or_insert_with(|| (snap.value.clone(), snap.formula.clone()));
                }
            }
        }

        self.apply_snapshots(&after)?;
        self.engine.recalculate();
        let mut updates = self.refresh_computed_values()?;

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
            if workbook.origin_xlsx_bytes.is_some() {
                for snap in &before {
                    workbook
                        .cell_input_baseline
                        .entry((snap.sheet_id.clone(), snap.row, snap.col))
                        .or_insert_with(|| (snap.value.clone(), snap.formula.clone()));
                }
            }
        }

        self.apply_snapshots(&after)?;
        self.engine.recalculate();
        let mut updates = self.refresh_computed_values()?;

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
        self.apply_snapshots(&entry.before)?;
        self.engine.recalculate();
        let mut updates = self.refresh_computed_values()?;

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
        self.apply_snapshots(&entry.after)?;
        self.engine.recalculate();
        let mut updates = self.refresh_computed_values()?;

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

            let new_cell = match (&snap.formula, &snap.value) {
                (Some(formula), _) => Cell::from_formula(formula.clone()),
                (None, Some(value)) => Cell::from_literal(Some(value.clone())),
                (None, None) => Cell::empty(),
            };

            sheet.set_cell(snap.row, snap.col, new_cell);
            apply_snapshot_to_engine(
                engine,
                &sheet_name,
                snap.row,
                snap.col,
                &snap.value,
                &snap.formula,
            );
        }

        Ok(())
    }

    fn rebuild_engine_from_workbook(&mut self) -> Result<(), AppStateError> {
        let workbook = self
            .workbook
            .as_ref()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
        self.engine = FormulaEngine::new();
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

        for sheet in &workbook.sheets {
            let sheet_name = &sheet.name;
            for ((row, col), cell) in sheet.cells_iter() {
                let addr = coord_to_a1(row, col);
                if let Some(formula) = &cell.formula {
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

        Ok(())
    }

    fn refresh_computed_values(&mut self) -> Result<Vec<CellUpdateData>, AppStateError> {
        let workbook = self
            .workbook
            .as_mut()
            .ok_or(AppStateError::NoWorkbookLoaded)?;
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
                    updates.push(CellUpdateData {
                        sheet_id: sheet_id.clone(),
                        row: *row,
                        col: *col,
                        value: new_value,
                        formula: cell.formula.clone(),
                    });
                }
            }
        }

        Ok(updates)
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

    let sheet_id = if sheet_name.eq_ignore_ascii_case(default_sheet_name) {
        default_sheet_id.to_string()
    } else {
        workbook
            .sheets
            .iter()
            .find(|s| s.name.eq_ignore_ascii_case(&sheet_name))
            .map(|s| s.id.clone())
            .ok_or_else(|| AppStateError::UnknownSheet(sheet_name.clone()))?
    };

    let addr = parse_a1(&addr).map_err(|e| AppStateError::WhatIf(e.to_string()))?;
    Ok((sheet_id, addr.row as usize, addr.col as usize))
}

fn normalize_formula(formula: Option<String>) -> Option<String> {
    let mut formula = formula?.trim().to_string();
    if formula.is_empty() {
        return None;
    }
    if !formula.starts_with('=') {
        formula.insert(0, '=');
    }
    Some(formula)
}

fn coord_to_a1(row: usize, col: usize) -> String {
    format!("{}{}", col_index_to_letters(col), row + 1)
}

fn quote_sheet_name(name: &str) -> String {
    // Excel escapes single quotes inside a quoted sheet name by doubling them.
    let escaped = name.replace('\'', "''");
    format!("'{escaped}'")
}

fn col_index_to_letters(mut col: usize) -> String {
    // Excel columns are base-26 with A=1..Z=26.
    col += 1;
    let mut letters = Vec::new();
    while col > 0 {
        let rem = (col - 1) % 26;
        letters.push((b'A' + rem as u8) as char);
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
        EngineValue::Bool(b) => CellScalar::Bool(b),
        EngineValue::Error(e) => CellScalar::Error(e.as_code().to_string()),
        EngineValue::Array(arr) => engine_value_to_scalar(arr.top_left()),
        EngineValue::Spill { .. } => CellScalar::Error("#SPILL!".to_string()),
    }
}

fn parse_error_kind(value: &str) -> Option<ErrorKind> {
    match value.trim() {
        "#NULL!" | "Null" => Some(ErrorKind::Null),
        "#DIV/0!" | "Div0" => Some(ErrorKind::Div0),
        "#VALUE!" | "Value" => Some(ErrorKind::Value),
        "#REF!" | "Ref" => Some(ErrorKind::Ref),
        "#NAME?" | "Name" => Some(ErrorKind::Name),
        "#NUM!" | "Num" => Some(ErrorKind::Num),
        "#N/A" | "#N/A!" | "NA" => Some(ErrorKind::NA),
        "#SPILL!" | "Spill" => Some(ErrorKind::Spill),
        "#CALC!" | "Calc" => Some(ErrorKind::Calc),
        _ => None,
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
        .position(|s| s.sheet_name.eq_ignore_ascii_case(sheet_name))
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::file_io::{read_xlsx_blocking, write_xlsx_blocking};
    use formula_engine::what_if::monte_carlo::{Distribution, InputDistribution};
    use formula_model::import::{import_csv_to_columnar_table, CsvOptions};

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

        workbook
            .sheet_mut(&sheet_id)
            .unwrap()
            .set_cell(0, 1, Cell::from_formula("=SUM(A1:A3)".to_string()));

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
    fn col_index_to_letters_matches_excel_a1() {
        assert_eq!(col_index_to_letters(0), "A");
        assert_eq!(col_index_to_letters(25), "Z");
        assert_eq!(col_index_to_letters(26), "AA");
        assert_eq!(col_index_to_letters(27), "AB");
        assert_eq!(col_index_to_letters(701), "ZZ");
        assert_eq!(col_index_to_letters(702), "AAA");
    }
}
