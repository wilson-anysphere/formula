use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PrintCellRange {
    pub start_row: u32,
    pub end_row: u32,
    pub start_col: u32,
    pub end_col: u32,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PrintRowRange {
    pub start: u32,
    pub end: u32,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PrintColRange {
    pub start: u32,
    pub end: u32,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PrintTitles {
    pub repeat_rows: Option<PrintRowRange>,
    pub repeat_cols: Option<PrintColRange>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "lowercase")]
pub enum PageOrientation {
    Portrait,
    Landscape,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PageMargins {
    pub left: f64,
    pub right: f64,
    pub top: f64,
    pub bottom: f64,
    pub header: f64,
    pub footer: f64,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(tag = "kind")]
pub enum PageScaling {
    #[serde(rename = "percent")]
    Percent { percent: u16 },
    #[serde(rename = "fitTo")]
    FitTo { width_pages: u16, height_pages: u16 },
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PageSetup {
    pub orientation: PageOrientation,
    pub paper_size: u16,
    pub margins: PageMargins,
    pub scaling: PageScaling,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct ManualPageBreaks {
    pub row_breaks_after: Vec<u32>,
    pub col_breaks_after: Vec<u32>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct SheetPrintSettings {
    pub sheet_name: String,
    pub print_area: Option<Vec<PrintCellRange>>,
    pub print_titles: Option<PrintTitles>,
    pub page_setup: PageSetup,
    pub manual_page_breaks: ManualPageBreaks,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct CellValue {
    pub value: Option<JsonValue>,
    pub formula: Option<String>,
    pub display_value: String,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct CellUpdate {
    pub sheet_id: String,
    pub row: usize,
    pub col: usize,
    pub value: Option<JsonValue>,
    pub formula: Option<String>,
    pub display_value: String,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct RangeData {
    pub values: Vec<Vec<CellValue>>,
    pub start_row: usize,
    pub start_col: usize,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct SheetInfo {
    pub id: String,
    pub name: String,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct WorkbookInfo {
    pub path: Option<String>,
    pub origin_path: Option<String>,
    pub sheets: Vec<SheetInfo>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct RangeCellEdit {
    pub value: Option<JsonValue>,
    pub formula: Option<String>,
}

#[cfg(feature = "desktop")]
use crate::file_io::{read_xlsx, write_xlsx};
#[cfg(feature = "desktop")]
use crate::state::{AppState, AppStateError, CellUpdateData, SharedAppState};
#[cfg(feature = "desktop")]
use std::path::PathBuf;
#[cfg(feature = "desktop")]
use tauri::State;

#[cfg(feature = "desktop")]
fn app_error(err: AppStateError) -> String {
    err.to_string()
}

#[cfg(feature = "desktop")]
fn coerce_save_path_to_xlsx(path: &str) -> String {
    let mut buf = PathBuf::from(path);
    if buf
        .extension()
        .and_then(|s| s.to_str())
        .is_some_and(|ext| ext.eq_ignore_ascii_case("xls"))
    {
        buf.set_extension("xlsx");
        return buf.to_string_lossy().to_string();
    }

    path.to_string()
}

#[cfg(feature = "desktop")]
fn cell_value_from_state(
    state: &AppState,
    sheet_id: &str,
    row: usize,
    col: usize,
) -> Result<CellValue, String> {
    let cell = state.get_cell(sheet_id, row, col).map_err(app_error)?;
    Ok(CellValue {
        value: cell.value.as_json(),
        formula: cell.formula,
        display_value: cell.value.display(),
    })
}

#[cfg(feature = "desktop")]
fn cell_update_from_state(update: CellUpdateData) -> CellUpdate {
    CellUpdate {
        sheet_id: update.sheet_id,
        row: update.row,
        col: update.col,
        value: update.value.as_json(),
        formula: update.formula,
        display_value: update.value.display(),
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn open_workbook(
    path: String,
    state: State<'_, SharedAppState>,
) -> Result<WorkbookInfo, String> {
    let workbook = read_xlsx(path.clone()).await.map_err(|e| e.to_string())?;
    let shared = state.inner().clone();
    let info = tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        state.load_workbook(workbook)
    })
    .await
    .map_err(|e| e.to_string())?;

    Ok(WorkbookInfo {
        path: info.path,
        origin_path: info.origin_path,
        sheets: info
            .sheets
            .into_iter()
            .map(|s| SheetInfo {
                id: s.id,
                name: s.name,
            })
            .collect(),
    })
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn save_workbook(
    path: Option<String>,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let (save_path, workbook) = {
        let state = state.inner().lock().unwrap();
        let workbook = state.get_workbook().map_err(app_error)?.clone();
        let save_path = path
            .clone()
            .or_else(|| workbook.path.clone())
            .ok_or_else(|| "no save path provided".to_string())?;
        (save_path, workbook)
    };

    let save_path = coerce_save_path_to_xlsx(&save_path);

    write_xlsx(save_path.clone(), workbook)
        .await
        .map_err(|e| e.to_string())?;

    {
        let mut state = state.inner().lock().unwrap();
        state.mark_saved(Some(save_path)).map_err(app_error)?;
    }

    Ok(())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_cell(
    sheet_id: String,
    row: usize,
    col: usize,
    state: State<'_, SharedAppState>,
) -> Result<CellValue, String> {
    let state = state.inner().lock().unwrap();
    cell_value_from_state(&state, &sheet_id, row, col)
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn set_cell(
    sheet_id: String,
    row: usize,
    col: usize,
    value: Option<JsonValue>,
    formula: Option<String>,
    state: State<'_, SharedAppState>,
) -> Result<Vec<CellUpdate>, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let updates = state
            .set_cell(&sheet_id, row, col, value, formula)
            .map_err(app_error)?;
        Ok::<_, String>(updates.into_iter().map(cell_update_from_state).collect())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_range(
    sheet_id: String,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    state: State<'_, SharedAppState>,
) -> Result<RangeData, String> {
    let state = state.inner().lock().unwrap();
    let cells = state
        .get_range(&sheet_id, start_row, start_col, end_row, end_col)
        .map_err(app_error)?;
    let values = cells
        .into_iter()
        .map(|row| {
            row.into_iter()
                .map(|cell| CellValue {
                    value: cell.value.as_json(),
                    formula: cell.formula,
                    display_value: cell.value.display(),
                })
                .collect::<Vec<_>>()
        })
        .collect();

    Ok(RangeData {
        values,
        start_row,
        start_col,
    })
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn set_range(
    sheet_id: String,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    values: Vec<Vec<RangeCellEdit>>,
    state: State<'_, SharedAppState>,
) -> Result<Vec<CellUpdate>, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let normalized = values
            .into_iter()
            .map(|row| {
                row.into_iter()
                    .map(|c| (c.value, c.formula))
                    .collect::<Vec<_>>()
            })
            .collect::<Vec<_>>();
        let updates = state
            .set_range(
                &sheet_id, start_row, start_col, end_row, end_col, normalized,
            )
            .map_err(app_error)?;
        Ok::<_, String>(updates.into_iter().map(cell_update_from_state).collect())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn recalculate(state: State<'_, SharedAppState>) -> Result<Vec<CellUpdate>, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let updates = state.recalculate_all().map_err(app_error)?;
        Ok::<_, String>(updates.into_iter().map(cell_update_from_state).collect())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn undo(state: State<'_, SharedAppState>) -> Result<Vec<CellUpdate>, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let updates = state.undo().map_err(app_error)?;
        Ok::<_, String>(updates.into_iter().map(cell_update_from_state).collect())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn redo(state: State<'_, SharedAppState>) -> Result<Vec<CellUpdate>, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let updates = state.redo().map_err(app_error)?;
        Ok::<_, String>(updates.into_iter().map(cell_update_from_state).collect())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
fn to_core_page_setup(setup: &PageSetup) -> formula_xlsx::print::PageSetup {
    use formula_xlsx::print as core;

    let orientation = match setup.orientation {
        PageOrientation::Portrait => core::Orientation::Portrait,
        PageOrientation::Landscape => core::Orientation::Landscape,
    };

    let margins = core::PageMargins {
        left: setup.margins.left,
        right: setup.margins.right,
        top: setup.margins.top,
        bottom: setup.margins.bottom,
        header: setup.margins.header,
        footer: setup.margins.footer,
    };

    let scaling = match setup.scaling {
        PageScaling::Percent { percent } => core::Scaling::Percent(percent),
        PageScaling::FitTo {
            width_pages,
            height_pages,
        } => core::Scaling::FitTo {
            width: width_pages,
            height: height_pages,
        },
    };

    core::PageSetup {
        orientation,
        paper_size: core::PaperSize {
            code: setup.paper_size,
        },
        margins,
        scaling,
    }
}

#[cfg(feature = "desktop")]
fn from_core_page_setup(setup: &formula_xlsx::print::PageSetup) -> PageSetup {
    let orientation = match setup.orientation {
        formula_xlsx::print::Orientation::Portrait => PageOrientation::Portrait,
        formula_xlsx::print::Orientation::Landscape => PageOrientation::Landscape,
    };

    let scaling = match setup.scaling {
        formula_xlsx::print::Scaling::Percent(percent) => PageScaling::Percent { percent },
        formula_xlsx::print::Scaling::FitTo { width, height } => PageScaling::FitTo {
            width_pages: width,
            height_pages: height,
        },
    };

    PageSetup {
        orientation,
        paper_size: setup.paper_size.code,
        margins: PageMargins {
            left: setup.margins.left,
            right: setup.margins.right,
            top: setup.margins.top,
            bottom: setup.margins.bottom,
            header: setup.margins.header,
            footer: setup.margins.footer,
        },
        scaling,
    }
}

#[cfg(feature = "desktop")]
fn from_core_sheet_print_settings(
    settings: &formula_xlsx::print::SheetPrintSettings,
) -> SheetPrintSettings {
    SheetPrintSettings {
        sheet_name: settings.sheet_name.clone(),
        print_area: settings.print_area.as_ref().map(|ranges| {
            ranges
                .iter()
                .map(|r| PrintCellRange {
                    start_row: r.start_row,
                    end_row: r.end_row,
                    start_col: r.start_col,
                    end_col: r.end_col,
                })
                .collect()
        }),
        print_titles: settings.print_titles.as_ref().map(|t| PrintTitles {
            repeat_rows: t.repeat_rows.map(|r| PrintRowRange {
                start: r.start,
                end: r.end,
            }),
            repeat_cols: t.repeat_cols.map(|r| PrintColRange {
                start: r.start,
                end: r.end,
            }),
        }),
        page_setup: from_core_page_setup(&settings.page_setup),
        manual_page_breaks: ManualPageBreaks {
            row_breaks_after: settings
                .manual_page_breaks
                .row_breaks_after
                .iter()
                .copied()
                .collect(),
            col_breaks_after: settings
                .manual_page_breaks
                .col_breaks_after
                .iter()
                .copied()
                .collect(),
        },
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_sheet_print_settings(
    sheet_id: String,
    state: State<'_, SharedAppState>,
) -> Result<SheetPrintSettings, String> {
    let state = state.inner().lock().unwrap();
    let settings = state.sheet_print_settings(&sheet_id).map_err(app_error)?;
    Ok(from_core_sheet_print_settings(&settings))
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn set_sheet_page_setup(
    sheet_id: String,
    page_setup: PageSetup,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let mut state = state.inner().lock().unwrap();
    state
        .set_sheet_page_setup(&sheet_id, to_core_page_setup(&page_setup))
        .map_err(app_error)?;
    Ok(())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn set_sheet_print_area(
    sheet_id: String,
    print_area: Option<Vec<PrintCellRange>>,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let print_area = print_area.map(|ranges| {
        ranges
            .into_iter()
            .map(|r| formula_xlsx::print::CellRange {
                start_row: r.start_row,
                end_row: r.end_row,
                start_col: r.start_col,
                end_col: r.end_col,
            })
            .collect()
    });

    let mut state = state.inner().lock().unwrap();
    state
        .set_sheet_print_area(&sheet_id, print_area)
        .map_err(app_error)?;
    Ok(())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn export_sheet_range_pdf(
    sheet_id: String,
    range: PrintCellRange,
    col_widths_points: Option<Vec<f64>>,
    row_heights_points: Option<Vec<f64>>,
    state: State<'_, SharedAppState>,
) -> Result<String, String> {
    use base64::{engine::general_purpose::STANDARD, Engine as _};

    let state_guard = state.inner().lock().unwrap();
    let workbook = state_guard.get_workbook().map_err(app_error)?;
    let sheet = workbook
        .sheet(&sheet_id)
        .ok_or_else(|| app_error(AppStateError::UnknownSheet(sheet_id.clone())))?;

    let settings = state_guard
        .sheet_print_settings(&sheet_id)
        .map_err(app_error)?;

    let print_area = formula_xlsx::print::CellRange {
        start_row: range.start_row,
        end_row: range.end_row,
        start_col: range.start_col,
        end_col: range.end_col,
    };

    let mut col_widths = col_widths_points.unwrap_or_default();
    let mut row_heights = row_heights_points.unwrap_or_default();

    let needed_cols = print_area.end_col.max(1) as usize;
    let needed_rows = print_area.end_row.max(1) as usize;

    if col_widths.len() < needed_cols {
        col_widths.resize(needed_cols, 64.0);
    }
    if row_heights.len() < needed_rows {
        row_heights.resize(needed_rows, 20.0);
    }

    let pdf_bytes = formula_xlsx::print::export_range_to_pdf_bytes(
        &sheet.name,
        print_area,
        &col_widths,
        &row_heights,
        &settings.page_setup,
        &settings.manual_page_breaks,
        |row, col| {
            let value = workbook.cell_value(&sheet_id, (row - 1) as usize, (col - 1) as usize);
            let text = value.display();
            if text.is_empty() {
                None
            } else {
                Some(text)
            }
        },
    )
    .map_err(|e| e.to_string())?;

    Ok(STANDARD.encode(pdf_bytes))
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroInfo {
    pub id: String,
    pub name: String,
    pub language: String,
    pub module: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "snake_case")]
pub enum MacroPermission {
    FilesystemRead,
    FilesystemWrite,
    Network,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroError {
    pub message: String,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroRunResult {
    pub ok: bool,
    pub output: Vec<String>,
    pub updates: Vec<CellUpdate>,
    pub error: Option<MacroError>,
}

#[cfg(feature = "desktop")]
fn build_vba_program(state: &AppState) -> Result<formula_vba_runtime::VbaProgram, String> {
    let workbook = state.get_workbook().map_err(app_error)?;
    let Some(vba_bin) = workbook.vba_project_bin.as_ref() else {
        return Ok(formula_vba_runtime::VbaProgram::new());
    };

    let project = formula_vba::VBAProject::parse(vba_bin).map_err(|e| e.to_string())?;
    let combined = project
        .modules
        .iter()
        .map(|m| m.code.as_str())
        .collect::<Vec<_>>()
        .join("\n\n");

    formula_vba_runtime::parse_program(&combined).map_err(|e| e.to_string())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn list_macros(
    workbook_id: Option<String>,
    state: State<'_, SharedAppState>,
) -> Result<Vec<MacroInfo>, String> {
    let _ = workbook_id;
    let state = state.inner().lock().unwrap();
    let program = build_vba_program(&state)?;

    let mut macros = program
        .procedures
        .values()
        .map(|proc| MacroInfo {
            id: proc.name.clone(),
            name: proc.name.clone(),
            language: "vba".to_string(),
            module: None,
        })
        .collect::<Vec<_>>();
    macros.sort_by(|a, b| a.name.cmp(&b.name));
    Ok(macros)
}

#[cfg(feature = "desktop")]
struct AppStateSpreadsheet<'a> {
    state: &'a mut AppState,
    active_sheet: usize,
    active_cell: (u32, u32),
    output: Vec<String>,
    updates: Vec<CellUpdateData>,
}

#[cfg(feature = "desktop")]
impl<'a> AppStateSpreadsheet<'a> {
    fn new(state: &'a mut AppState) -> Result<Self, formula_vba_runtime::VbaError> {
        let workbook = state
            .get_workbook()
            .map_err(|e| formula_vba_runtime::VbaError::Runtime(e.to_string()))?;
        let active_sheet = if workbook.sheets.is_empty() { 0 } else { 0 };
        Ok(Self {
            state,
            active_sheet,
            active_cell: (1, 1),
            output: Vec::new(),
            updates: Vec::new(),
        })
    }

    fn sheet_id(&self, sheet: usize) -> Result<String, formula_vba_runtime::VbaError> {
        let workbook = self
            .state
            .get_workbook()
            .map_err(|e| formula_vba_runtime::VbaError::Runtime(e.to_string()))?;
        workbook
            .sheets
            .get(sheet)
            .map(|s| s.id.clone())
            .ok_or_else(|| {
                formula_vba_runtime::VbaError::Runtime(format!("Unknown sheet index: {sheet}"))
            })
    }

    fn cell_scalar_to_vba(value: crate::state::CellScalar) -> formula_vba_runtime::VbaValue {
        match value {
            crate::state::CellScalar::Empty => formula_vba_runtime::VbaValue::Empty,
            crate::state::CellScalar::Number(n) => formula_vba_runtime::VbaValue::Double(n),
            crate::state::CellScalar::Text(s) => formula_vba_runtime::VbaValue::String(s),
            crate::state::CellScalar::Bool(b) => formula_vba_runtime::VbaValue::Boolean(b),
            crate::state::CellScalar::Error(e) => formula_vba_runtime::VbaValue::String(e),
        }
    }

    fn vba_value_to_json(
        value: &formula_vba_runtime::VbaValue,
    ) -> Result<Option<JsonValue>, formula_vba_runtime::VbaError> {
        Ok(match value {
            formula_vba_runtime::VbaValue::Empty | formula_vba_runtime::VbaValue::Null => None,
            formula_vba_runtime::VbaValue::Boolean(b) => Some(JsonValue::from(*b)),
            formula_vba_runtime::VbaValue::Double(n) => Some(JsonValue::from(*n)),
            formula_vba_runtime::VbaValue::String(s) => Some(JsonValue::from(s.clone())),
            other => {
                return Err(formula_vba_runtime::VbaError::Runtime(format!(
                    "Unsupported VBA value for cell assignment: {other:?}"
                )))
            }
        })
    }

    fn take_output(&mut self) -> Vec<String> {
        std::mem::take(&mut self.output)
    }

    fn take_updates(&mut self) -> Vec<CellUpdateData> {
        std::mem::take(&mut self.updates)
    }
}

#[cfg(feature = "desktop")]
impl formula_vba_runtime::Spreadsheet for AppStateSpreadsheet<'_> {
    fn sheet_count(&self) -> usize {
        self.state
            .get_workbook()
            .ok()
            .map(|w| w.sheets.len())
            .unwrap_or(0)
    }

    fn sheet_name(&self, sheet: usize) -> Option<&str> {
        self.state
            .get_workbook()
            .ok()
            .and_then(|w| w.sheets.get(sheet))
            .map(|s| s.name.as_str())
    }

    fn sheet_index(&self, name: &str) -> Option<usize> {
        self.state.get_workbook().ok().and_then(|w| {
            w.sheets
                .iter()
                .position(|s| s.name.eq_ignore_ascii_case(name))
        })
    }

    fn active_sheet(&self) -> usize {
        self.active_sheet
    }

    fn set_active_sheet(&mut self, sheet: usize) -> Result<(), formula_vba_runtime::VbaError> {
        if sheet >= self.sheet_count() {
            return Err(formula_vba_runtime::VbaError::Runtime(format!(
                "Sheet index out of range: {sheet}"
            )));
        }
        self.active_sheet = sheet;
        Ok(())
    }

    fn active_cell(&self) -> (u32, u32) {
        self.active_cell
    }

    fn set_active_cell(&mut self, row: u32, col: u32) -> Result<(), formula_vba_runtime::VbaError> {
        if row == 0 || col == 0 {
            return Err(formula_vba_runtime::VbaError::Runtime(
                "ActiveCell is 1-based".to_string(),
            ));
        }
        self.active_cell = (row, col);
        Ok(())
    }

    fn get_cell_value(
        &self,
        sheet: usize,
        row: u32,
        col: u32,
    ) -> Result<formula_vba_runtime::VbaValue, formula_vba_runtime::VbaError> {
        let sheet_id = self.sheet_id(sheet)?;
        if row == 0 || col == 0 {
            return Err(formula_vba_runtime::VbaError::Runtime(
                "Row/col are 1-based".to_string(),
            ));
        }
        let cell = self
            .state
            .get_cell(&sheet_id, (row - 1) as usize, (col - 1) as usize)
            .map_err(|e| formula_vba_runtime::VbaError::Runtime(e.to_string()))?;
        Ok(Self::cell_scalar_to_vba(cell.value))
    }

    fn set_cell_value(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
        value: formula_vba_runtime::VbaValue,
    ) -> Result<(), formula_vba_runtime::VbaError> {
        let sheet_id = self.sheet_id(sheet)?;
        if row == 0 || col == 0 {
            return Err(formula_vba_runtime::VbaError::Runtime(
                "Row/col are 1-based".to_string(),
            ));
        }

        let json = Self::vba_value_to_json(&value)?;
        let updates = self
            .state
            .set_cell(
                &sheet_id,
                (row - 1) as usize,
                (col - 1) as usize,
                json,
                None,
            )
            .map_err(|e| formula_vba_runtime::VbaError::Runtime(e.to_string()))?;
        self.updates.extend(updates);
        Ok(())
    }

    fn get_cell_formula(
        &self,
        sheet: usize,
        row: u32,
        col: u32,
    ) -> Result<Option<String>, formula_vba_runtime::VbaError> {
        let sheet_id = self.sheet_id(sheet)?;
        if row == 0 || col == 0 {
            return Err(formula_vba_runtime::VbaError::Runtime(
                "Row/col are 1-based".to_string(),
            ));
        }
        let cell = self
            .state
            .get_cell(&sheet_id, (row - 1) as usize, (col - 1) as usize)
            .map_err(|e| formula_vba_runtime::VbaError::Runtime(e.to_string()))?;
        Ok(cell.formula)
    }

    fn set_cell_formula(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
        formula: String,
    ) -> Result<(), formula_vba_runtime::VbaError> {
        let sheet_id = self.sheet_id(sheet)?;
        if row == 0 || col == 0 {
            return Err(formula_vba_runtime::VbaError::Runtime(
                "Row/col are 1-based".to_string(),
            ));
        }

        let updates = self
            .state
            .set_cell(
                &sheet_id,
                (row - 1) as usize,
                (col - 1) as usize,
                None,
                Some(formula),
            )
            .map_err(|e| formula_vba_runtime::VbaError::Runtime(e.to_string()))?;
        self.updates.extend(updates);
        Ok(())
    }

    fn log(&mut self, message: String) {
        self.output.push(message);
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn run_macro(
    workbook_id: Option<String>,
    macro_id: String,
    permissions: Option<Vec<MacroPermission>>,
    timeout_ms: Option<u64>,
    state: State<'_, SharedAppState>,
) -> Result<MacroRunResult, String> {
    use std::time::Duration;

    let _ = workbook_id;
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();

        let program = build_vba_program(&state)?;

        let mut policy = formula_vba_runtime::VbaSandboxPolicy::default();
        if let Some(timeout_ms) = timeout_ms {
            policy.max_execution_time = Duration::from_millis(timeout_ms);
        }
        if let Some(perms) = permissions {
            for perm in perms {
                match perm {
                    MacroPermission::FilesystemRead => policy.allow_filesystem_read = true,
                    MacroPermission::FilesystemWrite => policy.allow_filesystem_write = true,
                    MacroPermission::Network => policy.allow_network = true,
                }
            }
        }

        let runtime = formula_vba_runtime::VbaRuntime::new(program).with_sandbox_policy(policy);
        let mut sheet = AppStateSpreadsheet::new(&mut state).map_err(|err| err.to_string())?;

        let exec = runtime.execute(&mut sheet, &macro_id, &[]);
        let output = sheet.take_output();
        let updates = sheet
            .take_updates()
            .into_iter()
            .map(cell_update_from_state)
            .collect::<Vec<_>>();

        Ok::<_, String>(match exec {
            Ok(_) => MacroRunResult {
                ok: true,
                output,
                updates,
                error: None,
            },
            Err(err) => MacroRunResult {
                ok: false,
                output,
                updates,
                error: Some(MacroError {
                    message: err.to_string(),
                }),
            },
        })
    })
    .await
    .map_err(|e| e.to_string())?
}
