use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;
use formula_engine::pivot::PivotConfig;

use crate::macro_trust::MacroTrustDecision;

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
pub struct SheetUsedRange {
    pub start_row: usize,
    pub end_row: usize,
    pub start_col: usize,
    pub end_col: usize,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct WorkbookThemePalette {
    pub dk1: String,
    pub lt1: String,
    pub dk2: String,
    pub lt2: String,
    pub accent1: String,
    pub accent2: String,
    pub accent3: String,
    pub accent4: String,
    pub accent5: String,
    pub accent6: String,
    pub hlink: String,
    pub followed_hlink: String,
}

#[cfg(any(feature = "desktop", test))]
fn rgb_hex(argb: u32) -> String {
    format!("#{:06X}", argb & 0x00FF_FFFF)
}

#[cfg(any(feature = "desktop", test))]
fn workbook_theme_palette(workbook: &crate::file_io::Workbook) -> Option<WorkbookThemePalette> {
    let palette = workbook.theme_palette.as_ref()?;
    Some(WorkbookThemePalette {
        dk1: rgb_hex(palette.dk1),
        lt1: rgb_hex(palette.lt1),
        dk2: rgb_hex(palette.dk2),
        lt2: rgb_hex(palette.lt2),
        accent1: rgb_hex(palette.accent1),
        accent2: rgb_hex(palette.accent2),
        accent3: rgb_hex(palette.accent3),
        accent4: rgb_hex(palette.accent4),
        accent5: rgb_hex(palette.accent5),
        accent6: rgb_hex(palette.accent6),
        hlink: rgb_hex(palette.hlink),
        followed_hlink: rgb_hex(palette.followed_hlink),
    })
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct RangeCellEdit {
    pub value: Option<JsonValue>,
    pub formula: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PivotCellRange {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PivotDestination {
    pub sheet_id: String,
    pub row: usize,
    pub col: usize,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct CreatePivotTableRequest {
    pub name: String,
    pub source_sheet_id: String,
    pub source_range: PivotCellRange,
    pub destination: PivotDestination,
    pub config: PivotConfig,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct CreatePivotTableResponse {
    pub pivot_id: String,
    pub updates: Vec<CellUpdate>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct RefreshPivotTableRequest {
    pub pivot_id: String,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PivotTableSummary {
    pub id: String,
    pub name: String,
    pub source_sheet_id: String,
    pub source_range: PivotCellRange,
    pub destination: PivotDestination,
}

#[cfg(feature = "desktop")]
use crate::file_io::{read_csv, read_xlsx, write_xlsx};
#[cfg(feature = "desktop")]
use crate::state::{AppState, AppStateError, CellUpdateData, SharedAppState};
#[cfg(feature = "desktop")]
use crate::{
    macro_trust::{compute_macro_fingerprint, SharedMacroTrustStore},
    file_io::Workbook,
};
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
    if buf.extension().and_then(|s| s.to_str()).is_some_and(|ext| {
        ext.eq_ignore_ascii_case("xls")
            || ext.eq_ignore_ascii_case("xlsb")
            || ext.eq_ignore_ascii_case("csv")
    }) {
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
    let ext = PathBuf::from(&path)
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or_default()
        .to_ascii_lowercase();
    let workbook = match ext.as_str() {
        "csv" => read_csv(path.clone()).await.map_err(|e| e.to_string())?,
        _ => read_xlsx(path.clone()).await.map_err(|e| e.to_string())?,
    };
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
pub async fn new_workbook(state: State<'_, SharedAppState>) -> Result<WorkbookInfo, String> {
    let shared = state.inner().clone();
    let info = tauri::async_runtime::spawn_blocking(move || {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

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
pub fn get_workbook_theme_palette(
    state: State<'_, SharedAppState>,
) -> Result<Option<WorkbookThemePalette>, String> {
    let state = state.inner().lock().unwrap();
    let workbook = state.get_workbook().map_err(app_error)?;
    Ok(workbook_theme_palette(workbook))
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

    let written_bytes = write_xlsx(save_path.clone(), workbook)
        .await
        .map_err(|e| e.to_string())?;

    {
        let mut state = state.inner().lock().unwrap();
        state
            .mark_saved(Some(save_path), Some(written_bytes))
            .map_err(app_error)?;
    }

    Ok(())
}

/// Mark the in-memory workbook state as saved (clears the dirty flag) without writing a file.
///
/// This is useful when the frontend returns to the last-saved state via undo/redo and wants the
/// close prompt to match `DocumentController.isDirty`.
#[cfg(feature = "desktop")]
#[tauri::command]
pub fn mark_saved(state: State<'_, SharedAppState>) -> Result<(), String> {
    let mut state = state.inner().lock().unwrap();
    state.mark_saved(None, None).map_err(app_error)?;
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
pub fn get_sheet_used_range(
    sheet_id: String,
    state: State<'_, SharedAppState>,
) -> Result<Option<SheetUsedRange>, String> {
    let state = state.inner().lock().unwrap();
    let workbook = state.get_workbook().map_err(app_error)?;
    let sheet = workbook
        .sheet(&sheet_id)
        .ok_or_else(|| app_error(AppStateError::UnknownSheet(sheet_id)))?;

    if let Some(table) = &sheet.columnar {
        let rows = table.row_count();
        let cols = table.column_count();
        if rows == 0 || cols == 0 {
            return Ok(None);
        }
        return Ok(Some(SheetUsedRange {
            start_row: 0,
            end_row: rows.saturating_sub(1),
            start_col: 0,
            end_col: cols.saturating_sub(1),
        }));
    }

    let mut min_row = usize::MAX;
    let mut min_col = usize::MAX;
    let mut max_row = 0usize;
    let mut max_col = 0usize;
    let mut has_any = false;

    for ((row, col), _cell) in sheet.cells_iter() {
        has_any = true;
        min_row = min_row.min(row);
        min_col = min_col.min(col);
        max_row = max_row.max(row);
        max_col = max_col.max(col);
    }

    if !has_any {
        return Ok(None);
    }

    Ok(Some(SheetUsedRange {
        start_row: min_row,
        end_row: max_row,
        start_col: min_col,
        end_col: max_col,
    }))
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
pub async fn create_pivot_table(
    request: CreatePivotTableRequest,
    state: State<'_, SharedAppState>,
) -> Result<CreatePivotTableResponse, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let (pivot_id, updates) = state
            .create_pivot_table(
                request.name,
                request.source_sheet_id,
                crate::state::CellRect {
                    start_row: request.source_range.start_row,
                    start_col: request.source_range.start_col,
                    end_row: request.source_range.end_row,
                    end_col: request.source_range.end_col,
                },
                crate::state::PivotDestination {
                    sheet_id: request.destination.sheet_id,
                    row: request.destination.row,
                    col: request.destination.col,
                },
                request.config,
            )
            .map_err(app_error)?;

        Ok::<_, String>(CreatePivotTableResponse {
            pivot_id,
            updates: updates.into_iter().map(cell_update_from_state).collect(),
        })
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn refresh_pivot_table(
    request: RefreshPivotTableRequest,
    state: State<'_, SharedAppState>,
) -> Result<Vec<CellUpdate>, String> {
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let updates = state
            .refresh_pivot_table(&request.pivot_id)
            .map_err(app_error)?;
        Ok::<_, String>(updates.into_iter().map(cell_update_from_state).collect())
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn list_pivot_tables(state: State<'_, SharedAppState>) -> Result<Vec<PivotTableSummary>, String> {
    let state = state.inner().lock().unwrap();
    Ok(state
        .list_pivot_tables()
        .into_iter()
        .map(|pivot| PivotTableSummary {
            id: pivot.id,
            name: pivot.name,
            source_sheet_id: pivot.source_sheet_id,
            source_range: PivotCellRange {
                start_row: pivot.source_range.start_row,
                start_col: pivot.source_range.start_col,
                end_row: pivot.source_range.end_row,
                end_col: pivot.source_range.end_col,
            },
            destination: PivotDestination {
                sheet_id: pivot.destination.sheet_id,
                row: pivot.destination.row,
                col: pivot.destination.col,
            },
        })
        .collect())
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

pub use crate::macros::{MacroInfo, MacroPermission, MacroPermissionRequest};

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "snake_case")]
pub enum MacroSignatureStatus {
    Unsigned,
    SignedUnverified,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroSignatureInfo {
    pub status: MacroSignatureStatus,
    pub signer_subject: Option<String>,
    /// Raw signature blob, base64 encoded. May be omitted in the future if it grows large.
    pub signature_base64: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroSecurityStatus {
    pub has_macros: bool,
    pub origin_path: Option<String>,
    pub workbook_fingerprint: Option<String>,
    pub signature: Option<MacroSignatureInfo>,
    pub trust: MacroTrustDecision,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "snake_case")]
pub enum MacroBlockedReason {
    NotTrusted,
    SignatureRequired,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroBlockedError {
    pub reason: MacroBlockedReason,
    pub status: MacroSecurityStatus,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroError {
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub code: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub blocked: Option<MacroBlockedError>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroRunResult {
    pub ok: bool,
    pub output: Vec<String>,
    pub updates: Vec<CellUpdate>,
    pub error: Option<MacroError>,
    pub permission_request: Option<MacroPermissionRequest>,
}

#[cfg(feature = "desktop")]
fn workbook_identity_for_trust(workbook: &Workbook, workbook_id: Option<&str>) -> String {
    workbook
        .origin_path
        .as_deref()
        .or(workbook.path.as_deref())
        .or(workbook_id)
        .unwrap_or("untitled")
        .to_string()
}

#[cfg(feature = "desktop")]
fn compute_workbook_fingerprint(workbook: &mut Workbook, workbook_id: Option<&str>) -> Option<String> {
    if workbook.vba_project_bin.is_none() {
        return None;
    }
    if let Some(fp) = workbook.macro_fingerprint.as_deref() {
        return Some(fp.to_string());
    }
    let id = workbook_identity_for_trust(workbook, workbook_id);
    let vba = workbook
        .vba_project_bin
        .as_deref()
        .expect("checked is_some above");
    let fp = compute_macro_fingerprint(&id, vba);
    workbook.macro_fingerprint = Some(fp.clone());
    Some(fp)
}

#[cfg(feature = "desktop")]
fn build_macro_security_status(
    workbook: &mut Workbook,
    workbook_id: Option<&str>,
    trust_store: &crate::macro_trust::MacroTrustStore,
) -> Result<MacroSecurityStatus, String> {
    use base64::Engine as _;

    let has_macros = workbook.vba_project_bin.is_some();
    let fingerprint = compute_workbook_fingerprint(workbook, workbook_id);

    let signature = if let Some(vba_bin) = workbook.vba_project_bin.as_deref() {
        // Signature parsing is best-effort: failures should not prevent macro listing or
        // execution (trust decisions are still enforced by the fingerprint).
        let parsed = formula_vba::parse_vba_digital_signature(vba_bin).ok().flatten();
        Some(match parsed {
            Some(sig) => MacroSignatureInfo {
                status: MacroSignatureStatus::SignedUnverified,
                signer_subject: sig.signer_subject,
                signature_base64: Some(
                    base64::engine::general_purpose::STANDARD.encode(sig.signature),
                ),
            },
            None => MacroSignatureInfo {
                status: MacroSignatureStatus::Unsigned,
                signer_subject: None,
                signature_base64: None,
            },
        })
    } else {
        None
    };

    let trust = fingerprint
        .as_deref()
        .map(|fp| trust_store.trust_state(fp))
        .unwrap_or(MacroTrustDecision::Blocked);

    Ok(MacroSecurityStatus {
        has_macros,
        origin_path: workbook.origin_path.clone(),
        workbook_fingerprint: fingerprint,
        signature,
        trust,
    })
}

#[cfg(feature = "desktop")]
fn enforce_macro_trust(
    workbook: &mut Workbook,
    workbook_id: Option<&str>,
    trust_store: &crate::macro_trust::MacroTrustStore,
) -> Result<Option<MacroBlockedError>, String> {
    let status = build_macro_security_status(workbook, workbook_id, trust_store)?;
    if !status.has_macros {
        return Ok(None);
    }

    let is_signed = matches!(
        status.signature.as_ref().map(|s| s.status),
        Some(MacroSignatureStatus::SignedUnverified)
    );
    let allowed = match status.trust {
        MacroTrustDecision::TrustedAlways | MacroTrustDecision::TrustedOnce => true,
        MacroTrustDecision::TrustedSignedOnly => is_signed,
        MacroTrustDecision::Blocked => false,
    };

    if allowed {
        return Ok(None);
    }

    let reason = match status.trust {
        MacroTrustDecision::TrustedSignedOnly => MacroBlockedReason::SignatureRequired,
        _ => MacroBlockedReason::NotTrusted,
    };

    Ok(Some(MacroBlockedError { reason, status }))
}

#[cfg(feature = "desktop")]
fn macro_blocked_result(blocked: MacroBlockedError) -> MacroRunResult {
    MacroRunResult {
        ok: false,
        output: Vec::new(),
        updates: Vec::new(),
        error: Some(MacroError {
            message: "Macros are blocked by Trust Center policy.".to_string(),
            code: Some("macro_blocked".to_string()),
            blocked: Some(blocked),
        }),
        permission_request: None,
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct VbaReferenceSummary {
    pub name: Option<String>,
    pub guid: Option<String>,
    pub major: Option<u16>,
    pub minor: Option<u16>,
    pub path: Option<String>,
    pub raw: String,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct VbaModuleSummary {
    pub name: String,
    pub module_type: String,
    pub code: String,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct VbaProjectSummary {
    pub name: Option<String>,
    pub constants: Option<String>,
    pub references: Vec<VbaReferenceSummary>,
    pub modules: Vec<VbaModuleSummary>,
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn get_macro_security_status(
    workbook_id: Option<String>,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroSecurityStatus, String> {
    let workbook_id = workbook_id.as_deref();
    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let mut trust_store = trust_shared.lock().unwrap();
        let workbook = state.get_workbook_mut().map_err(app_error)?;
        build_macro_security_status(workbook, workbook_id, &trust_store)
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn set_macro_trust(
    workbook_id: Option<String>,
    decision: MacroTrustDecision,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroSecurityStatus, String> {
    let workbook_id = workbook_id.as_deref();
    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let mut trust_store = trust_shared.lock().unwrap();

        let workbook = state.get_workbook_mut().map_err(app_error)?;
        let Some(fingerprint) = compute_workbook_fingerprint(workbook, workbook_id) else {
            return Err("workbook has no macros to trust".to_string());
        };

        trust_store
            .set_trust(fingerprint, decision)
            .map_err(|e| e.to_string())?;

        build_macro_security_status(workbook, workbook_id, &trust_store)
    })
    .await
    .map_err(|e| e.to_string())?
}
#[cfg(feature = "desktop")]
#[tauri::command]
pub fn get_vba_project(
    workbook_id: Option<String>,
    state: State<'_, SharedAppState>,
) -> Result<Option<VbaProjectSummary>, String> {
    let _ = workbook_id;
    let mut state = state.inner().lock().unwrap();
    let Some(project) = state.vba_project().map_err(|e| e.to_string())? else {
        return Ok(None);
    };
    Ok(Some(VbaProjectSummary {
        name: project.name,
        constants: project.constants,
        references: project
            .references
            .into_iter()
            .map(|r| VbaReferenceSummary {
                name: r.name,
                guid: r.guid,
                major: r.major,
                minor: r.minor,
                path: r.path,
                raw: r.raw,
            })
            .collect(),
        modules: project
            .modules
            .into_iter()
            .map(|m| VbaModuleSummary {
                name: m.name,
                module_type: format!("{:?}", m.module_type),
                code: m.code,
            })
            .collect(),
    }))
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn list_macros(
    workbook_id: Option<String>,
    state: State<'_, SharedAppState>,
) -> Result<Vec<MacroInfo>, String> {
    let _ = workbook_id;

    let mut state = state.inner().lock().unwrap();
    state.list_macros().map_err(|e| e.to_string())
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn run_macro(
    workbook_id: Option<String>,
    macro_id: String,
    permissions: Option<Vec<MacroPermission>>,
    timeout_ms: Option<u64>,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroRunResult, String> {
    let workbook_id_str = workbook_id.clone();
    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let blocked = {
            let trust_store = trust_shared.lock().unwrap();
            let workbook_id = workbook_id_str.as_deref();
            let workbook = state.get_workbook_mut().map_err(app_error)?;
            enforce_macro_trust(workbook, workbook_id, &trust_store)?
        };
        if let Some(blocked) = blocked {
            return Ok(macro_blocked_result(blocked));
        }

        let options = crate::macros::MacroExecutionOptions {
            permissions: permissions.unwrap_or_default(),
            timeout_ms,
        };
        let outcome = state
            .run_macro(&macro_id, options)
            .map_err(|e| e.to_string())?;
        Ok::<_, String>(macro_result_from_outcome(outcome))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
fn macro_result_from_outcome(outcome: crate::macros::MacroExecutionOutcome) -> MacroRunResult {
    MacroRunResult {
        ok: outcome.ok,
        output: outcome.output,
        updates: outcome
            .updates
            .into_iter()
            .map(cell_update_from_state)
            .collect(),
        error: outcome.error.map(|message| MacroError {
            message,
            code: None,
            blocked: None,
        }),
        permission_request: outcome.permission_request,
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn fire_workbook_open(
    workbook_id: Option<String>,
    permissions: Option<Vec<MacroPermission>>,
    timeout_ms: Option<u64>,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroRunResult, String> {
    let workbook_id_str = workbook_id.clone();
    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let blocked = {
            let trust_store = trust_shared.lock().unwrap();
            let workbook_id = workbook_id_str.as_deref();
            let workbook = state.get_workbook_mut().map_err(app_error)?;
            enforce_macro_trust(workbook, workbook_id, &trust_store)?
        };
        if let Some(blocked) = blocked {
            return Ok(macro_blocked_result(blocked));
        }
        let options = crate::macros::MacroExecutionOptions {
            permissions: permissions.unwrap_or_default(),
            timeout_ms,
        };
        let outcome = state
            .fire_workbook_open(options)
            .map_err(|e| e.to_string())?;
        Ok::<_, String>(macro_result_from_outcome(outcome))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn fire_workbook_before_close(
    workbook_id: Option<String>,
    permissions: Option<Vec<MacroPermission>>,
    timeout_ms: Option<u64>,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroRunResult, String> {
    let workbook_id_str = workbook_id.clone();
    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let blocked = {
            let trust_store = trust_shared.lock().unwrap();
            let workbook_id = workbook_id_str.as_deref();
            let workbook = state.get_workbook_mut().map_err(app_error)?;
            enforce_macro_trust(workbook, workbook_id, &trust_store)?
        };
        if let Some(blocked) = blocked {
            return Ok(macro_blocked_result(blocked));
        }
        let options = crate::macros::MacroExecutionOptions {
            permissions: permissions.unwrap_or_default(),
            timeout_ms,
        };
        let outcome = state
            .fire_workbook_before_close(options)
            .map_err(|e| e.to_string())?;
        Ok::<_, String>(macro_result_from_outcome(outcome))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn fire_worksheet_change(
    workbook_id: Option<String>,
    sheet_id: String,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    permissions: Option<Vec<MacroPermission>>,
    timeout_ms: Option<u64>,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroRunResult, String> {
    let workbook_id_str = workbook_id.clone();
    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let blocked = {
            let trust_store = trust_shared.lock().unwrap();
            let workbook_id = workbook_id_str.as_deref();
            let workbook = state.get_workbook_mut().map_err(app_error)?;
            enforce_macro_trust(workbook, workbook_id, &trust_store)?
        };
        if let Some(blocked) = blocked {
            return Ok(macro_blocked_result(blocked));
        }
        let options = crate::macros::MacroExecutionOptions {
            permissions: permissions.unwrap_or_default(),
            timeout_ms,
        };
        let outcome = state
            .fire_worksheet_change(&sheet_id, start_row, start_col, end_row, end_col, options)
            .map_err(|e| e.to_string())?;
        Ok::<_, String>(macro_result_from_outcome(outcome))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn fire_selection_change(
    workbook_id: Option<String>,
    sheet_id: String,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
    permissions: Option<Vec<MacroPermission>>,
    timeout_ms: Option<u64>,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MacroRunResult, String> {
    let workbook_id_str = workbook_id.clone();
    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        let blocked = {
            let trust_store = trust_shared.lock().unwrap();
            let workbook_id = workbook_id_str.as_deref();
            let workbook = state.get_workbook_mut().map_err(app_error)?;
            enforce_macro_trust(workbook, workbook_id, &trust_store)?
        };
        if let Some(blocked) = blocked {
            return Ok(macro_blocked_result(blocked));
        }
        let options = crate::macros::MacroExecutionOptions {
            permissions: permissions.unwrap_or_default(),
            timeout_ms,
        };
        let outcome = state
            .fire_selection_change(&sheet_id, start_row, start_col, end_row, end_col, options)
            .map_err(|e| e.to_string())?;
        Ok::<_, String>(macro_result_from_outcome(outcome))
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::file_io::read_xlsx_blocking;
    use std::path::Path;

    #[test]
    fn workbook_theme_palette_is_exposed_for_rt_simple_fixture() {
        let fixture_path = Path::new(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../../crates/formula-xlsx/tests/fixtures/rt_simple.xlsx"
        ));

        let workbook = read_xlsx_blocking(fixture_path).expect("read fixture workbook");
        let palette = workbook_theme_palette(&workbook).expect("palette should be present");

        for value in [
            palette.dk1,
            palette.lt1,
            palette.dk2,
            palette.lt2,
            palette.accent1,
            palette.accent2,
            palette.accent3,
            palette.accent4,
            palette.accent5,
            palette.accent6,
            palette.hlink,
            palette.followed_hlink,
        ] {
            assert!(
                value.len() == 7
                    && value.starts_with('#')
                    && value
                        .chars()
                        .skip(1)
                        .all(|c| c.is_ascii_hexdigit() && !c.is_ascii_lowercase()),
                "expected hex color like '#RRGGBB', got {value}"
            );
        }
    }
}
