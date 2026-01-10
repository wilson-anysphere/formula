use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;

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
use tauri::State;

#[cfg(feature = "desktop")]
fn app_error(err: AppStateError) -> String {
    err.to_string()
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
    let info = {
        let mut state = state.inner().lock().unwrap();
        state.load_workbook(workbook)
    };

    Ok(WorkbookInfo {
        path: info.path,
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
