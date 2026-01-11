use formula_engine::pivot::PivotConfig;
use serde::{Deserialize, Serialize};
use serde_json::Value as JsonValue;

use crate::macro_trust::MacroTrustDecision;
#[cfg(feature = "desktop")]
use crate::storage::power_query_cache_key::{PowerQueryCacheKey, PowerQueryCacheKeyStore};
#[cfg(feature = "desktop")]
use crate::storage::power_query_credentials::{
    PowerQueryCredentialEntry, PowerQueryCredentialListEntry, PowerQueryCredentialStore,
};
#[cfg(feature = "desktop")]
use crate::storage::power_query_refresh_state::PowerQueryRefreshStateStore;

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
pub struct DefinedNameInfo {
    pub name: String,
    pub refers_to: String,
    pub sheet_id: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct TableInfo {
    pub name: String,
    pub sheet_id: String,
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
    pub columns: Vec<String>,
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
use crate::file_io::{read_csv, read_xlsx};
#[cfg(feature = "desktop")]
use crate::persistence::{
    autosave_db_path_for_new_workbook, autosave_db_path_for_workbook, WorkbookPersistenceLocation,
};
#[cfg(feature = "desktop")]
use crate::state::SharedAppState;
#[cfg(any(feature = "desktop", test))]
use crate::state::{AppState, AppStateError, CellUpdateData};
#[cfg(feature = "desktop")]
use crate::{
    file_io::Workbook,
    macro_trust::{compute_macro_fingerprint, SharedMacroTrustStore},
};
#[cfg(feature = "desktop")]
use std::path::PathBuf;
#[cfg(feature = "desktop")]
use std::sync::Arc;
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
        .is_some_and(|ext| ext.eq_ignore_ascii_case("xls") || ext.eq_ignore_ascii_case("csv"))
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

#[cfg(any(feature = "desktop", test))]
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
    let location = autosave_db_path_for_workbook(&path)
        .map(WorkbookPersistenceLocation::OnDisk)
        .unwrap_or(WorkbookPersistenceLocation::InMemory);

    let shared = state.inner().clone();
    let info = tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        state
            .load_workbook_persistent(workbook, location)
            .map_err(app_error)
    })
    .await
    .map_err(|e| e.to_string())?;
    let info = info?;

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
    let location = autosave_db_path_for_new_workbook()
        .map(WorkbookPersistenceLocation::OnDisk)
        .unwrap_or(WorkbookPersistenceLocation::InMemory);
    let info = tauri::async_runtime::spawn_blocking(move || {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        let mut state = shared.lock().unwrap();
        state
            .load_workbook_persistent(workbook, location)
            .map_err(app_error)
    })
    .await
    .map_err(|e| e.to_string())?;
    let info = info?;

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

/// Read a local text file on behalf of the frontend.
///
/// This exists so the desktop webview can power-query local sources (CSV/JSON) without
/// depending on the optional Tauri FS plugin.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn read_text_file(path: String) -> Result<String, String> {
    tauri::async_runtime::spawn_blocking(move || std::fs::read_to_string(path))
        .await
        .map_err(|e| e.to_string())?
        .map_err(|e| e.to_string())
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct FileStat {
    pub mtime_ms: u64,
    pub size_bytes: u64,
}

/// Stat a local file and return its modification time and size.
///
/// This is used by Power Query's cache validation logic to decide whether cached results can be
/// reused when reading local sources (CSV/JSON/Parquet).
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn stat_file(path: String) -> Result<FileStat, String> {
    use std::time::UNIX_EPOCH;

    tauri::async_runtime::spawn_blocking(move || {
        let metadata = std::fs::metadata(path).map_err(|e| e.to_string())?;
        let modified = metadata.modified().map_err(|e| e.to_string())?;
        let duration = modified
            .duration_since(UNIX_EPOCH)
            .map_err(|e| e.to_string())?;
        Ok::<_, String>(FileStat {
            mtime_ms: duration.as_millis() as u64,
            size_bytes: metadata.len(),
        })
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Read a local file and return its contents as base64.
///
/// The frontend decodes this into a `Uint8Array` for Parquet sources.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn read_binary_file(path: String) -> Result<String, String> {
    use base64::{engine::general_purpose::STANDARD, Engine as _};

    let bytes = tauri::async_runtime::spawn_blocking(move || std::fs::read(path))
        .await
        .map_err(|e| e.to_string())?
        .map_err(|e| e.to_string())?;

    Ok(STANDARD.encode(bytes))
}

/// Read a byte range from a local file and return the contents as base64.
///
/// This enables streaming reads for large local sources (e.g. CSV/Parquet) without first
/// materializing the full file into memory.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn read_binary_file_range(
    path: String,
    offset: u64,
    length: u64,
) -> Result<String, String> {
    use base64::{engine::general_purpose::STANDARD, Engine as _};
    use std::io::{Read, Seek, SeekFrom};

    tauri::async_runtime::spawn_blocking(move || {
        if length == 0 {
            return Ok::<_, String>(String::new());
        }

        let mut file = std::fs::File::open(path).map_err(|e| e.to_string())?;
        file.seek(SeekFrom::Start(offset))
            .map_err(|e| e.to_string())?;

        let len = usize::try_from(length)
            .map_err(|_| "Requested length exceeds platform limits".to_string())?;
        let mut buf = vec![0u8; len];
        let read = file.read(&mut buf).map_err(|e| e.to_string())?;
        buf.truncate(read);

        Ok::<_, String>(STANDARD.encode(buf))
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: retrieve (or create) the AES-256-GCM key used to encrypt cached
/// query results at rest.
///
/// The key material is stored in the OS keychain so cached results remain
/// decryptable across app restarts.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_cache_key_get_or_create() -> Result<PowerQueryCacheKey, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryCacheKeyStore::open_default();
        store.get_or_create().map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: retrieve a persisted credential entry by scope key.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_credential_get(
    scope_key: String,
) -> Result<Option<PowerQueryCredentialEntry>, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryCredentialStore::open_default().map_err(|e| e.to_string())?;
        store.get(&scope_key).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: persist a credential entry for a scope key.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_credential_set(
    scope_key: String,
    secret: JsonValue,
) -> Result<PowerQueryCredentialEntry, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryCredentialStore::open_default().map_err(|e| e.to_string())?;
        store.set(&scope_key, secret).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: delete any persisted credential entry for a scope key.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_credential_delete(scope_key: String) -> Result<(), String> {
    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryCredentialStore::open_default().map_err(|e| e.to_string())?;
        store.delete(&scope_key).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: list persisted credential scope keys and IDs (for debugging).
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_credential_list() -> Result<Vec<PowerQueryCredentialListEntry>, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryCredentialStore::open_default().map_err(|e| e.to_string())?;
        store.list().map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: load the persisted refresh scheduling state for a workbook.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_refresh_state_get(
    workbook_id: String,
) -> Result<Option<JsonValue>, String> {
    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryRefreshStateStore::open_default().map_err(|e| e.to_string())?;
        store.load(&workbook_id).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Power Query: persist refresh scheduling state for a workbook.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn power_query_refresh_state_set(
    workbook_id: String,
    state: JsonValue,
) -> Result<(), String> {
    tauri::async_runtime::spawn_blocking(move || {
        let store = PowerQueryRefreshStateStore::open_default().map_err(|e| e.to_string())?;
        store.save(&workbook_id, state).map_err(|e| e.to_string())
    })
    .await
    .map_err(|e| e.to_string())?
}

/// Execute a SQL query against a local database connection.
///
/// Used by the desktop Power Query engine (`source.type === "database"`).
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn sql_query(
    connection: JsonValue,
    sql: String,
    params: Option<Vec<JsonValue>>,
    credentials: Option<JsonValue>,
) -> Result<crate::sql::SqlQueryResult, String> {
    crate::sql::sql_query(connection, sql, params.unwrap_or_default(), credentials)
        .await
        .map_err(|e| e.to_string())
}

/// Describe a SQL query (columns/types) without returning data rows.
#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn sql_get_schema(
    connection: JsonValue,
    sql: String,
    credentials: Option<JsonValue>,
) -> Result<crate::sql::SqlSchemaResult, String> {
    crate::sql::sql_get_schema(connection, sql, credentials)
        .await
        .map_err(|e| e.to_string())
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
pub fn list_defined_names(
    state: State<'_, SharedAppState>,
) -> Result<Vec<DefinedNameInfo>, String> {
    let state = state.inner().lock().unwrap();
    let workbook = state.get_workbook().map_err(app_error)?;

    let names = workbook
        .defined_names
        .iter()
        .filter(|n| !n.hidden)
        .filter(|n| !n.name.trim().is_empty())
        .filter(|n| !n.name.to_ascii_lowercase().starts_with("_xlnm."))
        .map(|n| DefinedNameInfo {
            name: n.name.clone(),
            refers_to: n.refers_to.clone(),
            sheet_id: n.sheet_id.clone(),
        })
        .collect();

    Ok(names)
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn list_tables(state: State<'_, SharedAppState>) -> Result<Vec<TableInfo>, String> {
    let state = state.inner().lock().unwrap();
    let workbook = state.get_workbook().map_err(app_error)?;

    let tables = workbook
        .tables
        .iter()
        .filter(|t| !t.name.trim().is_empty())
        .filter(|t| !t.columns.is_empty())
        .map(|t| TableInfo {
            name: t.name.clone(),
            sheet_id: t.sheet_id.clone(),
            start_row: t.start_row,
            start_col: t.start_col,
            end_row: t.end_row,
            end_col: t.end_col,
            columns: t.columns.clone(),
        })
        .collect();

    Ok(tables)
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn save_workbook(
    path: Option<String>,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let (save_path, workbook, storage, memory, workbook_id, autosave) = {
        let state = state.inner().lock().unwrap();
        let workbook = state.get_workbook().map_err(app_error)?.clone();
        let storage = state
            .persistent_storage()
            .ok_or_else(|| "no persistent storage available".to_string())?;
        let memory = state
            .persistent_memory_manager()
            .ok_or_else(|| "no memory manager available".to_string())?;
        let workbook_id = state
            .persistent_workbook_id()
            .ok_or_else(|| "no persistent workbook id available".to_string())?;
        let autosave = state.autosave_manager();
        let save_path = path
            .clone()
            .or_else(|| workbook.path.clone())
            .ok_or_else(|| "no save path provided".to_string())?;
        (save_path, workbook, storage, memory, workbook_id, autosave)
    };

    let save_path = coerce_save_path_to_xlsx(&save_path);

    let wants_origin_bytes = PathBuf::from(&save_path)
        .extension()
        .and_then(|s| s.to_str())
        .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx") || ext.eq_ignore_ascii_case("xlsm"));

    if let Some(autosave) = autosave.as_ref() {
        autosave.flush().await.map_err(|e| e.to_string())?;
    }

    // Always flush the paging cache before exporting to ensure changes are
    // applied even if the autosave task has exited unexpectedly.
    memory.flush_dirty_pages().map_err(|e| e.to_string())?;

    let save_path_clone = save_path.clone();
    let written_bytes = tauri::async_runtime::spawn_blocking(move || {
        let path = std::path::Path::new(&save_path_clone);
        let ext = path.extension().and_then(|s| s.to_str()).unwrap_or_default();

        // XLSB saves must go through the `formula-xlsb` round-trip writer. The storage export
        // path only knows how to generate XLSX.
        if ext.eq_ignore_ascii_case("xlsb") {
            return crate::file_io::write_xlsx_blocking(path, &workbook).map_err(|e| e.to_string());
        }

        // Prefer the existing patch-based save path when we have the original XLSX bytes.
        // This preserves unknown parts (theme, comments, conditional formatting, etc.) by
        // rewriting only the modified worksheet XML.
        //
        // Fall back to the storage->model export path for non-XLSX origins (csv/xls) and
        // for new workbooks without an `origin_xlsx_bytes` baseline.
        if workbook.origin_xlsx_bytes.is_some() {
            crate::file_io::write_xlsx_blocking(path, &workbook).map_err(|e| e.to_string())
        } else {
            crate::persistence::write_xlsx_from_storage(&storage, workbook_id, &workbook, path)
                .map_err(|e| e.to_string())
        }
    })
    .await
    .map_err(|e| e.to_string())??;

    {
        let mut state = state.inner().lock().unwrap();
        state
            .mark_saved(Some(save_path), wants_origin_bytes.then_some(written_bytes))
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

    for ((row, col), cell) in sheet.cells_iter() {
        // Ignore format-only cells (the UI considers used range based on value/formula).
        if cell.formula.is_none() && cell.input_value.is_none() {
            continue;
        }
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
pub fn list_pivot_tables(
    state: State<'_, SharedAppState>,
) -> Result<Vec<PivotTableSummary>, String> {
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

#[derive(Clone, Copy, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum MacroSignatureStatus {
    Unsigned,
    SignedVerified,
    SignedInvalid,
    SignedParseError,
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

#[derive(Clone, Copy, Debug, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum MacroBlockedReason {
    NotTrusted,
    SignatureRequired,
}

/// Evaluate whether Trust Center policy allows macro execution.
///
/// Note: This is intentionally a pure function so it can be unit-tested without
/// requiring the full Tauri "desktop" feature.
pub fn evaluate_macro_trust(
    trust: MacroTrustDecision,
    signature_status: MacroSignatureStatus,
) -> Result<(), MacroBlockedReason> {
    match trust {
        MacroTrustDecision::TrustedAlways | MacroTrustDecision::TrustedOnce => Ok(()),
        MacroTrustDecision::Blocked => Err(MacroBlockedReason::NotTrusted),
        MacroTrustDecision::TrustedSignedOnly => match signature_status {
            MacroSignatureStatus::SignedVerified => Ok(()),
            _ => Err(MacroBlockedReason::SignatureRequired),
        },
    }
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

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "lowercase")]
pub enum PythonFilesystemPermission {
    None,
    Read,
    Readwrite,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "lowercase")]
pub enum PythonNetworkPermission {
    None,
    Allowlist,
    Full,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct PythonPermissions {
    pub filesystem: PythonFilesystemPermission,
    pub network: PythonNetworkPermission,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub network_allowlist: Option<Vec<String>>,
}

impl Default for PythonPermissions {
    fn default() -> Self {
        Self {
            filesystem: PythonFilesystemPermission::None,
            network: PythonNetworkPermission::None,
            network_allowlist: None,
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PythonSelection {
    pub sheet_id: String,
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PythonRunContext {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub active_sheet_id: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub selection: Option<PythonSelection>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PythonError {
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub stack: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct PythonRunResult {
    pub ok: bool,
    pub stdout: String,
    pub stderr: String,
    pub updates: Vec<CellUpdate>,
    pub error: Option<PythonError>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct TypeScriptRunResult {
    pub ok: bool,
    pub updates: Vec<CellUpdate>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<String>,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroSelectionRect {
    pub start_row: usize,
    pub start_col: usize,
    pub end_row: usize,
    pub end_col: usize,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "lowercase")]
pub enum MigrationTarget {
    Python,
    TypeScript,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct MigrationValidationMismatch {
    pub sheet_id: String,
    pub row: usize,
    pub col: usize,
    pub vba: CellValue,
    pub script: CellValue,
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct MigrationValidationReport {
    pub ok: bool,
    pub macro_id: String,
    pub target: MigrationTarget,
    pub mismatches: Vec<MigrationValidationMismatch>,
    pub vba: MacroRunResult,
    pub python: Option<PythonRunResult>,
    pub typescript: Option<TypeScriptRunResult>,
    pub error: Option<String>,
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
fn compute_workbook_fingerprint(
    workbook: &mut Workbook,
    workbook_id: Option<&str>,
) -> Option<String> {
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
        let parsed = formula_vba::verify_vba_digital_signature(vba_bin)
            .ok()
            .flatten();
        Some(match parsed {
            Some(sig) => MacroSignatureInfo {
                status: match sig.verification {
                    formula_vba::VbaSignatureVerification::SignedVerified => {
                        MacroSignatureStatus::SignedVerified
                    }
                    formula_vba::VbaSignatureVerification::SignedInvalid => {
                        MacroSignatureStatus::SignedInvalid
                    }
                    formula_vba::VbaSignatureVerification::SignedParseError => {
                        MacroSignatureStatus::SignedParseError
                    }
                    formula_vba::VbaSignatureVerification::SignedButUnverified => {
                        MacroSignatureStatus::SignedUnverified
                    }
                },
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

    let signature_status = status
        .signature
        .as_ref()
        .map(|s| s.status)
        .unwrap_or(MacroSignatureStatus::Unsigned);

    match evaluate_macro_trust(status.trust, signature_status) {
        Ok(()) => Ok(None),
        Err(reason) => Ok(Some(MacroBlockedError { reason, status })),
    }
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
pub fn set_macro_ui_context(
    workbook_id: Option<String>,
    sheet_id: String,
    active_row: usize,
    active_col: usize,
    selection: Option<MacroSelectionRect>,
    state: State<'_, SharedAppState>,
) -> Result<(), String> {
    let _ = workbook_id;
    let mut state = state.inner().lock().unwrap();
    let selection = selection.map(|rect| crate::state::CellRect {
        start_row: rect.start_row,
        start_col: rect.start_col,
        end_row: rect.end_row,
        end_col: rect.end_col,
    });
    state
        .set_macro_ui_context(&sheet_id, active_row, active_col, selection)
        .map_err(|e| e.to_string())?;
    Ok(())
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
#[tauri::command]
pub async fn run_python_script(
    workbook_id: Option<String>,
    code: String,
    permissions: Option<PythonPermissions>,
    timeout_ms: Option<u64>,
    max_memory_bytes: Option<u64>,
    context: Option<PythonRunContext>,
    state: State<'_, SharedAppState>,
) -> Result<PythonRunResult, String> {
    let _ = workbook_id;
    let shared = state.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut state = shared.lock().unwrap();
        crate::python::run_python_script(
            &mut state,
            &code,
            permissions,
            timeout_ms,
            max_memory_bytes,
            context,
        )
    })
    .await
    .map_err(|e| e.to_string())?
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_string_literal_prefix(input: &str) -> Result<(String, &str), String> {
    let trimmed = input.trim_start();
    let mut chars = trimmed.char_indices();
    let Some((_, quote)) = chars.next() else {
        return Err("expected string literal".to_string());
    };
    if quote != '"' && quote != '\'' {
        return Err("expected string literal".to_string());
    }

    let mut out = String::new();
    let mut escape = false;
    for (idx, ch) in chars {
        if escape {
            let translated = match ch {
                'n' => '\n',
                'r' => '\r',
                't' => '\t',
                '\\' => '\\',
                '"' => '"',
                '\'' => '\'',
                other => other,
            };
            out.push(translated);
            escape = false;
            continue;
        }

        if ch == '\\' {
            escape = true;
            continue;
        }

        if ch == quote {
            let remainder = &trimmed[idx + ch.len_utf8()..];
            return Ok((out, remainder));
        }

        out.push(ch);
    }

    Err("unterminated string literal".to_string())
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_string_literal(expr: &str) -> Result<String, String> {
    let trimmed = expr.trim().trim_end_matches(';').trim();
    let (value, remainder) = parse_typescript_string_literal_prefix(trimmed)?;
    if !remainder.trim().is_empty() {
        return Err(format!(
            "unexpected trailing tokens after string literal: {remainder}"
        ));
    }
    Ok(value)
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_value_expr(expr: &str) -> Result<Option<JsonValue>, String> {
    let trimmed = expr.trim().trim_end_matches(';').trim();
    if trimmed.eq_ignore_ascii_case("null") || trimmed.eq_ignore_ascii_case("undefined") {
        return Ok(None);
    }
    if trimmed.eq_ignore_ascii_case("true") {
        return Ok(Some(JsonValue::Bool(true)));
    }
    if trimmed.eq_ignore_ascii_case("false") {
        return Ok(Some(JsonValue::Bool(false)));
    }

    if trimmed.starts_with('"') || trimmed.starts_with('\'') {
        let value = parse_typescript_string_literal(trimmed)?;
        return Ok(Some(JsonValue::String(value)));
    }

    if let Ok(int_value) = trimmed.parse::<i64>() {
        return Ok(Some(JsonValue::from(int_value)));
    }

    if let Ok(float_value) = trimmed.parse::<f64>() {
        let num = serde_json::Number::from_f64(float_value)
            .ok_or_else(|| format!("invalid numeric literal: {trimmed}"))?;
        return Ok(Some(JsonValue::Number(num)));
    }

    Err(format!("unsupported TypeScript literal: {trimmed}"))
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_call_args(input: &str) -> Result<(String, &str), String> {
    let trimmed = input.trim_start();
    let mut in_string: Option<char> = None;
    let mut escape = false;
    let mut paren_depth: i32 = 0;
    let mut bracket_depth: i32 = 0;
    let mut brace_depth: i32 = 0;

    for (idx, ch) in trimmed.char_indices() {
        if let Some(quote) = in_string {
            if escape {
                escape = false;
                continue;
            }
            if ch == '\\' {
                escape = true;
                continue;
            }
            if ch == quote {
                in_string = None;
            }
            continue;
        }

        match ch {
            '"' | '\'' => in_string = Some(ch),
            '(' => paren_depth += 1,
            ')' => {
                if paren_depth == 0 && bracket_depth == 0 && brace_depth == 0 {
                    let args = trimmed[..idx].to_string();
                    let rest = &trimmed[idx + ch.len_utf8()..];
                    return Ok((args, rest));
                }
                paren_depth -= 1;
            }
            '[' => bracket_depth += 1,
            ']' => bracket_depth = (bracket_depth - 1).max(0),
            '{' => brace_depth += 1,
            '}' => brace_depth = (brace_depth - 1).max(0),
            _ => {}
        }
    }

    Err("unterminated function call".to_string())
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_cell_selector(expr: &str) -> Result<(usize, usize), String> {
    let (start_row, start_col, _end_row, _end_col) = parse_typescript_range_selector(expr)?;
    Ok((start_row, start_col))
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_range_selector(expr: &str) -> Result<(usize, usize, usize, usize), String> {
    if let Some(idx) = expr.rfind(".getRange(") {
        let after = &expr[idx + ".getRange(".len()..];
        let (addr, remainder) = parse_typescript_string_literal_prefix(after)?;
        let remainder = remainder.trim_start();
        if !remainder.starts_with(')') {
            return Err(format!("unsupported getRange expression: {expr}"));
        }
        return parse_typescript_a1_range(&addr);
    }

    if let Some(idx) = expr.rfind(".range(") {
        let after = &expr[idx + ".range(".len()..];
        let (addr, remainder) = parse_typescript_string_literal_prefix(after)?;
        let remainder = remainder.trim_start();
        if !remainder.starts_with(')') {
            return Err(format!("unsupported range expression: {expr}"));
        }
        return parse_typescript_a1_range(&addr);
    }

    if let Some(idx) = expr.rfind(".getCell(") {
        let after = &expr[idx + ".getCell(".len()..];
        let (args, _remainder) = parse_typescript_call_args(after)?;
        let mut parts = args.split(',');
        let row_str = parts.next().unwrap_or("").trim();
        let col_str = parts.next().unwrap_or("").trim();
        if parts.next().is_some() {
            return Err(format!("unsupported getCell expression: {expr}"));
        }
        let row = row_str
            .parse::<usize>()
            .map_err(|_| format!("invalid row in getCell(): {row_str:?}"))?;
        let col = col_str
            .parse::<usize>()
            .map_err(|_| format!("invalid col in getCell(): {col_str:?}"))?;
        return Ok((row, col, row, col));
    }

    if let Some(idx) = expr.rfind(".cell(") {
        let after = &expr[idx + ".cell(".len()..];
        let (args, _remainder) = parse_typescript_call_args(after)?;
        let mut parts = args.split(',');
        let row_str = parts.next().unwrap_or("").trim();
        let col_str = parts.next().unwrap_or("").trim();
        if parts.next().is_some() {
            return Err(format!("unsupported cell expression: {expr}"));
        }
        let row_1 = match row_str.parse::<usize>() {
            Ok(v) if v > 0 => v - 1,
            _ => return Err(format!("invalid row in cell(): {row_str:?}")),
        };
        let col_1 = match col_str.parse::<usize>() {
            Ok(v) if v > 0 => v - 1,
            _ => return Err(format!("invalid col in cell(): {col_str:?}")),
        };
        return Ok((row_1, col_1, row_1, col_1));
    }

    Err(format!("unsupported cell selector: {expr}"))
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_a1_range(addr: &str) -> Result<(usize, usize, usize, usize), String> {
    let addr = addr.trim();
    let mut parts = addr.split(':');
    let start_raw = parts.next().unwrap_or("").trim();
    let end_raw = parts.next().unwrap_or("").trim();
    if parts.next().is_some() {
        return Err(format!("invalid A1 range: {addr:?}"));
    }

    let start = formula_engine::eval::parse_a1(start_raw)
        .map_err(|e| format!("invalid A1 address {start_raw:?}: {e}"))?;
    let end = if end_raw.is_empty() {
        start
    } else {
        formula_engine::eval::parse_a1(end_raw)
            .map_err(|e| format!("invalid A1 address {end_raw:?}: {e}"))?
    };

    let start_row = std::cmp::min(start.row, end.row) as usize;
    let end_row = std::cmp::max(start.row, end.row) as usize;
    let start_col = std::cmp::min(start.col, end.col) as usize;
    let end_col = std::cmp::max(start.col, end.col) as usize;
    Ok((start_row, start_col, end_row, end_col))
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_value_prefix(input: &str) -> Result<(Option<JsonValue>, &str), String> {
    let trimmed = input.trim_start();
    if trimmed.is_empty() {
        return Err("expected literal".to_string());
    }

    let lower = trimmed.to_ascii_lowercase();
    for (token, value) in [
        ("null", None),
        ("undefined", None),
        ("true", Some(JsonValue::Bool(true))),
        ("false", Some(JsonValue::Bool(false))),
    ] {
        if lower.starts_with(token) {
            let remainder = &trimmed[token.len()..];
            let next = remainder.chars().next();
            if matches!(next, Some(ch) if ch.is_ascii_alphanumeric() || ch == '_') {
                // e.g. "nullish" -> not a token.
            } else {
                return Ok((value, remainder));
            }
        }
    }

    if trimmed.starts_with('"') || trimmed.starts_with('\'') {
        let (value, remainder) = parse_typescript_string_literal_prefix(trimmed)?;
        return Ok((Some(JsonValue::String(value)), remainder));
    }

    // Parse a simple number literal (no exponent).
    let bytes = trimmed.as_bytes();
    let mut idx = 0usize;
    if matches!(bytes.get(idx), Some(b'+' | b'-')) {
        idx += 1;
    }
    let start_digits = idx;
    while matches!(bytes.get(idx), Some(b) if b.is_ascii_digit()) {
        idx += 1;
    }
    if idx == start_digits {
        return Err(format!("unsupported TypeScript literal: {trimmed}"));
    }
    if matches!(bytes.get(idx), Some(b'.')) {
        idx += 1;
        let start_frac = idx;
        while matches!(bytes.get(idx), Some(b) if b.is_ascii_digit()) {
            idx += 1;
        }
        if idx == start_frac {
            return Err(format!("invalid numeric literal: {trimmed}"));
        }
    }
    let literal = &trimmed[..idx];
    let remainder = &trimmed[idx..];
    if literal.contains('.') {
        let float_value = literal
            .parse::<f64>()
            .map_err(|_| format!("invalid numeric literal: {literal}"))?;
        let num = serde_json::Number::from_f64(float_value)
            .ok_or_else(|| format!("invalid numeric literal: {literal}"))?;
        return Ok((Some(JsonValue::Number(num)), remainder));
    }
    let int_value = literal
        .parse::<i64>()
        .map_err(|_| format!("invalid numeric literal: {literal}"))?;
    Ok((Some(JsonValue::from(int_value)), remainder))
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_value_matrix(expr: &str) -> Result<Vec<Vec<Option<JsonValue>>>, String> {
    let mut rest = expr.trim_start();
    rest = rest
        .strip_prefix('[')
        .ok_or_else(|| "expected matrix literal like [[1,2],[3,4]]".to_string())?;

    let mut rows: Vec<Vec<Option<JsonValue>>> = Vec::new();
    loop {
        rest = rest.trim_start();
        if let Some(next) = rest.strip_prefix(']') {
            rest = next;
            break;
        }
        rest = rest
            .strip_prefix('[')
            .ok_or_else(|| "expected row literal like [1,2]".to_string())?;

        let mut row: Vec<Option<JsonValue>> = Vec::new();
        loop {
            rest = rest.trim_start();
            if let Some(next) = rest.strip_prefix(']') {
                rest = next;
                break;
            }
            let (value, remainder) = parse_typescript_value_prefix(rest)?;
            row.push(value);
            rest = remainder.trim_start();
            if let Some(next) = rest.strip_prefix(',') {
                rest = next;
                continue;
            }
            if let Some(next) = rest.strip_prefix(']') {
                rest = next;
                break;
            }
            return Err(format!("expected ',' or ']' in row literal, got {rest:?}"));
        }
        rows.push(row);

        rest = rest.trim_start();
        if let Some(next) = rest.strip_prefix(',') {
            rest = next;
            continue;
        }
        if let Some(next) = rest.strip_prefix(']') {
            rest = next;
            break;
        }
        return Err(format!(
            "expected ',' or ']' after row literal, got {rest:?}"
        ));
    }

    if !rest.trim().trim_end_matches(';').trim().is_empty() {
        return Err(format!(
            "unexpected trailing tokens after matrix literal: {rest}"
        ));
    }

    Ok(rows)
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_formulas_matrix(expr: &str) -> Result<Vec<Vec<Option<String>>>, String> {
    let matrix = parse_typescript_value_matrix(expr)?;
    let mut out: Vec<Vec<Option<String>>> = Vec::new();
    for row in matrix {
        let mut row_out: Vec<Option<String>> = Vec::new();
        for value in row {
            match value {
                None => row_out.push(None),
                Some(JsonValue::String(s)) => row_out.push(Some(s)),
                Some(other) => {
                    return Err(format!(
                        "expected formula string literal or null, got {other}"
                    ))
                }
            }
        }
        out.push(row_out);
    }
    Ok(out)
}

#[cfg(any(feature = "desktop", test))]
#[derive(Clone, Debug)]
enum TypeScriptBinding {
    Scalar(Option<JsonValue>),
    Matrix(Vec<Vec<Option<JsonValue>>>),
}

#[cfg(any(feature = "desktop", test))]
fn is_typescript_identifier(input: &str) -> bool {
    let mut chars = input.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !(first.is_ascii_alphabetic() || first == '_') {
        return false;
    }
    chars.all(|ch| ch.is_ascii_alphanumeric() || ch == '_')
}

#[cfg(any(feature = "desktop", test))]
fn split_typescript_top_level_commas(input: &str) -> Vec<&str> {
    let trimmed = input.trim();
    let mut parts: Vec<&str> = Vec::new();
    let mut start = 0usize;

    let mut in_string: Option<char> = None;
    let mut escape = false;
    let mut paren_depth: i32 = 0;
    let mut bracket_depth: i32 = 0;
    let mut brace_depth: i32 = 0;

    for (idx, ch) in trimmed.char_indices() {
        if let Some(quote) = in_string {
            if escape {
                escape = false;
                continue;
            }
            if ch == '\\' {
                escape = true;
                continue;
            }
            if ch == quote {
                in_string = None;
            }
            continue;
        }

        match ch {
            '"' | '\'' => in_string = Some(ch),
            '(' => paren_depth += 1,
            ')' => paren_depth = (paren_depth - 1).max(0),
            '[' => bracket_depth += 1,
            ']' => bracket_depth = (bracket_depth - 1).max(0),
            '{' => brace_depth += 1,
            '}' => brace_depth = (brace_depth - 1).max(0),
            ',' if paren_depth == 0 && bracket_depth == 0 && brace_depth == 0 => {
                parts.push(trimmed[start..idx].trim());
                start = idx + ch.len_utf8();
            }
            _ => {}
        }
    }

    parts.push(trimmed[start..].trim());
    parts
}

#[cfg(any(feature = "desktop", test))]
fn resolve_typescript_scalar_expr(
    expr: &str,
    bindings: &std::collections::HashMap<String, TypeScriptBinding>,
) -> Result<Option<JsonValue>, String> {
    match parse_typescript_value_expr(expr) {
        Ok(value) => Ok(value),
        Err(parse_err) => {
            let ident = expr.trim().trim_end_matches(';').trim();
            if is_typescript_identifier(ident) {
                if let Some(TypeScriptBinding::Scalar(value)) = bindings.get(ident) {
                    return Ok(value.clone());
                }
            }
            Err(parse_err)
        }
    }
}

#[cfg(any(feature = "desktop", test))]
fn parse_typescript_array_from_fill_matrix(
    expr: &str,
    bindings: &std::collections::HashMap<String, TypeScriptBinding>,
) -> Result<Vec<Vec<Option<JsonValue>>>, String> {
    let trimmed = expr.trim().trim_end_matches(';').trim();
    let after = trimmed
        .strip_prefix("Array.from")
        .ok_or_else(|| format!("unsupported matrix expression: {trimmed}"))?;
    let after = after.trim_start();
    let after = after
        .strip_prefix('(')
        .ok_or_else(|| format!("unsupported matrix expression: {trimmed}"))?;

    let (args, remainder) = parse_typescript_call_args(after)?;
    if !remainder.trim().is_empty() {
        return Err(format!(
            "unexpected trailing tokens after Array.from(...): {remainder}"
        ));
    }

    let parts = split_typescript_top_level_commas(&args);
    if parts.len() != 2 {
        return Err(format!("unsupported Array.from(...) arguments: {args}"));
    }

    let length_arg = parts[0];
    let mut rest = length_arg.trim();
    rest = rest
        .strip_prefix('{')
        .ok_or_else(|| format!("unsupported Array.from length arg: {length_arg}"))?;
    rest = rest
        .strip_suffix('}')
        .ok_or_else(|| format!("unsupported Array.from length arg: {length_arg}"))?;
    rest = rest.trim_start();
    rest = rest
        .strip_prefix("length")
        .ok_or_else(|| format!("unsupported Array.from length arg: {length_arg}"))?;
    rest = rest.trim_start();
    rest = rest
        .strip_prefix(':')
        .ok_or_else(|| format!("unsupported Array.from length arg: {length_arg}"))?;
    rest = rest.trim_start();
    let digits = rest
        .chars()
        .take_while(|ch| ch.is_ascii_digit())
        .collect::<String>();
    if digits.is_empty() {
        return Err(format!("unsupported Array.from length arg: {length_arg}"));
    }
    let rows = digits
        .parse::<usize>()
        .map_err(|_| format!("invalid Array.from length: {digits}"))?;
    if rows == 0 {
        return Err("Array.from length must be > 0".to_string());
    }

    let fill_arg = parts[1];
    let fill_str = fill_arg.trim();
    let array_idx = fill_str
        .find("Array(")
        .ok_or_else(|| format!("unsupported Array.from fill arg: {fill_arg}"))?;
    let after_array = &fill_str[array_idx + "Array(".len()..];
    let after_array = after_array.trim_start();
    let digits = after_array
        .chars()
        .take_while(|ch| ch.is_ascii_digit())
        .collect::<String>();
    if digits.is_empty() {
        return Err(format!("unsupported Array.from fill arg: {fill_arg}"));
    }
    let cols = digits
        .parse::<usize>()
        .map_err(|_| format!("invalid Array(...) length: {digits}"))?;
    if cols == 0 {
        return Err("Array(...) length must be > 0".to_string());
    }

    let mut rest = after_array[digits.len()..].trim_start();
    rest = rest
        .strip_prefix(')')
        .ok_or_else(|| format!("unsupported Array.from fill arg: {fill_arg}"))?;
    rest = rest.trim_start();
    rest = rest
        .strip_prefix(".fill(")
        .ok_or_else(|| format!("unsupported Array.from fill arg: {fill_arg}"))?;

    let (fill_expr, remainder) = parse_typescript_call_args(rest)?;
    if !remainder.trim().is_empty() {
        return Err(format!(
            "unexpected trailing tokens after Array(...).fill(...): {remainder}"
        ));
    }

    let fill_value = resolve_typescript_scalar_expr(&fill_expr, bindings)?;
    let mut matrix: Vec<Vec<Option<JsonValue>>> = Vec::new();
    for _ in 0..rows {
        matrix.push((0..cols).map(|_| fill_value.clone()).collect());
    }
    Ok(matrix)
}

#[cfg(any(feature = "desktop", test))]
fn resolve_typescript_value_matrix_expr(
    expr: &str,
    bindings: &std::collections::HashMap<String, TypeScriptBinding>,
) -> Result<Vec<Vec<Option<JsonValue>>>, String> {
    let trimmed = expr.trim().trim_end_matches(';').trim();
    if trimmed.starts_with('[') {
        return parse_typescript_value_matrix(trimmed);
    }
    if trimmed.starts_with("Array.from") {
        return parse_typescript_array_from_fill_matrix(trimmed, bindings);
    }
    if is_typescript_identifier(trimmed) {
        match bindings.get(trimmed) {
            Some(TypeScriptBinding::Matrix(matrix)) => return Ok(matrix.clone()),
            Some(TypeScriptBinding::Scalar(_)) => {
                return Err(format!("expected matrix expression, got scalar {trimmed}"))
            }
            None => return Err(format!("unknown identifier: {trimmed}")),
        }
    }
    Err(format!(
        "unsupported TypeScript matrix expression: {trimmed}"
    ))
}

#[cfg(any(feature = "desktop", test))]
fn resolve_typescript_formulas_matrix_expr(
    expr: &str,
    bindings: &std::collections::HashMap<String, TypeScriptBinding>,
) -> Result<Vec<Vec<Option<String>>>, String> {
    let trimmed = expr.trim().trim_end_matches(';').trim();
    if trimmed.starts_with('[') {
        return parse_typescript_formulas_matrix(trimmed);
    }

    let matrix = resolve_typescript_value_matrix_expr(trimmed, bindings)?;
    let mut out: Vec<Vec<Option<String>>> = Vec::new();
    for row in matrix {
        let mut row_out: Vec<Option<String>> = Vec::new();
        for value in row {
            match value {
                None => row_out.push(None),
                Some(JsonValue::String(s)) => row_out.push(Some(s)),
                Some(other) => {
                    return Err(format!(
                        "expected formula string literal or null, got {other}"
                    ))
                }
            }
        }
        out.push(row_out);
    }
    Ok(out)
}

#[cfg(any(feature = "desktop", test))]
fn run_typescript_migration_script(state: &mut AppState, code: &str) -> TypeScriptRunResult {
    use std::collections::HashMap;

    let active_sheet_id = match state.get_workbook() {
        Ok(workbook) => {
            let active_index = state.macro_runtime_context().active_sheet;
            workbook
                .sheets
                .get(active_index)
                .or_else(|| workbook.sheets.first())
                .map(|s| s.id.clone())
                .unwrap_or_else(|| "Sheet1".to_string())
        }
        Err(err) => {
            return TypeScriptRunResult {
                ok: false,
                updates: Vec::new(),
                error: Some(err.to_string()),
            }
        }
    };

    let mut updates = Vec::<CellUpdateData>::new();
    let mut error: Option<String> = None;
    let mut bindings: HashMap<String, TypeScriptBinding> = HashMap::new();

    for raw_line in code.lines() {
        let line = raw_line.trim();
        if line.is_empty() {
            continue;
        }
        if line.starts_with("//") {
            continue;
        }
        if line.starts_with("export ") || line.starts_with("import ") {
            continue;
        }
        if line == "{" || line == "}" {
            continue;
        }
        if line.starts_with("const ") || line.starts_with("let ") || line.starts_with("var ") {
            let rest = if let Some(rest) = line.strip_prefix("const ") {
                rest
            } else if let Some(rest) = line.strip_prefix("let ") {
                rest
            } else {
                line.strip_prefix("var ").unwrap_or("")
            };

            if let Some((name_raw, expr_raw)) = rest.split_once('=') {
                let name = name_raw.trim();
                let expr = expr_raw.trim().trim_end_matches(';').trim();
                if is_typescript_identifier(name) {
                    if let Ok(value) = parse_typescript_value_expr(expr) {
                        bindings.insert(name.to_string(), TypeScriptBinding::Scalar(value));
                    } else if expr.trim_start().starts_with('[') {
                        if let Ok(matrix) = parse_typescript_value_matrix(expr) {
                            bindings.insert(name.to_string(), TypeScriptBinding::Matrix(matrix));
                        }
                    } else if expr.trim_start().starts_with("Array.from") {
                        if let Ok(matrix) = parse_typescript_array_from_fill_matrix(expr, &bindings)
                        {
                            bindings.insert(name.to_string(), TypeScriptBinding::Matrix(matrix));
                        }
                    }
                }
            }

            continue;
        }

        if let Some(idx) = line.find(".setValue(") {
            let target_expr = line[..idx].trim();
            let after = &line[idx + ".setValue(".len()..];
            match parse_typescript_call_args(after) {
                Ok((args, remainder)) => {
                    if !remainder
                        .trim_start()
                        .trim_start_matches(';')
                        .trim()
                        .is_empty()
                    {
                        error = Some(format!("unsupported setValue call: {line}"));
                        break;
                    }
                    match parse_typescript_range_selector(target_expr) {
                        Ok((start_row, start_col, end_row, end_col)) => {
                            if start_row != end_row || start_col != end_col {
                                error = Some(format!(
                                    "setValue is only valid for single cells (got range {start_row},{start_col}..{end_row},{end_col})"
                                ));
                                break;
                            }
                            match resolve_typescript_scalar_expr(&args, &bindings) {
                                Ok(value) => {
                                    match state.set_cell(
                                        &active_sheet_id,
                                        start_row,
                                        start_col,
                                        value,
                                        None,
                                    ) {
                                        Ok(mut changed) => updates.append(&mut changed),
                                        Err(e) => {
                                            error = Some(e.to_string());
                                            break;
                                        }
                                    }
                                }
                                Err(e) => {
                                    error = Some(e);
                                    break;
                                }
                            }
                        }
                        Err(e) => {
                            error = Some(e);
                            break;
                        }
                    }
                    continue;
                }
                Err(e) => {
                    error = Some(e);
                    break;
                }
            }
        }

        if let Some(idx) = line.find(".setValues(") {
            let target_expr = line[..idx].trim();
            let after = &line[idx + ".setValues(".len()..];
            match parse_typescript_call_args(after) {
                Ok((args, remainder)) => {
                    if !remainder
                        .trim_start()
                        .trim_start_matches(';')
                        .trim()
                        .is_empty()
                    {
                        error = Some(format!("unsupported setValues call: {line}"));
                        break;
                    }

                    let (start_row, start_col, end_row, end_col) =
                        match parse_typescript_range_selector(target_expr) {
                            Ok(coords) => coords,
                            Err(e) => {
                                error = Some(e);
                                break;
                            }
                        };

                    let matrix = match resolve_typescript_value_matrix_expr(&args, &bindings) {
                        Ok(m) => m,
                        Err(e) => {
                            error = Some(e);
                            break;
                        }
                    };

                    let row_count = end_row - start_row + 1;
                    let col_count = end_col - start_col + 1;
                    if matrix.len() != row_count || matrix.iter().any(|row| row.len() != col_count)
                    {
                        error = Some(format!(
                            "setValues expected {row_count}x{col_count} matrix for range ({start_row},{start_col})..({end_row},{end_col}), got {}x{}",
                            matrix.len(),
                            matrix.first().map(|row| row.len()).unwrap_or(0)
                        ));
                        break;
                    }

                    let mut payload: Vec<Vec<(Option<JsonValue>, Option<String>)>> = Vec::new();
                    for row in matrix {
                        payload.push(
                            row.into_iter()
                                .map(|value| (value, None))
                                .collect::<Vec<_>>(),
                        );
                    }

                    match state.set_range(
                        &active_sheet_id,
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                        payload,
                    ) {
                        Ok(mut changed) => updates.append(&mut changed),
                        Err(e) => {
                            error = Some(e.to_string());
                            break;
                        }
                    }

                    continue;
                }
                Err(e) => {
                    error = Some(e);
                    break;
                }
            }
        }

        if let Some(idx) = line.find(".setFormulas(") {
            let target_expr = line[..idx].trim();
            let after = &line[idx + ".setFormulas(".len()..];
            match parse_typescript_call_args(after) {
                Ok((args, remainder)) => {
                    if !remainder
                        .trim_start()
                        .trim_start_matches(';')
                        .trim()
                        .is_empty()
                    {
                        error = Some(format!("unsupported setFormulas call: {line}"));
                        break;
                    }
                    let (start_row, start_col, end_row, end_col) =
                        match parse_typescript_range_selector(target_expr) {
                            Ok(coords) => coords,
                            Err(e) => {
                                error = Some(e);
                                break;
                            }
                        };

                    let matrix = match resolve_typescript_formulas_matrix_expr(&args, &bindings) {
                        Ok(m) => m,
                        Err(e) => {
                            error = Some(e);
                            break;
                        }
                    };

                    let row_count = end_row - start_row + 1;
                    let col_count = end_col - start_col + 1;
                    if matrix.len() != row_count || matrix.iter().any(|row| row.len() != col_count)
                    {
                        error = Some(format!(
                            "setFormulas expected {row_count}x{col_count} matrix for range ({start_row},{start_col})..({end_row},{end_col}), got {}x{}",
                            matrix.len(),
                            matrix.first().map(|row| row.len()).unwrap_or(0)
                        ));
                        break;
                    }

                    let mut payload: Vec<Vec<(Option<JsonValue>, Option<String>)>> = Vec::new();
                    for row in matrix {
                        payload.push(
                            row.into_iter()
                                .map(|formula| (None, formula))
                                .collect::<Vec<_>>(),
                        );
                    }

                    match state.set_range(
                        &active_sheet_id,
                        start_row,
                        start_col,
                        end_row,
                        end_col,
                        payload,
                    ) {
                        Ok(mut changed) => updates.append(&mut changed),
                        Err(e) => {
                            error = Some(e.to_string());
                            break;
                        }
                    }

                    continue;
                }
                Err(e) => {
                    error = Some(e);
                    break;
                }
            }
        }

        let Some((lhs_raw, rhs_raw)) = line.split_once('=') else {
            // Ignore non-assignment statements (loops, conditionals, etc.).
            continue;
        };
        let lhs_raw = lhs_raw.trim();
        let rhs_raw = rhs_raw.trim();

        let (assign_kind, lhs_target) = if let Some(prefix) = lhs_raw.strip_suffix(".value") {
            ("value", prefix.trim())
        } else if let Some(prefix) = lhs_raw.strip_suffix(".formula") {
            ("formula", prefix.trim())
        } else {
            continue;
        };

        let (row, col) = match parse_typescript_cell_selector(lhs_target) {
            Ok(coords) => coords,
            Err(_) => continue,
        };

        let result = match assign_kind {
            "value" => match resolve_typescript_scalar_expr(rhs_raw, &bindings) {
                Ok(value) => state.set_cell(&active_sheet_id, row, col, value, None),
                Err(e) => Err(AppStateError::WhatIf(e)),
            },
            "formula" => match resolve_typescript_scalar_expr(rhs_raw, &bindings) {
                Ok(None) => state.set_cell(&active_sheet_id, row, col, None, None),
                Ok(Some(JsonValue::String(formula))) => {
                    state.set_cell(&active_sheet_id, row, col, None, Some(formula))
                }
                Ok(Some(other)) => Err(AppStateError::WhatIf(format!(
                    "expected formula string literal or null, got {other}"
                ))),
                Err(e) => Err(AppStateError::WhatIf(e)),
            },
            _ => continue,
        };

        match result {
            Ok(mut changed) => updates.append(&mut changed),
            Err(e) => {
                error = Some(e.to_string());
                break;
            }
        }
    }

    // De-dupe updates by last write (keep report stable).
    let mut out: Vec<CellUpdateData> = Vec::new();
    let mut idx_by_key: HashMap<(String, usize, usize), usize> = HashMap::new();
    for update in updates {
        let key = (update.sheet_id.clone(), update.row, update.col);
        if let Some(idx) = idx_by_key.get(&key).copied() {
            out[idx] = update;
        } else {
            idx_by_key.insert(key, out.len());
            out.push(update);
        }
    }

    TypeScriptRunResult {
        ok: error.is_none(),
        updates: out.into_iter().map(cell_update_from_state).collect(),
        error,
    }
}

#[cfg(feature = "desktop")]
#[tauri::command]
pub async fn validate_vba_migration(
    workbook_id: Option<String>,
    macro_id: String,
    target: MigrationTarget,
    code: String,
    state: State<'_, SharedAppState>,
    trust: State<'_, SharedMacroTrustStore>,
) -> Result<MigrationValidationReport, String> {
    let workbook_id_str = workbook_id.clone();
    let shared = state.inner().clone();
    let trust_shared = trust.inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        use crate::macros::MacroExecutionOptions;
        use std::collections::BTreeSet;

        let (vba_blocked_result, workbook, macro_ctx) = {
            let mut state = shared.lock().unwrap();
            let trust_store = trust_shared.lock().unwrap();

            let blocked = {
                let workbook_id = workbook_id_str.as_deref();
                let workbook = state.get_workbook_mut().map_err(app_error)?;
                enforce_macro_trust(workbook, workbook_id, &trust_store)?
            };

            let vba_blocked_result = blocked.map(macro_blocked_result);
            let macro_ctx = state.macro_runtime_context();
            let workbook = state.get_workbook_mut().map_err(app_error)?.clone();
            (vba_blocked_result, workbook, macro_ctx)
        };

        let mut vba_state = AppState::new();
        vba_state.load_workbook(workbook.clone());
        vba_state
            .set_macro_runtime_context(macro_ctx)
            .map_err(|e| e.to_string())?;

        let mut script_state = AppState::new();
        script_state.load_workbook(workbook);
        script_state
            .set_macro_runtime_context(macro_ctx)
            .map_err(|e| e.to_string())?;

        let vba = if let Some(blocked_result) = vba_blocked_result {
            blocked_result
        } else {
            match vba_state.run_macro(&macro_id, MacroExecutionOptions::default()) {
                Ok(outcome) => macro_result_from_outcome(outcome),
                Err(err) => MacroRunResult {
                    ok: false,
                    output: Vec::new(),
                    updates: Vec::new(),
                    error: Some(MacroError {
                        message: err.to_string(),
                        code: Some("macro_error".to_string()),
                        blocked: None,
                    }),
                    permission_request: None,
                },
            }
        };

        let mut python = None;
        let mut typescript = None;

        match target {
            MigrationTarget::Python => {
                let python_context = {
                    let workbook = script_state.get_workbook().map_err(|e| e.to_string())?;
                    let fallback_sheet_id = workbook
                        .sheets
                        .first()
                        .map(|s| s.id.clone())
                        .ok_or_else(|| "workbook contains no sheets".to_string())?;
                    let active_sheet_id = workbook
                        .sheets
                        .get(macro_ctx.active_sheet)
                        .map(|s| s.id.clone())
                        .unwrap_or_else(|| fallback_sheet_id.clone());

                    let selection =
                        macro_ctx
                            .selection
                            .unwrap_or(formula_vba_runtime::VbaRangeRef {
                                sheet: macro_ctx.active_sheet,
                                start_row: macro_ctx.active_cell.0,
                                start_col: macro_ctx.active_cell.1,
                                end_row: macro_ctx.active_cell.0,
                                end_col: macro_ctx.active_cell.1,
                            });
                    let selection_sheet_id = workbook
                        .sheets
                        .get(selection.sheet)
                        .map(|s| s.id.clone())
                        .unwrap_or_else(|| active_sheet_id.clone());
                    PythonRunContext {
                        active_sheet_id: Some(active_sheet_id.clone()),
                        selection: Some(PythonSelection {
                            sheet_id: selection_sheet_id,
                            start_row: selection.start_row.saturating_sub(1) as usize,
                            start_col: selection.start_col.saturating_sub(1) as usize,
                            end_row: selection.end_row.saturating_sub(1) as usize,
                            end_col: selection.end_col.saturating_sub(1) as usize,
                        }),
                    }
                };
                python = Some(
                    crate::python::run_python_script(
                        &mut script_state,
                        &code,
                        None,
                        None,
                        None,
                        Some(python_context),
                    )
                    .map_err(|e| e.to_string())?,
                );
            }
            MigrationTarget::TypeScript => {
                typescript = Some(run_typescript_migration_script(&mut script_state, &code));
            }
        };

        let mut mismatches = Vec::new();

        let mut touched = BTreeSet::<(String, usize, usize)>::new();
        for update in &vba.updates {
            touched.insert((update.sheet_id.clone(), update.row, update.col));
        }
        if let Some(python_run) = python.as_ref() {
            for update in &python_run.updates {
                touched.insert((update.sheet_id.clone(), update.row, update.col));
            }
        }
        if let Some(ts_run) = typescript.as_ref() {
            for update in &ts_run.updates {
                touched.insert((update.sheet_id.clone(), update.row, update.col));
            }
        }

        for (sheet_id, row, col) in touched {
            let vba_cell = cell_value_from_state(&vba_state, &sheet_id, row, col)?;
            let script_cell = cell_value_from_state(&script_state, &sheet_id, row, col)?;
            if vba_cell != script_cell {
                mismatches.push(MigrationValidationMismatch {
                    sheet_id,
                    row,
                    col,
                    vba: vba_cell,
                    script: script_cell,
                });
            }
        }

        let script_ok = match target {
            MigrationTarget::Python => python.as_ref().map(|r| r.ok).unwrap_or(false),
            MigrationTarget::TypeScript => typescript.as_ref().map(|r| r.ok).unwrap_or(false),
        };

        let mut error_messages: Vec<String> = Vec::new();
        if !vba.ok {
            if let Some(err) = vba.error.as_ref() {
                error_messages.push(err.message.clone());
            } else {
                error_messages.push("VBA macro failed".to_string());
            }
        }
        if !script_ok {
            match target {
                MigrationTarget::Python => {
                    if let Some(run) = python.as_ref() {
                        if let Some(err) = run.error.as_ref() {
                            error_messages.push(err.message.clone());
                        } else {
                            error_messages.push("Python script failed".to_string());
                        }
                    }
                }
                MigrationTarget::TypeScript => {
                    if let Some(run) = typescript.as_ref() {
                        if let Some(err) = run.error.as_ref() {
                            error_messages.push(err.clone());
                        } else {
                            error_messages.push("TypeScript migration failed".to_string());
                        }
                    }
                }
            }
        }
        let error = if error_messages.is_empty() {
            None
        } else {
            Some(error_messages.join(" | "))
        };

        let ok = vba.ok && script_ok && mismatches.is_empty();

        Ok(MigrationValidationReport {
            ok,
            macro_id,
            target,
            mismatches,
            vba,
            python,
            typescript,
            error,
        })
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

#[cfg(feature = "desktop")]
#[tauri::command]
pub fn quit_app() {
    // We intentionally use a hard process exit here. The desktop shell already delegates
    // "should we quit?" decisions (event macros + unsaved prompts) to the frontend.
    // Once the frontend invokes this command, exiting immediately avoids re-entering the
    // CloseRequested handler (which prevents default close to support hide-to-tray).
    std::process::exit(0);
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

    #[test]
    fn typescript_migration_interpreter_applies_basic_range_and_cell_assignments() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        let mut state = crate::state::AppState::new();
        state.load_workbook(workbook);

        let code = r#"
export default async function main(ctx) {
  const sheet = ctx.activeSheet;
  const fill = 7;
  const values = Array.from({ length: 2 }, () => Array(2).fill(fill));
  await sheet.getRange("$A$1:$B$2").setValues(values);
  await sheet.getRange("C1").setValue(fill);

  const formula = "=A1+B1";
  await sheet.getRange("D1:E1").setFormulas(Array.from({ length: 1 }, () => Array(2).fill(formula)));
 }
"#;

        let result = run_typescript_migration_script(&mut state, code);
        assert!(result.ok, "expected ok, got {:?}", result.error);

        let a1 = state.get_cell("Sheet1", 0, 0).expect("A1 exists");
        assert_eq!(a1.value.display(), "7");

        let b2 = state.get_cell("Sheet1", 1, 1).expect("B2 exists");
        assert_eq!(b2.value.display(), "7");

        let c1 = state.get_cell("Sheet1", 0, 2).expect("C1 exists");
        assert_eq!(c1.value.display(), "7");

        let d1 = state.get_cell("Sheet1", 0, 3).expect("D1 exists");
        assert_eq!(d1.formula.as_deref(), Some("=A1+B1"));
        assert_eq!(d1.value.display(), "14");

        let e1 = state.get_cell("Sheet1", 0, 4).expect("E1 exists");
        assert_eq!(e1.formula.as_deref(), Some("=A1+B1"));
        assert_eq!(e1.value.display(), "14");
    }

    #[test]
    fn typescript_migration_interpreter_respects_active_sheet_from_macro_context() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());

        let mut state = crate::state::AppState::new();
        state.load_workbook(workbook);

        // Simulate the user having Sheet2 selected when kicking off a migration validation.
        state
            .set_macro_ui_context("Sheet2", 0, 0, None)
            .expect("set macro ui context");

        let code = r#"
export default async function main(ctx) {
  const sheet = ctx.activeSheet;
  sheet.range("A1").value = 99;
}
"#;

        let result = run_typescript_migration_script(&mut state, code);
        assert!(result.ok, "expected ok, got {:?}", result.error);

        let sheet2_a1 = state.get_cell("Sheet2", 0, 0).expect("Sheet2!A1 exists");
        assert_eq!(sheet2_a1.value.display(), "99");

        let sheet1_a1 = state.get_cell("Sheet1", 0, 0).expect("Sheet1!A1 exists");
        assert_eq!(
            sheet1_a1.value.display(),
            "",
            "expected Sheet1!A1 to remain empty"
        );
    }
}
