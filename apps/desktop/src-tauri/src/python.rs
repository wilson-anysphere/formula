use crate::commands::{
    CellUpdate, PythonError, PythonFilesystemPermission, PythonNetworkPermission,
    PythonPermissions, PythonRunContext, PythonRunResult, PythonSelection,
};
use crate::resource_limits::{MAX_RANGE_CELLS_PER_CALL, MAX_RANGE_DIM};
use crate::sheet_name::sheet_name_eq_case_insensitive;
use crate::state::{AppState, AppStateError, CellUpdateData};
use serde::Deserialize;
use serde_json::Value as JsonValue;
use std::collections::HashMap;
use std::ffi::OsString;
use std::io::{BufRead, BufReader, Read, Write};
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{mpsc, Arc, Mutex};
use std::thread;
use std::time::Duration;

const DEFAULT_TIMEOUT_MS: u64 = 5_000;
const DEFAULT_MAX_MEMORY_BYTES: u64 = 256 * 1024 * 1024;
const PYTHON_PERMISSION_ESCALATION_ERROR: &str =
    "Python permission escalation is not supported yet";

fn unsafe_python_permissions_enabled() -> bool {
    // Escape hatch for local development only.
    //
    // Release builds should not permit enabling elevated filesystem/network access
    // from the frontend since the native Python sandbox is not a hardened boundary.
    if !cfg!(debug_assertions) {
        return false;
    }

    let value = std::env::var("FORMULA_UNSAFE_PYTHON_PERMISSIONS").unwrap_or_default();
    matches!(value.trim(), "1" | "true" | "TRUE" | "yes" | "YES")
}

#[derive(Debug, Deserialize)]
#[serde(tag = "type")]
enum RunnerMessage {
    #[serde(rename = "rpc")]
    Rpc {
        id: u64,
        method: String,
        #[serde(default)]
        params: Option<JsonValue>,
    },
    #[serde(rename = "result")]
    Result {
        success: bool,
        #[serde(default)]
        error: Option<String>,
        #[serde(default)]
        traceback: Option<String>,
    },
}

#[derive(Debug, serde::Serialize)]
struct ExecuteMessage<'a> {
    #[serde(rename = "type")]
    msg_type: &'static str,
    code: &'a str,
    permissions: PythonPermissions,
    timeout_ms: u64,
    max_memory_bytes: u64,
}

#[derive(Debug, serde::Serialize)]
struct RpcResponseMessage {
    #[serde(rename = "type")]
    msg_type: &'static str,
    id: u64,
    result: JsonValue,
    error: Option<String>,
}

#[derive(Clone, Debug, Deserialize)]
struct RangeRef {
    sheet_id: String,
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
}

fn resolve_repo_root() -> PathBuf {
    // apps/desktop/src-tauri -> repo root
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../..")
}

fn resolve_formula_api_path() -> PathBuf {
    resolve_repo_root().join("python").join("formula_api")
}

fn python_path_env(formula_api_path: &Path) -> OsString {
    let mut paths: Vec<PathBuf> = Vec::new();
    paths.push(formula_api_path.to_path_buf());
    if let Some(existing) = std::env::var_os("PYTHONPATH") {
        paths.extend(std::env::split_paths(&existing));
    }
    std::env::join_paths(paths).unwrap_or_else(|_| formula_api_path.as_os_str().to_os_string())
}

fn cell_update_from_state(update: CellUpdateData) -> CellUpdate {
    CellUpdate {
        sheet_id: update.sheet_id,
        row: update.row,
        col: update.col,
        value: update.value.as_json(),
        formula: update.formula,
        display_value: update.display_value,
    }
}

fn dedupe_updates(updates: Vec<CellUpdateData>) -> Vec<CellUpdateData> {
    let mut out = Vec::new();
    let mut index_by_key = HashMap::<(String, usize, usize), usize>::new();
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

fn normalize_formula_text(raw: &str) -> Option<String> {
    let display = formula_model::display_formula_text(raw);
    if display.is_empty() {
        None
    } else {
        Some(display)
    }
}

fn parse_cell_input(input: &JsonValue) -> (Option<JsonValue>, Option<String>) {
    match input {
        JsonValue::Null => (None, None),
        JsonValue::String(s) => {
            if let Some(rest) = s.strip_prefix('\'') {
                return (Some(JsonValue::String(rest.to_string())), None);
            }
            let trimmed = s.trim_start();
            if trimmed.starts_with('=') {
                return (None, normalize_formula_text(trimmed));
            }
            (Some(JsonValue::String(s.clone())), None)
        }
        JsonValue::Bool(_) | JsonValue::Number(_) => (Some(input.clone()), None),
        other => (Some(other.clone()), None),
    }
}

struct PythonRpcHost<'a> {
    state: &'a mut AppState,
    active_sheet_id: String,
    selection: PythonSelection,
    updates: Vec<CellUpdateData>,
}

impl<'a> PythonRpcHost<'a> {
    fn new(state: &'a mut AppState, context: Option<PythonRunContext>) -> Result<Self, String> {
        let workbook = state.get_workbook().map_err(|e| e.to_string())?;
        let fallback_sheet_id = workbook
            .sheets
            .first()
            .map(|s| s.id.clone())
            .ok_or_else(|| "workbook contains no sheets".to_string())?;

        let (active_sheet_id, selection) = match context {
            Some(ctx) => {
                let active_sheet_id = ctx
                    .active_sheet_id
                    .filter(|id| workbook.sheet(id).is_some())
                    .unwrap_or_else(|| fallback_sheet_id.clone());
                let selection = ctx.selection.and_then(|sel| {
                    if workbook.sheet(&sel.sheet_id).is_some() {
                        Some(sel)
                    } else {
                        None
                    }
                });
                let selection = selection.unwrap_or_else(|| PythonSelection {
                    sheet_id: active_sheet_id.clone(),
                    start_row: 0,
                    start_col: 0,
                    end_row: 0,
                    end_col: 0,
                });
                (active_sheet_id, selection)
            }
            None => {
                let active_sheet_id = fallback_sheet_id.clone();
                let selection = PythonSelection {
                    sheet_id: active_sheet_id.clone(),
                    start_row: 0,
                    start_col: 0,
                    end_row: 0,
                    end_col: 0,
                };
                (active_sheet_id, selection)
            }
        };

        Ok(Self {
            state,
            active_sheet_id,
            selection,
            updates: Vec::new(),
        })
    }

    fn extend_updates(&mut self, updates: Vec<CellUpdateData>) {
        self.updates.extend(updates);
        self.updates = dedupe_updates(std::mem::take(&mut self.updates));
    }

    fn take_updates(&mut self) -> Vec<CellUpdateData> {
        std::mem::take(&mut self.updates)
    }

    fn parse_range(params: &JsonValue) -> Result<RangeRef, String> {
        let range = params
            .get("range")
            .ok_or_else(|| "missing params.range".to_string())?;
        serde_json::from_value::<RangeRef>(range.clone())
            .map_err(|e| format!("invalid range reference: {e}"))
    }

    fn ensure_single_cell(range: &RangeRef, method: &str) -> Result<(), String> {
        if range.start_row != range.end_row || range.start_col != range.end_col {
            return Err(format!("{method} expects a single cell range"));
        }
        Ok(())
    }

    fn enforce_range_limits(range: &RangeRef) -> Result<(usize, usize), String> {
        if range.start_row > range.end_row || range.start_col > range.end_col {
            return Err(AppStateError::InvalidRange {
                start_row: range.start_row,
                start_col: range.start_col,
                end_row: range.end_row,
                end_col: range.end_col,
            }
            .to_string());
        }

        let row_count = range
            .end_row
            .checked_sub(range.start_row)
            .and_then(|d| d.checked_add(1))
            .unwrap_or(usize::MAX);
        let col_count = range
            .end_col
            .checked_sub(range.start_col)
            .and_then(|d| d.checked_add(1))
            .unwrap_or(usize::MAX);

        if row_count > MAX_RANGE_DIM || col_count > MAX_RANGE_DIM {
            return Err(AppStateError::RangeDimensionTooLarge {
                rows: row_count,
                cols: col_count,
                limit: MAX_RANGE_DIM,
            }
            .to_string());
        }

        let cell_count = (row_count as u128) * (col_count as u128);
        if cell_count > MAX_RANGE_CELLS_PER_CALL as u128 {
            return Err(AppStateError::RangeTooLarge {
                rows: row_count,
                cols: col_count,
                limit: MAX_RANGE_CELLS_PER_CALL,
            }
            .to_string());
        }

        Ok((row_count, col_count))
    }

    fn handle_rpc(&mut self, method: &str, params: Option<JsonValue>) -> Result<JsonValue, String> {
        let params = params.unwrap_or(JsonValue::Null);
        match method {
            "get_active_sheet_id" => Ok(JsonValue::String(self.active_sheet_id.clone())),
            "get_sheet_id" => {
                let name = params
                    .get("name")
                    .and_then(|v| v.as_str())
                    .unwrap_or_default();
                let workbook = self.state.get_workbook().map_err(|e| e.to_string())?;
                let found = workbook
                    .sheets
                    .iter()
                    .find(|s| sheet_name_eq_case_insensitive(&s.name, name))
                    .map(|s| JsonValue::String(s.id.clone()))
                    .unwrap_or(JsonValue::Null);
                Ok(found)
            }
            "create_sheet" => {
                let name = params
                    .get("name")
                    .and_then(|v| v.as_str())
                    .unwrap_or_default()
                    .trim()
                    .to_string();
                formula_model::validate_sheet_name(&name).map_err(|e| e.to_string())?;

                {
                    let workbook = self.state.get_workbook().map_err(|e| e.to_string())?;
                    if workbook
                        .sheets
                        .iter()
                        .any(|sheet| sheet_name_eq_case_insensitive(&sheet.name, &name))
                    {
                        return Err(formula_model::SheetNameError::DuplicateName.to_string());
                    }
                }

                let index = params
                    .get("index")
                    .and_then(|v| v.as_u64())
                    .map(|v| v as usize);

                let insert_after = index.is_none().then(|| self.active_sheet_id.clone());

                let sheet = self
                    .state
                    .add_sheet(name, None, insert_after, index)
                    .map_err(|e| e.to_string())?;
                Ok(JsonValue::String(sheet.id))
            }
            "get_sheet_name" => {
                let sheet_id = params
                    .get("sheet_id")
                    .and_then(|v| v.as_str())
                    .ok_or_else(|| "get_sheet_name expects { sheet_id }".to_string())?;
                let workbook = self.state.get_workbook().map_err(|e| e.to_string())?;
                let sheet = workbook
                    .sheet(sheet_id)
                    .ok_or_else(|| AppStateError::UnknownSheet(sheet_id.to_string()).to_string())?;
                Ok(JsonValue::String(sheet.name.clone()))
            }
            "rename_sheet" => {
                let sheet_id = params
                    .get("sheet_id")
                    .and_then(|v| v.as_str())
                    .ok_or_else(|| "rename_sheet expects { sheet_id, name }".to_string())?;
                let name = params
                    .get("name")
                    .and_then(|v| v.as_str())
                    .ok_or_else(|| "rename_sheet expects { sheet_id, name }".to_string())?
                    .to_string();
                self.state
                    .rename_sheet(sheet_id, name)
                    .map_err(|e| match e {
                        // Match `create_sheet` behavior: surface canonical sheet-name validation
                        // strings without the AppStateError wrapper prefix.
                        AppStateError::WhatIf(msg) => msg,
                        other => other.to_string(),
                    })?;
                Ok(JsonValue::Null)
            }
            "get_selection" => Ok(serde_json::to_value(&self.selection).unwrap_or(JsonValue::Null)),
            "set_selection" => {
                let selection = params
                    .get("selection")
                    .ok_or_else(|| "set_selection expects { selection }".to_string())?;
                let parsed: PythonSelection = serde_json::from_value(selection.clone())
                    .map_err(|e| format!("invalid selection: {e}"))?;
                // Best effort validation.
                if self
                    .state
                    .get_workbook()
                    .map_err(|e| e.to_string())?
                    .sheet(&parsed.sheet_id)
                    .is_none()
                {
                    return Err(format!("unknown sheet id: {}", parsed.sheet_id));
                }
                self.active_sheet_id = parsed.sheet_id.clone();
                self.selection = parsed;
                Ok(JsonValue::Null)
            }
            "get_range_values" => {
                let range = Self::parse_range(&params)?;
                let (row_count, col_count) = Self::enforce_range_limits(&range)?;
                let workbook = self.state.get_workbook().map_err(|e| e.to_string())?;
                let sheet = workbook.sheet(&range.sheet_id).ok_or_else(|| {
                    AppStateError::UnknownSheet(range.sheet_id.clone()).to_string()
                })?;

                let mut out: Vec<Vec<JsonValue>> = Vec::with_capacity(row_count);
                for row in range.start_row..=range.end_row {
                    let mut row_vals = Vec::with_capacity(col_count);
                    for col in range.start_col..=range.end_col {
                        let cell = sheet.get_cell(row, col);
                        row_vals.push(cell.computed_value.as_json().unwrap_or(JsonValue::Null));
                    }
                    out.push(row_vals);
                }
                Ok(serde_json::to_value(out).unwrap_or(JsonValue::Null))
            }
            "set_range_values" => {
                let range = Self::parse_range(&params)?;
                let values = params.get("values").unwrap_or(&JsonValue::Null);
                let (row_count, col_count) = Self::enforce_range_limits(&range)?;

                let normalized: Vec<Vec<(Option<JsonValue>, Option<String>)>> = match values {
                    JsonValue::Array(rows)
                        if rows
                            .first()
                            .is_some_and(|v| matches!(v, JsonValue::Array(_))) =>
                    {
                        (0..row_count)
                            .map(|r| {
                                let row = rows.get(r).and_then(|v| v.as_array());
                                (0..col_count)
                                    .map(|c| {
                                        let cell = row
                                            .and_then(|rv| rv.get(c))
                                            .unwrap_or(&JsonValue::Null);
                                        parse_cell_input(cell)
                                    })
                                    .collect::<Vec<_>>()
                            })
                            .collect::<Vec<_>>()
                    }
                    scalar => {
                        let parsed = parse_cell_input(scalar);
                        vec![vec![parsed; col_count]; row_count]
                    }
                };

                let edits = normalized
                    .into_iter()
                    .map(|row| {
                        row.into_iter()
                            .map(|(value, formula)| (value, formula))
                            .collect::<Vec<_>>()
                    })
                    .collect::<Vec<_>>();

                let updates = self
                    .state
                    .set_range(
                        &range.sheet_id,
                        range.start_row,
                        range.start_col,
                        range.end_row,
                        range.end_col,
                        edits,
                    )
                    .map_err(|e| e.to_string())?;
                self.extend_updates(updates);
                Ok(JsonValue::Null)
            }
            "set_cell_value" => {
                let range = Self::parse_range(&params)?;
                Self::ensure_single_cell(&range, method)?;
                let value = params.get("value").unwrap_or(&JsonValue::Null);
                let (value, formula) = parse_cell_input(value);
                let updates = self
                    .state
                    .set_cell(
                        &range.sheet_id,
                        range.start_row,
                        range.start_col,
                        value,
                        formula,
                    )
                    .map_err(|e| e.to_string())?;
                self.extend_updates(updates);
                Ok(JsonValue::Null)
            }
            "get_cell_formula" => {
                let range = Self::parse_range(&params)?;
                Self::ensure_single_cell(&range, method)?;
                let workbook = self.state.get_workbook().map_err(|e| e.to_string())?;
                let sheet = workbook.sheet(&range.sheet_id).ok_or_else(|| {
                    AppStateError::UnknownSheet(range.sheet_id.clone()).to_string()
                })?;
                let cell = sheet.get_cell(range.start_row, range.start_col);
                Ok(cell
                    .formula
                    .map(JsonValue::String)
                    .unwrap_or(JsonValue::Null))
            }
            "set_cell_formula" => {
                let range = Self::parse_range(&params)?;
                Self::ensure_single_cell(&range, method)?;
                let raw = params
                    .get("formula")
                    .and_then(|v| v.as_str())
                    .ok_or_else(|| "set_cell_formula expects { range, formula }".to_string())?;
                let formula =
                    normalize_formula_text(raw).ok_or_else(|| "empty formula".to_string())?;
                let updates = self
                    .state
                    .set_cell(
                        &range.sheet_id,
                        range.start_row,
                        range.start_col,
                        None,
                        Some(formula),
                    )
                    .map_err(|e| e.to_string())?;
                self.extend_updates(updates);
                Ok(JsonValue::Null)
            }
            "clear_range" => {
                let range = Self::parse_range(&params)?;
                let updates = self
                    .state
                    .clear_range(
                        &range.sheet_id,
                        range.start_row,
                        range.start_col,
                        range.end_row,
                        range.end_col,
                    )
                    .map_err(|e| e.to_string())?;
                self.extend_updates(updates);
                Ok(JsonValue::Null)
            }
            "set_range_format" => {
                let range = Self::parse_range(&params)?;
                let format = params.get("format").unwrap_or(&JsonValue::Null);
                let patch = parse_number_format_patch(format)?;
                if let Some(number_format) = patch {
                    self.state
                        .set_range_number_format(
                            &range.sheet_id,
                            range.start_row,
                            range.start_col,
                            range.end_row,
                            range.end_col,
                            number_format,
                        )
                        .map_err(|e| e.to_string())?;
                }
                Ok(JsonValue::Null)
            }
            "get_range_format" => {
                let range = Self::parse_range(&params)?;
                let workbook = self.state.get_workbook().map_err(|e| e.to_string())?;
                let sheet = workbook.sheet(&range.sheet_id).ok_or_else(|| {
                    AppStateError::UnknownSheet(range.sheet_id.clone()).to_string()
                })?;
                let cell = sheet.get_cell(range.start_row, range.start_col);
                let Some(fmt) = cell.number_format else {
                    return Ok(JsonValue::Object(Default::default()));
                };
                let mut out = serde_json::Map::new();
                out.insert("numberFormat".to_string(), JsonValue::String(fmt));
                Ok(JsonValue::Object(out))
            }
            other => Err(format!("Unknown RPC method: {other}")),
        }
    }
}

fn parse_number_format_patch(format: &JsonValue) -> Result<Option<Option<String>>, String> {
    match format {
        JsonValue::Null => Ok(Some(None)),
        JsonValue::String(s) => Ok(Some(Some(s.clone()))),
        JsonValue::Object(map) => {
            let value = map
                .get("numberFormat")
                .or_else(|| map.get("number_format"))
                .or_else(|| map.get("numberformat"));
            let Some(value) = value else {
                return Ok(None);
            };
            match value {
                JsonValue::Null => Ok(Some(None)),
                JsonValue::String(s) => Ok(Some(Some(s.clone()))),
                _ => Ok(None),
            }
        }
        _ => Ok(None),
    }
}

struct TimeoutGuard {
    done_tx: Option<mpsc::Sender<()>>,
    handle: Option<thread::JoinHandle<()>>,
}

impl TimeoutGuard {
    fn new(
        timeout_ms: u64,
        child: Arc<Mutex<std::process::Child>>,
        timed_out: Arc<AtomicBool>,
    ) -> Self {
        if timeout_ms == 0 {
            return Self {
                done_tx: None,
                handle: None,
            };
        }
        let (tx, rx) = mpsc::channel::<()>();
        let handle = thread::spawn(move || {
            if rx.recv_timeout(Duration::from_millis(timeout_ms)).is_ok() {
                return;
            }
            timed_out.store(true, Ordering::SeqCst);
            if let Ok(mut child) = child.lock() {
                let _ = child.kill();
            }
        });
        Self {
            done_tx: Some(tx),
            handle: Some(handle),
        }
    }
}

impl Drop for TimeoutGuard {
    fn drop(&mut self) {
        if let Some(tx) = self.done_tx.take() {
            let _ = tx.send(());
        }
        if let Some(handle) = self.handle.take() {
            let _ = handle.join();
        }
    }
}

pub fn run_python_script(
    state: &mut AppState,
    code: &str,
    permissions: Option<PythonPermissions>,
    timeout_ms: Option<u64>,
    max_memory_bytes: Option<u64>,
    context: Option<PythonRunContext>,
) -> Result<PythonRunResult, String> {
    crate::ipc_limits::enforce_script_code_size(code)?;
    let permissions = permissions.unwrap_or_default();
    if !unsafe_python_permissions_enabled()
        && (permissions.filesystem != PythonFilesystemPermission::None
            || permissions.network != PythonNetworkPermission::None)
    {
        return Err(PYTHON_PERMISSION_ESCALATION_ERROR.to_string());
    }
    let timeout_ms = timeout_ms.unwrap_or(DEFAULT_TIMEOUT_MS);
    let max_memory_bytes = max_memory_bytes.unwrap_or(DEFAULT_MAX_MEMORY_BYTES);

    let repo_root = resolve_repo_root();
    let formula_api_path = resolve_formula_api_path();

    let candidates: Vec<String> = match std::env::var("FORMULA_PYTHON_EXECUTABLE") {
        Ok(explicit) if !explicit.trim().is_empty() => vec![explicit],
        _ => vec!["python3".to_string(), "python".to_string()],
    };

    let mut last_err: Option<std::io::Error> = None;
    let mut child: Option<std::process::Child> = None;
    let mut selected: Option<String> = None;

    for exe in &candidates {
        let mut cmd = Command::new(exe);
        cmd.args(["-u", "-m", "formula.runtime.stdio_runner"])
            .current_dir(&repo_root)
            .env("PYTHONPATH", python_path_env(&formula_api_path))
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::piped());

        match cmd.spawn() {
            Ok(spawned) => {
                child = Some(spawned);
                selected = Some(exe.clone());
                break;
            }
            Err(err) => {
                last_err = Some(err);
            }
        }
    }

    let mut child = child.ok_or_else(|| {
        let err = last_err
            .map(|e| e.to_string())
            .unwrap_or_else(|| "unknown error".to_string());
        format!("Failed to spawn python runner (tried {candidates:?}): {err}")
    })?;

    let mut stdin = child
        .stdin
        .take()
        .ok_or_else(|| "Failed to open python stdin".to_string())?;
    let stdout = child
        .stdout
        .take()
        .ok_or_else(|| "Failed to open python stdout".to_string())?;
    let stderr = child
        .stderr
        .take()
        .ok_or_else(|| "Failed to open python stderr".to_string())?;

    let stderr_buf = Arc::new(Mutex::new(String::new()));
    let stderr_buf_clone = stderr_buf.clone();
    let stderr_thread = thread::spawn(move || {
        let mut reader = BufReader::new(stderr);
        let mut out = String::new();
        let _ = reader.read_to_string(&mut out);
        if let Ok(mut guard) = stderr_buf_clone.lock() {
            guard.push_str(&out);
        }
    });

    let _python_exe = selected.unwrap_or_else(|| candidates[0].clone());
    let child = Arc::new(Mutex::new(child));
    let timed_out = Arc::new(AtomicBool::new(false));
    let _timeout_guard = TimeoutGuard::new(timeout_ms, child.clone(), timed_out.clone());

    // Kick off execution once listeners are attached.
    let exec = ExecuteMessage {
        msg_type: "execute",
        code,
        permissions,
        timeout_ms,
        max_memory_bytes,
    };
    serde_json::to_writer(&mut stdin, &exec).map_err(|e| e.to_string())?;
    stdin.write_all(b"\n").map_err(|e| e.to_string())?;
    stdin.flush().map_err(|e| e.to_string())?;

    let mut host = PythonRpcHost::new(state, context)?;
    let mut runner_result: Option<(bool, Option<String>, Option<String>)> = None;

    let reader = BufReader::new(stdout);
    for line in reader.lines() {
        let line = line.map_err(|e| e.to_string())?;
        let trimmed = line.trim();
        if trimmed.is_empty() {
            continue;
        }

        let msg: RunnerMessage = match serde_json::from_str(trimmed) {
            Ok(msg) => msg,
            Err(err) => {
                // Protocol corruption; abort.
                if let Ok(mut child) = child.lock() {
                    let _ = child.kill();
                }
                return Err(format!(
                    "Python runtime protocol error (invalid JSON line): {err}: {trimmed}"
                ));
            }
        };

        match msg {
            RunnerMessage::Rpc { id, method, params } => {
                let (result, error) = match host.handle_rpc(&method, params) {
                    Ok(value) => (value, None),
                    Err(err) => (JsonValue::Null, Some(err)),
                };
                let response = RpcResponseMessage {
                    msg_type: "rpc_response",
                    id,
                    result,
                    error,
                };
                serde_json::to_writer(&mut stdin, &response).map_err(|e| e.to_string())?;
                stdin.write_all(b"\n").map_err(|e| e.to_string())?;
                stdin.flush().map_err(|e| e.to_string())?;
            }
            RunnerMessage::Result {
                success,
                error,
                traceback,
            } => {
                runner_result = Some((success, error, traceback));
                break;
            }
        }
    }

    // Ensure the process is no longer running and capture stderr.
    let _ = child
        .lock()
        .map_err(|_| "python process mutex poisoned")?
        .wait();
    let _ = stderr_thread.join();

    let stderr_text = stderr_buf
        .lock()
        .map(|s| s.clone())
        .unwrap_or_else(|_| String::new());

    if timed_out.load(Ordering::SeqCst) {
        return Ok(PythonRunResult {
            ok: false,
            stdout: String::new(),
            stderr: stderr_text,
            updates: Vec::new(),
            error: Some(PythonError {
                message: format!("Python script timed out after {timeout_ms}ms"),
                stack: None,
            }),
        });
    }

    let Some((success, error, traceback)) = runner_result else {
        return Err("Python process exited unexpectedly without sending a result".to_string());
    };

    let updates = host
        .take_updates()
        .into_iter()
        .map(cell_update_from_state)
        .collect();

    if success {
        return Ok(PythonRunResult {
            ok: true,
            stdout: String::new(),
            stderr: stderr_text,
            updates,
            error: None,
        });
    }

    Ok(PythonRunResult {
        ok: false,
        stdout: String::new(),
        stderr: stderr_text,
        updates,
        error: Some(PythonError {
            message: error.unwrap_or_else(|| "Python script failed".to_string()),
            stack: traceback,
        }),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ipc_limits::MAX_SCRIPT_CODE_BYTES;
    use serde_json::json;

    #[test]
    fn run_python_script_rejects_oversized_code_without_spawning_python() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let oversized = "x".repeat(MAX_SCRIPT_CODE_BYTES + 1);
        let err = run_python_script(&mut state, &oversized, None, None, None, None)
            .expect_err("expected oversized code to be rejected");
        assert!(
            err.contains("Script is too large"),
            "unexpected error: {err}"
        );
        assert!(
            err.contains(&MAX_SCRIPT_CODE_BYTES.to_string()),
            "expected error to mention limit: {err}"
        );
    }

    #[test]
    fn python_rpc_create_sheet_inserts_after_active_sheet_by_default() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("First".to_string());
        workbook.add_sheet("Second".to_string());
        workbook.add_sheet("Third".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let context = PythonRunContext {
            active_sheet_id: Some("Second".to_string()),
            selection: None,
        };
        let mut host = PythonRpcHost::new(&mut state, Some(context)).expect("host should init");

        let sheet_id = host
            .handle_rpc("create_sheet", Some(json!({ "name": "Inserted" })))
            .expect("create_sheet should succeed");
        assert_eq!(sheet_id, JsonValue::String("Inserted".to_string()));

        let workbook = host.state.get_workbook().expect("workbook should exist");
        let sheet_names: Vec<_> = workbook.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(sheet_names, vec!["First", "Second", "Inserted", "Third"]);
    }

    #[test]
    fn python_rpc_create_sheet_honors_explicit_index_over_active_sheet() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("First".to_string());
        workbook.add_sheet("Second".to_string());
        workbook.add_sheet("Third".to_string());

        let mut state = AppState::new();
        state.load_workbook(workbook);

        let context = PythonRunContext {
            active_sheet_id: Some("Second".to_string()),
            selection: None,
        };
        let mut host = PythonRpcHost::new(&mut state, Some(context)).expect("host should init");

        let sheet_id = host
            .handle_rpc("create_sheet", Some(json!({ "name": "AtStart", "index": 0 })))
            .expect("create_sheet should succeed");
        assert_eq!(sheet_id, JsonValue::String("AtStart".to_string()));

        let workbook = host.state.get_workbook().expect("workbook should exist");
        let sheet_names: Vec<_> = workbook.sheets.iter().map(|s| s.name.as_str()).collect();
        assert_eq!(sheet_names, vec!["AtStart", "First", "Second", "Third"]);
    }

    #[test]
    fn python_rpc_create_sheet_rejects_blank_name_with_canonical_error() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        let err = host
            .handle_rpc("create_sheet", Some(json!({ "name": "   " })))
            .expect_err("expected create_sheet to reject blank name");
        assert_eq!(
            err,
            formula_model::SheetNameError::EmptyName.to_string(),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_rpc_create_sheet_rejects_duplicate_name_with_canonical_error() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        let err = host
            .handle_rpc("create_sheet", Some(json!({ "name": "sheet1" })))
            .expect_err("expected create_sheet to reject duplicate name");
        assert_eq!(
            err,
            formula_model::SheetNameError::DuplicateName.to_string(),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_rpc_create_sheet_rejects_duplicate_name_with_unicode_case_folding_expansion() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("straße".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        let err = host
            .handle_rpc("create_sheet", Some(json!({ "name": "STRASSE" })))
            .expect_err("expected create_sheet to reject duplicate name");
        assert_eq!(
            err,
            formula_model::SheetNameError::DuplicateName.to_string(),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_rpc_create_sheet_rejects_invalid_character_with_canonical_error() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        let err = host
            .handle_rpc("create_sheet", Some(json!({ "name": "Bad/Name" })))
            .expect_err("expected create_sheet to reject invalid name");
        assert_eq!(
            err,
            formula_model::SheetNameError::InvalidCharacter('/').to_string(),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_rpc_create_sheet_rejects_leading_or_trailing_apostrophe_with_canonical_error() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        let err = host
            .handle_rpc("create_sheet", Some(json!({ "name": "'Leading" })))
            .expect_err("expected create_sheet to reject invalid name");
        assert_eq!(
            err,
            formula_model::SheetNameError::LeadingOrTrailingApostrophe.to_string(),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_rpc_create_sheet_rejects_too_long_name_with_canonical_error() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        let long_name = "a".repeat(formula_model::EXCEL_MAX_SHEET_NAME_LEN + 1);
        let err = host
            .handle_rpc("create_sheet", Some(json!({ "name": long_name })))
            .expect_err("expected create_sheet to reject too-long name");
        assert_eq!(
            err,
            formula_model::SheetNameError::TooLong.to_string(),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_rpc_create_sheet_rejects_too_long_name_by_utf16_units_with_canonical_error() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        // 🙂 is 2 UTF-16 code units; 16 of them is 32 units (over Excel's 31-unit limit).
        let long_name = "🙂".repeat(16);
        let err = host
            .handle_rpc("create_sheet", Some(json!({ "name": long_name })))
            .expect_err("expected create_sheet to reject too-long name");
        assert_eq!(
            err,
            formula_model::SheetNameError::TooLong.to_string(),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_rpc_rename_sheet_rejects_blank_name_with_canonical_error() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        let err = host
            .handle_rpc(
                "rename_sheet",
                Some(json!({ "sheet_id": "Sheet1", "name": "   " })),
            )
            .expect_err("expected rename_sheet to reject blank name");
        assert_eq!(
            err,
            formula_model::SheetNameError::EmptyName.to_string(),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_rpc_rename_sheet_rejects_duplicate_name_with_canonical_error() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        workbook.add_sheet("Sheet2".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        let err = host
            .handle_rpc(
                "rename_sheet",
                Some(json!({ "sheet_id": "Sheet2", "name": "sheet1" })),
            )
            .expect_err("expected rename_sheet to reject duplicate name");
        assert_eq!(
            err,
            formula_model::SheetNameError::DuplicateName.to_string(),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_rpc_rename_sheet_rejects_duplicate_name_with_unicode_case_folding_expansion() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("straße".to_string());
        workbook.add_sheet("Sheet2".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        let err = host
            .handle_rpc(
                "rename_sheet",
                Some(json!({ "sheet_id": "Sheet2", "name": "STRASSE" })),
            )
            .expect_err("expected rename_sheet to reject duplicate name");
        assert_eq!(
            err,
            formula_model::SheetNameError::DuplicateName.to_string(),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_rpc_rename_sheet_rejects_invalid_character_with_canonical_error() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        let err = host
            .handle_rpc(
                "rename_sheet",
                Some(json!({ "sheet_id": "Sheet1", "name": "Bad/Name" })),
            )
            .expect_err("expected rename_sheet to reject invalid name");
        assert_eq!(
            err,
            formula_model::SheetNameError::InvalidCharacter('/').to_string(),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_rpc_rename_sheet_rejects_leading_or_trailing_apostrophe_with_canonical_error() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        let err = host
            .handle_rpc(
                "rename_sheet",
                Some(json!({ "sheet_id": "Sheet1", "name": "'Leading" })),
            )
            .expect_err("expected rename_sheet to reject invalid name");
        assert_eq!(
            err,
            formula_model::SheetNameError::LeadingOrTrailingApostrophe.to_string(),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_rpc_rename_sheet_rejects_too_long_name_with_canonical_error() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        let long_name = "a".repeat(formula_model::EXCEL_MAX_SHEET_NAME_LEN + 1);
        let err = host
            .handle_rpc(
                "rename_sheet",
                Some(json!({ "sheet_id": "Sheet1", "name": long_name })),
            )
            .expect_err("expected rename_sheet to reject too-long name");
        assert_eq!(
            err,
            formula_model::SheetNameError::TooLong.to_string(),
            "unexpected error: {err}"
        );
    }

    #[test]
    fn python_rpc_rename_sheet_rejects_too_long_name_by_utf16_units_with_canonical_error() {
        let mut workbook = crate::file_io::Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut host = PythonRpcHost::new(&mut state, None).expect("host should init");
        // 🙂 is 2 UTF-16 code units; 16 of them is 32 units (over Excel's 31-unit limit).
        let long_name = "🙂".repeat(16);
        let err = host
            .handle_rpc(
                "rename_sheet",
                Some(json!({ "sheet_id": "Sheet1", "name": long_name })),
            )
            .expect_err("expected rename_sheet to reject too-long name");
        assert_eq!(
            err,
            formula_model::SheetNameError::TooLong.to_string(),
            "unexpected error: {err}"
        );
    }
}
