use crate::file_io::Workbook;
use crate::resource_limits::{
    macro_output_max_bytes, macro_output_max_lines, MAX_CELL_FORMULA_BYTES,
    MAX_CELL_VALUE_STRING_BYTES, MAX_MACRO_UPDATES,
};
use crate::sheet_name::sheet_name_eq_case_insensitive;
use crate::state::{AppState, CellScalar, CellUpdateData};
use formula_vba_runtime::Spreadsheet;
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};
use std::hash::{Hash, Hasher};
use std::time::Duration;

/// Default execution timeout for `Workbook_BeforeClose` when invoked from the native desktop
/// window-close flow (see `src/main.rs`).
///
/// This does *not* affect user-invoked macros or the `fire_workbook_before_close` IPC command,
/// which accept an explicit `timeout_ms` argument.
pub const WORKBOOK_BEFORE_CLOSE_TIMEOUT_MS_DEFAULT: u64 = 2_000;

const WORKBOOK_BEFORE_CLOSE_TIMEOUT_ENV_VAR: &str = "FORMULA_WORKBOOK_BEFORE_CLOSE_TIMEOUT_MS";

/// Timeout (in milliseconds) applied to the `Workbook_BeforeClose` macro when invoked by the
/// native window close flow.
///
/// This can be overridden by `FORMULA_WORKBOOK_BEFORE_CLOSE_TIMEOUT_MS`. In debug builds the
/// override is honored as-is; in release builds we clamp the value to the default (or lower) to
/// prevent relaxing the guardrails in production.
pub fn workbook_before_close_timeout_ms() -> u64 {
    parse_clamped_timeout_ms(
        std::env::var(WORKBOOK_BEFORE_CLOSE_TIMEOUT_ENV_VAR)
            .ok()
            .as_deref(),
        WORKBOOK_BEFORE_CLOSE_TIMEOUT_MS_DEFAULT,
        cfg!(debug_assertions),
    )
}

fn parse_clamped_timeout_ms(raw: Option<&str>, default: u64, allow_relax: bool) -> u64 {
    let Some(value) = raw.and_then(|v| v.parse::<u64>().ok()) else {
        return default;
    };
    if allow_relax {
        value
    } else {
        value.min(default)
    }
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
    ObjectCreation,
}

impl MacroPermission {
    pub(crate) fn as_runtime_permission(&self) -> formula_vba_runtime::Permission {
        match self {
            MacroPermission::FilesystemRead => formula_vba_runtime::Permission::FileSystemRead,
            MacroPermission::FilesystemWrite => formula_vba_runtime::Permission::FileSystemWrite,
            MacroPermission::Network => formula_vba_runtime::Permission::Network,
            MacroPermission::ObjectCreation => formula_vba_runtime::Permission::ObjectCreation,
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize, PartialEq)]
pub struct MacroPermissionRequest {
    pub reason: String,
    pub macro_id: String,
    pub workbook_origin_path: Option<String>,
    pub requested: Vec<MacroPermission>,
}

#[derive(Clone, Debug, Default)]
pub struct MacroExecutionOptions {
    pub permissions: Vec<MacroPermission>,
    pub timeout_ms: Option<u64>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct MacroExecutionOutcome {
    pub ok: bool,
    pub output: Vec<String>,
    pub updates: Vec<CellUpdateData>,
    pub error: Option<String>,
    pub permission_request: Option<MacroPermissionRequest>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct MacroRuntimeContext {
    pub active_sheet: usize,
    pub active_cell: (u32, u32),
    pub selection: Option<formula_vba_runtime::VbaRangeRef>,
}

impl Default for MacroRuntimeContext {
    fn default() -> Self {
        Self {
            active_sheet: 0,
            active_cell: (1, 1),
            selection: None,
        }
    }
}

#[derive(Debug, thiserror::Error)]
pub enum MacroHostError {
    #[error("no workbook loaded")]
    NoWorkbookLoaded,
    #[error("vba project parse error: {0}")]
    ProjectParse(String),
    #[error("vba program parse error: {0}")]
    ProgramParse(String),
    #[error("macro runtime error: {0}")]
    Runtime(String),
}

#[derive(Debug, Default)]
pub struct MacroHost {
    vba_project_hash: Option<u64>,
    project: Option<formula_vba::VBAProject>,
    combined_source: Option<String>,
    procedure_module: HashMap<String, String>,
    runtime_context: MacroRuntimeContext,
    #[cfg(test)]
    program_compile_count: u64,
}

// SAFETY: `MacroHost` embeds `formula_vba_runtime` types that use `Rc<RefCell<...>>` internally
// (not `Send`). In the desktop shell we only access the macro runtime via the shared `AppState`,
// which is always protected by a `std::sync::Mutex` (see `SharedAppState` in `state.rs`).
//
// This ensures all `Rc` refcount mutations and `RefCell` borrows happen with mutual exclusion,
// even if Tauri invokes commands on different threads. We therefore treat `MacroHost` as
// effectively single-threaded state guarded by a mutex and mark it as `Send` so it can be stored
// in Tauri-managed state.
unsafe impl Send for MacroHost {}

impl MacroHost {
    pub fn invalidate(&mut self) {
        self.vba_project_hash = None;
        self.project = None;
        self.combined_source = None;
        self.procedure_module.clear();
        self.runtime_context = MacroRuntimeContext::default();
        #[cfg(test)]
        {
            self.program_compile_count = 0;
        }
    }

    pub fn runtime_context(&self) -> MacroRuntimeContext {
        self.runtime_context
    }

    pub fn set_runtime_context(&mut self, ctx: MacroRuntimeContext) {
        self.runtime_context = ctx;
    }

    pub(crate) fn sync_with_workbook(&mut self, workbook: &Workbook) {
        self.refresh_if_needed(workbook);
    }

    fn refresh_if_needed(&mut self, workbook: &Workbook) {
        let hash = workbook
            .vba_project_bin
            .as_ref()
            .map(|bytes| hash_bytes(bytes.as_slice()));
        if hash != self.vba_project_hash {
            self.vba_project_hash = hash;
            self.project = None;
            self.combined_source = None;
            self.procedure_module.clear();
            self.runtime_context = MacroRuntimeContext::default();
            #[cfg(test)]
            {
                self.program_compile_count = 0;
            }
        }
    }

    fn ensure_project_loaded(&mut self, workbook: &Workbook) -> Result<(), MacroHostError> {
        self.refresh_if_needed(workbook);
        let Some(vba_bin) = workbook.vba_project_bin.as_ref() else {
            return Ok(());
        };

        if self.project.is_some() {
            return Ok(());
        }

        let project = formula_vba::VBAProject::parse(vba_bin)
            .map_err(|e| MacroHostError::ProjectParse(e.to_string()))?;
        self.procedure_module = build_procedure_module_map(&project);
        self.project = Some(project);
        Ok(())
    }

    fn ensure_sources_loaded(&mut self, workbook: &Workbook) -> Result<(), MacroHostError> {
        self.ensure_project_loaded(workbook)?;
        if workbook.vba_project_bin.is_none() {
            return Ok(());
        }
        if self.combined_source.is_some() {
            return Ok(());
        }

        let project = self
            .project
            .as_ref()
            .ok_or_else(|| MacroHostError::Runtime("missing VBA project".to_string()))?;

        self.combined_source = Some(
            project
                .modules
                .iter()
                .map(|m| m.code.as_str())
                .collect::<Vec<_>>()
                .join("\n\n"),
        );

        Ok(())
    }
    pub fn project(
        &mut self,
        workbook: &Workbook,
    ) -> Result<Option<formula_vba::VBAProject>, MacroHostError> {
        self.ensure_project_loaded(workbook)?;
        Ok(self.project.clone())
    }

    pub fn program(
        &mut self,
        workbook: &Workbook,
    ) -> Result<Option<formula_vba_runtime::VbaProgram>, MacroHostError> {
        self.ensure_sources_loaded(workbook)?;
        let Some(source) = self.combined_source.as_deref() else {
            return Ok(None);
        };
        let program = formula_vba_runtime::parse_program(source)
            .map_err(|e| MacroHostError::ProgramParse(e.to_string()))?;
        #[cfg(test)]
        {
            self.program_compile_count = self.program_compile_count.saturating_add(1);
        }
        Ok(Some(program))
    }

    pub fn list_macros(&mut self, workbook: &Workbook) -> Result<Vec<MacroInfo>, MacroHostError> {
        let Some(program) = self.program(workbook)? else {
            return Ok(Vec::new());
        };
        let module_map = &self.procedure_module;

        let mut macros = program
            .procedures
            .values()
            .map(|proc| {
                let module = module_map.get(&proc.name.to_ascii_lowercase()).cloned();
                MacroInfo {
                    id: proc.name.clone(),
                    name: proc.name.clone(),
                    language: "vba".to_string(),
                    module,
                }
            })
            .collect::<Vec<_>>();
        macros.sort_by(|a, b| a.name.cmp(&b.name));
        Ok(macros)
    }

    #[cfg(test)]
    pub fn program_compile_count(&self) -> u64 {
        self.program_compile_count
    }
}

#[cfg(test)]
mod timeout_tests {
    use super::*;

    #[test]
    fn parse_clamped_timeout_ms_falls_back_to_default_on_missing_or_invalid() {
        assert_eq!(
            parse_clamped_timeout_ms(None, WORKBOOK_BEFORE_CLOSE_TIMEOUT_MS_DEFAULT, true),
            WORKBOOK_BEFORE_CLOSE_TIMEOUT_MS_DEFAULT
        );
        assert_eq!(
            parse_clamped_timeout_ms(
                Some("not-a-number"),
                WORKBOOK_BEFORE_CLOSE_TIMEOUT_MS_DEFAULT,
                true
            ),
            WORKBOOK_BEFORE_CLOSE_TIMEOUT_MS_DEFAULT
        );
    }

    #[test]
    fn parse_clamped_timeout_ms_allows_relax_in_debug_like_mode() {
        assert_eq!(
            parse_clamped_timeout_ms(Some("5000"), WORKBOOK_BEFORE_CLOSE_TIMEOUT_MS_DEFAULT, true),
            5000
        );
    }

    #[test]
    fn parse_clamped_timeout_ms_clamps_in_release_like_mode() {
        assert_eq!(
            parse_clamped_timeout_ms(
                Some("5000"),
                WORKBOOK_BEFORE_CLOSE_TIMEOUT_MS_DEFAULT,
                false
            ),
            WORKBOOK_BEFORE_CLOSE_TIMEOUT_MS_DEFAULT
        );
        assert_eq!(
            parse_clamped_timeout_ms(
                Some("1000"),
                WORKBOOK_BEFORE_CLOSE_TIMEOUT_MS_DEFAULT,
                false
            ),
            1000
        );
    }
}

fn build_procedure_module_map(project: &formula_vba::VBAProject) -> HashMap<String, String> {
    let mut map = HashMap::new();
    for module in &project.modules {
        let Ok(program) = formula_vba_runtime::parse_program(&module.code) else {
            continue;
        };
        for proc in program.procedures.values() {
            map.insert(proc.name.to_ascii_lowercase(), module.name.clone());
        }
    }
    map
}

fn hash_bytes(bytes: &[u8]) -> u64 {
    let mut hasher = std::collections::hash_map::DefaultHasher::new();
    bytes.hash(&mut hasher);
    hasher.finish()
}

#[derive(Clone, Debug)]
pub enum MacroInvocation {
    Procedure {
        macro_id: String,
    },
    WorkbookOpen,
    WorkbookBeforeClose,
    WorksheetChange {
        target: formula_vba_runtime::VbaRangeRef,
    },
    SelectionChange {
        target: formula_vba_runtime::VbaRangeRef,
    },
}

impl MacroInvocation {
    pub fn macro_id(&self) -> String {
        match self {
            MacroInvocation::Procedure { macro_id } => macro_id.clone(),
            MacroInvocation::WorkbookOpen => "Workbook_Open".to_string(),
            MacroInvocation::WorkbookBeforeClose => "Workbook_BeforeClose".to_string(),
            MacroInvocation::WorksheetChange { .. } => "Worksheet_Change".to_string(),
            MacroInvocation::SelectionChange { .. } => "Worksheet_SelectionChange".to_string(),
        }
    }

    fn kind(&self) -> &'static str {
        match self {
            MacroInvocation::Procedure { .. } => "run_macro",
            MacroInvocation::WorkbookOpen => "workbook_open",
            MacroInvocation::WorkbookBeforeClose => "workbook_before_close",
            MacroInvocation::WorksheetChange { .. } => "worksheet_change",
            MacroInvocation::SelectionChange { .. } => "selection_change",
        }
    }
}

#[derive(Debug, Serialize)]
struct MacroAuditEvent {
    event: String,
    kind: String,
    macro_id: String,
    workbook_origin_path: Option<String>,
    permissions: Vec<MacroPermission>,
    ok: Option<bool>,
    error: Option<String>,
    permission_request: Option<MacroPermissionRequest>,
}

fn emit_audit(event: MacroAuditEvent) {
    match serde_json::to_string(&event) {
        Ok(json) => crate::stdio::stderrln(format_args!("[macro_audit] {json}")),
        Err(err) => crate::stdio::stderrln(format_args!(
            "[macro_audit] failed to serialize audit event: {err}"
        )),
    }
}

struct MacroPermissionChecker {
    allowed: HashSet<formula_vba_runtime::Permission>,
}

impl formula_vba_runtime::PermissionChecker for MacroPermissionChecker {
    fn has_permission(&self, permission: formula_vba_runtime::Permission) -> bool {
        self.allowed.contains(&permission)
    }
}

pub fn execute_invocation(
    state: &mut AppState,
    program: formula_vba_runtime::VbaProgram,
    ctx: MacroRuntimeContext,
    workbook_origin_path: Option<String>,
    invocation: MacroInvocation,
    options: MacroExecutionOptions,
) -> Result<(MacroExecutionOutcome, MacroRuntimeContext), MacroHostError> {
    let mut policy = formula_vba_runtime::VbaSandboxPolicy::default();
    // `formula_vba_runtime::VbaSandboxPolicy` defaults to a very small execution time budget
    // (currently 250ms). That is helpful as a conservative sandbox default, but typical workbook
    // macros can legitimately take longer (especially in debug builds / under CI load).
    //
    // Prefer a modest, host-controlled default and allow callers to override via IPC.
    policy.max_execution_time = Duration::from_secs(5);
    if let Some(timeout_ms) = options.timeout_ms {
        policy.max_execution_time = Duration::from_millis(timeout_ms);
    }

    let mut allowed = HashSet::new();
    for perm in &options.permissions {
        match perm {
            MacroPermission::FilesystemRead => policy.allow_filesystem_read = true,
            MacroPermission::FilesystemWrite => policy.allow_filesystem_write = true,
            MacroPermission::Network => policy.allow_network = true,
            MacroPermission::ObjectCreation => policy.allow_object_creation = true,
        }
        allowed.insert(perm.as_runtime_permission());
    }

    let checker = MacroPermissionChecker { allowed };
    let runtime = formula_vba_runtime::VbaRuntime::new(program)
        .with_sandbox_policy(policy)
        .with_permission_checker(Box::new(checker));

    let macro_id = invocation.macro_id();
    emit_audit(MacroAuditEvent {
        event: "start".to_string(),
        kind: invocation.kind().to_string(),
        macro_id: macro_id.clone(),
        workbook_origin_path: workbook_origin_path.clone(),
        permissions: options.permissions.clone(),
        ok: None,
        error: None,
        permission_request: None,
    });

    let mut sheet = AppStateSpreadsheet::new(state, ctx)
        .map_err(|err| MacroHostError::Runtime(err.to_string()))?;

    let initial_selection = ctx
        .selection
        .filter(|sel| sel.sheet < sheet.sheet_count());
    let exec: Result<formula_vba_runtime::ExecutionResult, formula_vba_runtime::VbaError> =
        match &invocation {
        MacroInvocation::Procedure { macro_id } => {
            runtime.execute_with_selection(&mut sheet, macro_id, &[], initial_selection)
        }
        MacroInvocation::WorkbookOpen => {
            runtime.fire_workbook_open_with_selection(&mut sheet, initial_selection)
        }
        MacroInvocation::WorkbookBeforeClose => runtime.fire_workbook_before_close_with_selection(
            &mut sheet,
            initial_selection,
        ),
        MacroInvocation::WorksheetChange { target } => {
            runtime.fire_worksheet_change_with_selection(&mut sheet, *target, initial_selection)
        }
        MacroInvocation::SelectionChange { target } => {
            runtime.fire_worksheet_selection_change_with_selection(
                &mut sheet,
                *target,
                initial_selection,
            )
        }
    };

    let output = sheet.take_output();
    let updates = dedup_updates(sheet.take_updates());
    let selection = match &exec {
        Ok(res) => res.selection,
        Err(_) => initial_selection,
    };
    let selection = selection.filter(|sel| sel.sheet < sheet.sheet_count());
    let new_ctx = MacroRuntimeContext {
        active_sheet: sheet.active_sheet(),
        active_cell: sheet.active_cell(),
        selection,
    };

    let (ok, error, permission_request) = match exec {
        Ok(_) => (true, None, None),
        Err(err) => {
            let message = err.to_string();
            let permission_request = match &err {
                formula_vba_runtime::VbaError::Sandbox(reason) => {
                    permission_request_from_sandbox(reason, &macro_id, workbook_origin_path.clone())
                }
                _ => None,
            };
            (false, Some(message), permission_request)
        }
    };

    let outcome = MacroExecutionOutcome {
        ok,
        output,
        updates,
        error,
        permission_request: permission_request.clone(),
    };

    emit_audit(MacroAuditEvent {
        event: "end".to_string(),
        kind: invocation.kind().to_string(),
        macro_id,
        workbook_origin_path,
        permissions: options.permissions,
        ok: Some(outcome.ok),
        error: outcome.error.clone(),
        permission_request,
    });

    Ok((outcome, new_ctx))
}

fn permission_request_from_sandbox(
    reason: &str,
    macro_id: &str,
    workbook_origin_path: Option<String>,
) -> Option<MacroPermissionRequest> {
    let requested = parse_permission_from_sandbox_reason(reason)?;
    Some(MacroPermissionRequest {
        reason: reason.to_string(),
        macro_id: macro_id.to_string(),
        workbook_origin_path,
        requested: vec![requested],
    })
}

fn parse_permission_from_sandbox_reason(reason: &str) -> Option<MacroPermission> {
    let marker = "permission:";
    let idx = reason.to_ascii_lowercase().find(marker)?;
    let after = reason[idx + marker.len()..].trim();
    let token = after.split_whitespace().next().unwrap_or(after);
    match token.trim_matches(|c: char| c == ',' || c == '.' || c == ';') {
        "FileSystemRead" => Some(MacroPermission::FilesystemRead),
        "FileSystemWrite" => Some(MacroPermission::FilesystemWrite),
        "Network" => Some(MacroPermission::Network),
        "ObjectCreation" => Some(MacroPermission::ObjectCreation),
        _ => None,
    }
}

fn dedup_updates(updates: Vec<CellUpdateData>) -> Vec<CellUpdateData> {
    let mut out = Vec::new();
    let mut idx = HashMap::new();
    for update in updates {
        let key = (update.sheet_id.clone(), update.row, update.col);
        if let Some(existing) = idx.get(&key).copied() {
            out[existing] = update;
        } else {
            idx.insert(key, out.len());
            out.push(update);
        }
    }
    out
}

const VBA_MACRO_OUTPUT_TRUNCATED_SENTINEL: &str = "[truncated]";

#[derive(Debug)]
struct MacroOutputBuffer {
    max_lines: usize,
    max_bytes: usize,
    used_bytes: usize,
    truncated: bool,
    lines: Vec<String>,
}

impl MacroOutputBuffer {
    fn new(max_lines: usize, max_bytes: usize) -> Self {
        Self {
            max_lines,
            max_bytes,
            used_bytes: 0,
            truncated: false,
            lines: Vec::new(),
        }
    }

    fn take(&mut self) -> Vec<String> {
        self.used_bytes = 0;
        self.truncated = false;
        std::mem::take(&mut self.lines)
    }

    fn push(&mut self, message: String) {
        if self.truncated {
            return;
        }

        // Fast-path: accept the message if it fits.
        if self.lines.len() < self.max_lines && self.used_bytes.saturating_add(message.len()) <= self.max_bytes {
            self.used_bytes = self.used_bytes.saturating_add(message.len());
            self.lines.push(message);
            return;
        }

        // Otherwise, deterministically stop collecting output and append a truncation sentinel.
        self.truncated = true;
        self.ensure_truncation_sentinel();
    }

    fn ensure_truncation_sentinel(&mut self) {
        if self.max_lines == 0 || self.max_bytes == 0 {
            self.lines.clear();
            self.used_bytes = 0;
            return;
        }

        // Ensure the sentinel itself never exceeds the byte budget.
        let mut sentinel = VBA_MACRO_OUTPUT_TRUNCATED_SENTINEL.to_string();
        if sentinel.len() > self.max_bytes {
            sentinel = truncate_utf8_to_bytes(&sentinel, self.max_bytes);
        }
        let sentinel_len = sentinel.len();

        // Clamp to the maximum line count before inserting the sentinel.
        while self.lines.len() > self.max_lines {
            if let Some(removed) = self.lines.pop() {
                self.used_bytes = self.used_bytes.saturating_sub(removed.len());
            }
        }

        // We insert the sentinel as the last line. If we are already at the line cap we replace
        // the last line, otherwise we append a new one.
        let replace_last = self.lines.len() >= self.max_lines && !self.lines.is_empty();
        if replace_last {
            if let Some(removed) = self.lines.pop() {
                self.used_bytes = self.used_bytes.saturating_sub(removed.len());
            }
        }

        // Drop existing lines from the end until the sentinel fits within the byte budget.
        while !self.lines.is_empty()
            && self.used_bytes.saturating_add(sentinel_len) > self.max_bytes
        {
            if let Some(removed) = self.lines.pop() {
                self.used_bytes = self.used_bytes.saturating_sub(removed.len());
            }
        }

        // If the sentinel still doesn't fit, clear everything (this can only happen for extremely
        // small `max_bytes` values, e.g. via debug env overrides).
        if self.used_bytes.saturating_add(sentinel_len) > self.max_bytes {
            self.lines.clear();
            self.used_bytes = 0;
        }

        // At this point, adding the sentinel is guaranteed to be within both budgets.
        if self.lines.len() < self.max_lines {
            self.lines.push(sentinel);
            self.used_bytes = self.used_bytes.saturating_add(sentinel_len);
        }
    }
}

fn truncate_utf8_to_bytes(input: &str, max_bytes: usize) -> String {
    if input.len() <= max_bytes {
        return input.to_string();
    }

    let mut end = max_bytes.min(input.len());
    while end > 0 && !input.is_char_boundary(end) {
        end -= 1;
    }
    input[..end].to_string()
}

struct AppStateSpreadsheet<'a> {
    state: &'a mut AppState,
    active_sheet: usize,
    active_cell: (u32, u32),
    output: MacroOutputBuffer,
    updates: Vec<CellUpdateData>,
    update_index_by_cell: HashMap<(String, usize, usize), usize>,
    undo_entries_added: usize,
}

impl<'a> AppStateSpreadsheet<'a> {
    fn new(
        state: &'a mut AppState,
        ctx: MacroRuntimeContext,
    ) -> Result<Self, formula_vba_runtime::VbaError> {
        let workbook = state
            .get_workbook()
            .map_err(|e| formula_vba_runtime::VbaError::Runtime(e.to_string()))?;
        let active_sheet = if workbook.sheets.is_empty() {
            0
        } else {
            ctx.active_sheet.min(workbook.sheets.len() - 1)
        };
        Ok(Self {
            state,
            active_sheet,
            active_cell: ctx.active_cell,
            output: MacroOutputBuffer::new(macro_output_max_lines(), macro_output_max_bytes()),
            updates: Vec::new(),
            update_index_by_cell: HashMap::new(),
            undo_entries_added: 0,
        })
    }

    fn active_sheet(&self) -> usize {
        self.active_sheet
    }

    fn active_cell(&self) -> (u32, u32) {
        self.active_cell
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

    fn cell_scalar_to_vba(value: CellScalar) -> formula_vba_runtime::VbaValue {
        match value {
            CellScalar::Empty => formula_vba_runtime::VbaValue::Empty,
            CellScalar::Number(n) => formula_vba_runtime::VbaValue::Double(n),
            CellScalar::Text(s) => formula_vba_runtime::VbaValue::String(s),
            CellScalar::Bool(b) => formula_vba_runtime::VbaValue::Boolean(b),
            CellScalar::Error(e) => formula_vba_runtime::VbaValue::String(e),
        }
    }

    fn vba_value_to_cell_edit(
        value: &formula_vba_runtime::VbaValue,
    ) -> Result<(Option<serde_json::Value>, Option<String>), formula_vba_runtime::VbaError> {
        match value {
            formula_vba_runtime::VbaValue::Empty | formula_vba_runtime::VbaValue::Null => {
                Ok((None, None))
            }
            formula_vba_runtime::VbaValue::Boolean(b) => {
                Ok((Some(serde_json::Value::from(*b)), None))
            }
            formula_vba_runtime::VbaValue::Double(n) => {
                Ok((Some(serde_json::Value::from(*n)), None))
            }
            formula_vba_runtime::VbaValue::String(s) => {
                if s.starts_with('=') {
                    if s.len() > MAX_CELL_FORMULA_BYTES {
                        return Err(formula_vba_runtime::VbaError::Runtime(format!(
                            "cell formula is too large (max {MAX_CELL_FORMULA_BYTES} bytes)"
                        )));
                    }
                    Ok((None, Some(s.clone())))
                } else {
                    if s.len() > MAX_CELL_VALUE_STRING_BYTES {
                        return Err(formula_vba_runtime::VbaError::Runtime(format!(
                            "cell value string is too large (max {MAX_CELL_VALUE_STRING_BYTES} bytes)"
                        )));
                    }
                    Ok((Some(serde_json::Value::from(s.clone())), None))
                }
            }
            other => Err(formula_vba_runtime::VbaError::Runtime(format!(
                "Unsupported VBA value for cell assignment: {other:?}"
            ))),
        }
    }

    fn take_output(&mut self) -> Vec<String> {
        self.output.take()
    }

    fn take_updates(&mut self) -> Vec<CellUpdateData> {
        self.update_index_by_cell.clear();
        std::mem::take(&mut self.updates)
    }

    fn push_updates(
        &mut self,
        updates: Vec<CellUpdateData>,
    ) -> Result<(), formula_vba_runtime::VbaError> {
        if updates.is_empty() {
            return Ok(());
        }

        // `state.set_cell` pushed an undo entry for this edit (the no-op early return path yields
        // an empty updates vector, which we handled above).
        self.undo_entries_added = self.undo_entries_added.saturating_add(1);

        // Count the number of *new* distinct cells this batch would add. Existing cells are updated
        // in-place and do not increase the total update cardinality.
        let mut new_keys = HashSet::<(String, usize, usize)>::new();
        for update in &updates {
            let key = (update.sheet_id.clone(), update.row, update.col);
            if self.update_index_by_cell.contains_key(&key) {
                continue;
            }
            new_keys.insert(key);
        }

        let remaining = MAX_MACRO_UPDATES.saturating_sub(self.updates.len());
        if new_keys.len() > remaining {
            return Err(self.rollback_macro_edits_due_to_update_limit());
        }

        for update in updates {
            let key = (update.sheet_id.clone(), update.row, update.col);
            if let Some(idx) = self.update_index_by_cell.get(&key).copied() {
                self.updates[idx] = update;
            } else {
                self.update_index_by_cell.insert(key, self.updates.len());
                self.updates.push(update);
            }
        }
        Ok(())
    }

    fn rollback_macro_edits_due_to_update_limit(&mut self) -> formula_vba_runtime::VbaError {
        // The workbook has already been mutated by the macro. Roll back *all* edits applied during
        // this macro invocation so we don't return a partial update payload or leave the
        // frontend/backend out of sync.
        for _ in 0..self.undo_entries_added {
            let _ = self.state.undo();
        }
        self.undo_entries_added = 0;
        self.updates.clear();
        self.update_index_by_cell.clear();

        // Clear redo history so the rolled-back macro changes can't be "redone" later.
        self.state.mark_dirty();
        formula_vba_runtime::VbaError::Runtime(format!(
            "macro produced too many cell updates (limit {MAX_MACRO_UPDATES})"
        ))
    }
}

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
                .position(|s| sheet_name_eq_case_insensitive(&s.name, name))
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

        let row0 = (row - 1) as usize;
        let col0 = (col - 1) as usize;
        let key = (sheet_id.clone(), row0, col0);
        if !self.update_index_by_cell.contains_key(&key) && self.updates.len() >= MAX_MACRO_UPDATES {
            return Err(self.rollback_macro_edits_due_to_update_limit());
        }

        let (value_json, formula) = Self::vba_value_to_cell_edit(&value)?;
        let updates = self
            .state
            .set_cell(
                &sheet_id,
                row0,
                col0,
                value_json,
                formula,
            )
            .map_err(|e| formula_vba_runtime::VbaError::Runtime(e.to_string()))?;
        self.push_updates(updates)?;
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

        let row0 = (row - 1) as usize;
        let col0 = (col - 1) as usize;
        let key = (sheet_id.clone(), row0, col0);
        if !self.update_index_by_cell.contains_key(&key) && self.updates.len() >= MAX_MACRO_UPDATES {
            return Err(self.rollback_macro_edits_due_to_update_limit());
        }

        if formula.len() > MAX_CELL_FORMULA_BYTES {
            return Err(formula_vba_runtime::VbaError::Runtime(format!(
                "cell formula is too large (max {MAX_CELL_FORMULA_BYTES} bytes)"
            )));
        }

        let updates = self
            .state
            .set_cell(
                &sheet_id,
                row0,
                col0,
                None,
                Some(formula),
            )
            .map_err(|e| formula_vba_runtime::VbaError::Runtime(e.to_string()))?;
        self.push_updates(updates)?;
        Ok(())
    }

    fn clear_cell_contents(
        &mut self,
        sheet: usize,
        row: u32,
        col: u32,
    ) -> Result<(), formula_vba_runtime::VbaError> {
        let sheet_id = self.sheet_id(sheet)?;
        if row == 0 || col == 0 {
            return Err(formula_vba_runtime::VbaError::Runtime(
                "Row/col are 1-based".to_string(),
            ));
        }

        let row0 = (row - 1) as usize;
        let col0 = (col - 1) as usize;
        let key = (sheet_id.clone(), row0, col0);
        if !self.update_index_by_cell.contains_key(&key) && self.updates.len() >= MAX_MACRO_UPDATES {
            return Err(self.rollback_macro_edits_due_to_update_limit());
        }

        let updates = self
            .state
            .set_cell(
                &sheet_id,
                row0,
                col0,
                None,
                None,
            )
            .map_err(|e| formula_vba_runtime::VbaError::Runtime(e.to_string()))?;
        self.push_updates(updates)?;
        Ok(())
    }

    fn log(&mut self, message: String) {
        const MESSAGE_TRUNCATED_SUFFIX: &str = "...[truncated]";
        const MAX_LINE_BYTES_HARD: usize = 8 * 1024;

        // Limit a single log message to a bounded size, and ensure we don't retain a huge `String`
        // capacity even if the content is short.
        let max_line_bytes = MAX_LINE_BYTES_HARD.min(self.output.max_bytes);
        let mut message = message;
        if message.len() > max_line_bytes {
            let suffix_len = MESSAGE_TRUNCATED_SUFFIX.len();
            let prefix_budget = max_line_bytes.saturating_sub(suffix_len);
            let mut end = prefix_budget.min(message.len());
            while end > 0 && !message.is_char_boundary(end) {
                end -= 1;
            }
            // Build the truncated string with a bounded capacity so appending the suffix doesn't
            // trigger `String` growth (typically doubling capacity), which would cause us to
            // retain allocations larger than `max_line_bytes`.
            let mut truncated = String::new();
            let _ = truncated.try_reserve(max_line_bytes);
            truncated.push_str(&message[..end]);
            if truncated.len().saturating_add(suffix_len) <= max_line_bytes {
                truncated.push_str(MESSAGE_TRUNCATED_SUFFIX);
            }
            message = truncated;
        }
        if message.capacity() > max_line_bytes {
            // Even if the string is short enough (or was truncated), a malicious macro can still
            // force a huge allocation and pass it through to us. Ensure we don't retain that
            // capacity in the output buffer.
            message = message.as_str().to_string();
        }

        self.output.push(message);
    }

    fn last_used_row_in_column(&self, sheet: usize, col: u32, start_row: u32) -> Option<u32> {
        if col == 0 || start_row == 0 {
            return None;
        }
        let workbook = self.state.get_workbook().ok()?;
        let sheet = workbook.sheets.get(sheet)?;
        let col0 = (col - 1) as usize;

        let mut best: Option<u32> = None;

        if let Some(table) = &sheet.columnar {
            if col0 < table.column_count() && table.row_count() > 0 {
                let candidate = (table.row_count() as u32).min(start_row);
                if candidate > 0 {
                    best = Some(candidate);
                }
            }
        }

        for (&(row0, col_idx), cell) in sheet.cells.iter() {
            if col_idx != col0 {
                continue;
            }
            let row1 = (row0 + 1) as u32;
            if row1 > start_row {
                continue;
            }
            let has_content = cell.formula.is_some()
                || cell.input_value.is_some()
                || !matches!(cell.computed_value, CellScalar::Empty);
            if !has_content {
                continue;
            }
            best = Some(best.map(|b| b.max(row1)).unwrap_or(row1));
        }

        best
    }

    fn next_used_row_in_column(&self, sheet: usize, col: u32, start_row: u32) -> Option<u32> {
        if col == 0 || start_row == 0 {
            return None;
        }
        let workbook = self.state.get_workbook().ok()?;
        let sheet = workbook.sheets.get(sheet)?;
        let col0 = (col - 1) as usize;

        let mut best: Option<u32> = None;

        if let Some(table) = &sheet.columnar {
            if col0 < table.column_count() && table.row_count() > 0 {
                let row_count = table.row_count() as u32;
                if start_row <= row_count {
                    best = Some(start_row);
                }
            }
        }

        for (&(row0, col_idx), cell) in sheet.cells.iter() {
            if col_idx != col0 {
                continue;
            }
            let row1 = (row0 + 1) as u32;
            if row1 < start_row {
                continue;
            }
            let has_content = cell.formula.is_some()
                || cell.input_value.is_some()
                || !matches!(cell.computed_value, CellScalar::Empty);
            if !has_content {
                continue;
            }
            best = Some(best.map(|b| b.min(row1)).unwrap_or(row1));
        }

        best
    }

    fn last_used_col_in_row(&self, sheet: usize, row: u32, start_col: u32) -> Option<u32> {
        if row == 0 || start_col == 0 {
            return None;
        }
        let workbook = self.state.get_workbook().ok()?;
        let sheet = workbook.sheets.get(sheet)?;
        let row0 = (row - 1) as usize;

        let mut best: Option<u32> = None;

        if let Some(table) = &sheet.columnar {
            if row0 < table.row_count() && table.column_count() > 0 {
                let candidate = (table.column_count() as u32).min(start_col);
                if candidate > 0 {
                    best = Some(candidate);
                }
            }
        }

        for (&(row_idx, col0), cell) in sheet.cells.iter() {
            if row_idx != row0 {
                continue;
            }
            let col1 = (col0 + 1) as u32;
            if col1 > start_col {
                continue;
            }
            let has_content = cell.formula.is_some()
                || cell.input_value.is_some()
                || !matches!(cell.computed_value, CellScalar::Empty);
            if !has_content {
                continue;
            }
            best = Some(best.map(|b| b.max(col1)).unwrap_or(col1));
        }

        best
    }

    fn next_used_col_in_row(&self, sheet: usize, row: u32, start_col: u32) -> Option<u32> {
        if row == 0 || start_col == 0 {
            return None;
        }
        let workbook = self.state.get_workbook().ok()?;
        let sheet = workbook.sheets.get(sheet)?;
        let row0 = (row - 1) as usize;

        let mut best: Option<u32> = None;

        if let Some(table) = &sheet.columnar {
            if row0 < table.row_count() && table.column_count() > 0 {
                let col_count = table.column_count() as u32;
                if start_col <= col_count {
                    best = Some(start_col);
                }
            }
        }

        for (&(row_idx, col0), cell) in sheet.cells.iter() {
            if row_idx != row0 {
                continue;
            }
            let col1 = (col0 + 1) as u32;
            if col1 < start_col {
                continue;
            }
            let has_content = cell.formula.is_some()
                || cell.input_value.is_some()
                || !matches!(cell.computed_value, CellScalar::Empty);
            if !has_content {
                continue;
            }
            best = Some(best.map(|b| b.min(col1)).unwrap_or(col1));
        }

        best
    }

    fn used_cells_in_range(
        &self,
        range: formula_vba_runtime::VbaRangeRef,
    ) -> Option<Vec<(u32, u32)>> {
        let workbook = self.state.get_workbook().ok()?;
        let sheet = workbook.sheets.get(range.sheet)?;

        let mut out = Vec::new();
        for (&(row0, col0), cell) in sheet.cells.iter() {
            let row1 = (row0 + 1) as u32;
            let col1 = (col0 + 1) as u32;
            if row1 < range.start_row
                || row1 > range.end_row
                || col1 < range.start_col
                || col1 > range.end_col
            {
                continue;
            }
            let has_content = cell.formula.is_some()
                || cell.input_value.is_some()
                || !matches!(cell.computed_value, CellScalar::Empty);
            if !has_content {
                continue;
            }
            out.push((row1, col1));
        }

        Some(out)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::resource_limits::{
        MAX_CELL_FORMULA_BYTES, MAX_CELL_VALUE_STRING_BYTES, MAX_MACRO_OUTPUT_BYTES,
        MAX_MACRO_OUTPUT_LINES, MAX_MACRO_UPDATES,
    };
    use crate::state::Cell;

    fn empty_state_with_sheet() -> AppState {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);
        state
    }

    #[test]
    fn macro_output_is_capped_by_lines_and_bytes() {
        let mut state = empty_state_with_sheet();

        // Use a large literal string so the macro quickly exceeds the byte limit.
        let payload = "x".repeat(16 * 1024);
        let source = format!(
            r#"
Sub SpamOutput()
    Dim i As Integer
    For i = 1 To 500
        Debug.Print "{payload}"
    Next i
End Sub
"#
        );

        let program = formula_vba_runtime::parse_program(&source).expect("parse program");
        let (outcome, _ctx) = execute_invocation(
            &mut state,
            program,
            MacroRuntimeContext::default(),
            None,
            MacroInvocation::Procedure {
                macro_id: "SpamOutput".to_string(),
            },
            MacroExecutionOptions::default(),
        )
        .expect("execute macro");

        assert!(outcome.ok, "expected macro to succeed: {outcome:?}");

        let total_bytes: usize = outcome.output.iter().map(|s| s.len()).sum();
        assert!(
            outcome.output.len() <= MAX_MACRO_OUTPUT_LINES,
            "expected output lines <= {MAX_MACRO_OUTPUT_LINES}, got {}",
            outcome.output.len()
        );
        assert!(
            total_bytes <= MAX_MACRO_OUTPUT_BYTES,
            "expected output bytes <= {MAX_MACRO_OUTPUT_BYTES}, got {total_bytes}"
        );

        assert!(
            outcome
                .output
                .last()
                .is_some_and(|s| s == VBA_MACRO_OUTPUT_TRUNCATED_SENTINEL),
            "expected a single truncation marker at the end, got: {:?}",
            outcome.output.last()
        );

        // Ensure we never retain a giant `String` allocation in the output buffer.
        let max_line_bytes = (8 * 1024).min(MAX_MACRO_OUTPUT_BYTES);
        for line in &outcome.output {
            assert!(
                line.capacity() <= max_line_bytes,
                "expected output line capacity <= {max_line_bytes}, got {}",
                line.capacity()
            );
        }
    }

    #[test]
    fn macro_log_does_not_retain_giant_string_capacity() {
        let mut state = empty_state_with_sheet();
        let mut sheet =
            AppStateSpreadsheet::new(&mut state, MacroRuntimeContext::default()).expect("new sheet");

        // Simulate a malicious `Debug.Print` string that was allocated with an extreme capacity.
        let mut message = String::new();
        let _ = message.try_reserve(10 * 1024 * 1024);
        message.push_str("ok");
        sheet.log(message);

        let out = sheet.take_output();
        assert_eq!(out.len(), 1);
        let max_line_bytes = (8 * 1024).min(MAX_MACRO_OUTPUT_BYTES);
        assert!(
            out[0].capacity() <= max_line_bytes,
            "expected log output capacity <= {max_line_bytes}, got {}",
            out[0].capacity()
        );
    }

    #[test]
    fn macro_aborts_when_cell_value_string_is_too_large() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        let info = state.load_workbook(workbook);
        let sheet_id = info.sheets[0].id.clone();

        let payload = "x".repeat(MAX_CELL_VALUE_STRING_BYTES + 1);
        let source = format!(
            r#"
Sub BigValue()
    Range("A1").Value = "{payload}"
End Sub
"#
        );

        let program = formula_vba_runtime::parse_program(&source).expect("parse program");
        let (outcome, _ctx) = execute_invocation(
            &mut state,
            program,
            MacroRuntimeContext::default(),
            None,
            MacroInvocation::Procedure {
                macro_id: "BigValue".to_string(),
            },
            MacroExecutionOptions::default(),
        )
        .expect("execute macro");

        assert!(!outcome.ok, "expected macro to fail due to value size limit");
        let err = outcome.error.expect("expected error message");
        assert!(
            err.contains(&MAX_CELL_VALUE_STRING_BYTES.to_string()),
            "expected error to mention limit {MAX_CELL_VALUE_STRING_BYTES}, got: {err}"
        );
        assert_eq!(
            state.get_cell(&sheet_id, 0, 0).unwrap().value,
            CellScalar::Empty
        );
    }

    #[test]
    fn set_cell_formula_rejects_oversized_formulas() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());
        let mut state = AppState::new();
        state.load_workbook(workbook);

        let mut sheet =
            AppStateSpreadsheet::new(&mut state, MacroRuntimeContext::default()).expect("new sheet");

        let oversized = format!("={}", "1".repeat(MAX_CELL_FORMULA_BYTES));
        let err = sheet
            .set_cell_formula(0, 1, 1, oversized)
            .expect_err("expected oversized formula to be rejected");
        assert!(
            err.to_string().contains(&MAX_CELL_FORMULA_BYTES.to_string()),
            "expected error to mention limit {MAX_CELL_FORMULA_BYTES}, got: {err}"
        );
    }

    #[test]
    fn macro_aborts_when_update_limit_is_exceeded() {
        let mut workbook = Workbook::new_empty(None);
        workbook.add_sheet("Sheet1".to_string());

        // Seed many formulas that depend on A1 so a single write triggers a large fanout.
        {
            let sheet = workbook.sheets.get_mut(0).expect("Sheet1");
            sheet.cells.reserve(MAX_MACRO_UPDATES + 1);
            for row in 0..=MAX_MACRO_UPDATES {
                sheet
                    .cells
                    .insert((row, 1), Cell::from_formula("=A1".to_string()));
            }
        }

        let mut state = AppState::new();
        let info = state.load_workbook(workbook);
        let sheet_id = info.sheets[0].id.clone();

        let source = r#"
Sub TouchA1()
    Range("A1").Value = 1
End Sub
"#;
        let program = formula_vba_runtime::parse_program(source).expect("parse program");
        let (outcome, _ctx) = execute_invocation(
            &mut state,
            program,
            MacroRuntimeContext::default(),
            None,
            MacroInvocation::Procedure {
                macro_id: "TouchA1".to_string(),
            },
            MacroExecutionOptions::default(),
        )
        .expect("execute macro");

        assert!(!outcome.ok, "expected macro to fail due to update limit");
        let err = outcome.error.expect("expected error message");
        assert!(
            err.contains(&format!("limit {MAX_MACRO_UPDATES}")),
            "expected error to mention the limit {MAX_MACRO_UPDATES}, got: {err}"
        );

        // The last write should be rolled back so the backend workbook remains consistent with the
        // (empty) update payload returned for this failed invocation.
        assert_eq!(
            state.get_cell(&sheet_id, 0, 0).unwrap().value,
            CellScalar::Empty
        );
    }

    #[test]
    fn macro_output_truncates_by_line_count() {
        let mut out = MacroOutputBuffer::new(3, 1024);
        out.push("a".to_string());
        out.push("b".to_string());
        out.push("c".to_string());
        out.push("d".to_string());

        let lines = out.take();
        assert_eq!(
            lines,
            vec![
                "a".to_string(),
                "b".to_string(),
                VBA_MACRO_OUTPUT_TRUNCATED_SENTINEL.to_string()
            ]
        );
    }

    #[test]
    fn macro_output_truncates_by_total_byte_budget() {
        let msg = "123456789012345".to_string(); // 15 bytes
        let mut out = MacroOutputBuffer::new(10, 40);
        out.push(msg.clone());
        out.push(msg.clone());
        out.push(msg.clone());

        let lines = out.take();
        assert_eq!(
            lines,
            vec![msg.clone(), VBA_MACRO_OUTPUT_TRUNCATED_SENTINEL.to_string()]
        );
        let total: usize = lines.iter().map(|s| s.len()).sum();
        assert!(total <= 40);
    }

    #[test]
    fn truncation_sentinel_appears_at_most_once() {
        let mut out = MacroOutputBuffer::new(2, 1024);
        for i in 0..100 {
            out.push(format!("line {i}"));
        }

        let lines = out.take();
        let sentinel_count = lines
            .iter()
            .filter(|line| line.as_str() == VBA_MACRO_OUTPUT_TRUNCATED_SENTINEL)
            .count();
        assert_eq!(sentinel_count, 1);
        assert_eq!(
            lines.last().map(|s| s.as_str()),
            Some(VBA_MACRO_OUTPUT_TRUNCATED_SENTINEL)
        );
    }

    #[test]
    fn macro_output_handles_empty_and_multibyte_messages_without_panicking() {
        let mut out = MacroOutputBuffer::new(3, 25);
        out.push(String::new());
        out.push("💣".to_string()); // multi-byte UTF-8
        out.push("a".repeat(100));

        let lines = out.take();
        let total: usize = lines.iter().map(|s| s.len()).sum();
        assert!(total <= 25);
        assert!(lines.len() <= 3);
        assert_eq!(
            lines.last().map(|s| s.as_str()),
            Some(VBA_MACRO_OUTPUT_TRUNCATED_SENTINEL)
        );
    }
}
