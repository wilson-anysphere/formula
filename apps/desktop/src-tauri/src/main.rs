#![cfg_attr(
    all(not(debug_assertions), target_os = "windows"),
    windows_subsystem = "windows"
)]

mod asset_protocol;
mod menu;
mod shortcuts;
mod tray;

use desktop::clipboard;
use desktop::commands;
use desktop::ed25519_verifier;
use desktop::macro_trust::{compute_macro_fingerprint, MacroTrustStore, SharedMacroTrustStore};
use desktop::macros::MacroExecutionOptions;
use desktop::open_file;
use desktop::open_file_ipc::OpenFileState;
use desktop::oauth_redirect_ipc::OauthRedirectState;
#[cfg(target_os = "macos")]
use desktop::opened_urls;
use desktop::process_metrics;
use desktop::state::{AppState, CellUpdateData, SharedAppState};
use desktop::tray_status::{self, TrayStatusState};
use desktop::updater;
use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use std::net::{Ipv4Addr, Ipv6Addr, SocketAddr};
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::Instant;
use tauri::http::header::{HeaderName, HeaderValue};
use tauri::http::{Response, StatusCode};
use tauri::{Emitter, Listener, Manager, State};
use tauri_plugin_deep_link::DeepLinkExt;
use tauri_plugin_notification::NotificationExt;
use tokio::io::{AsyncReadExt, AsyncWriteExt};
use tokio::net::{TcpListener, TcpSocket};
use tokio::sync::oneshot;
use tokio::sync::watch;
use tokio::time::{timeout, Duration};
use url::Url;
use uuid::Uuid;

const WORKBOOK_ID: &str = "local-workbook";

/// Minimal HTML used by `--startup-bench`.
///
/// The goal of this mode is to measure the desktop shell + webview startup overhead without
/// depending on the built frontend assets in `apps/desktop/dist`.
const STARTUP_BENCH_HTML: &str =
    r#"<!doctype html><html><head><meta charset="utf-8" /><title>Formula</title></head><body></body></html>"#;

static CLOSE_REQUEST_IN_FLIGHT: AtomicBool = AtomicBool::new(false);

// Canonical Tauri event names exchanged between the Rust host and the frontend.
// Keep this list (and any new additions) in sync with the event allowlist in:
// `apps/desktop/src-tauri/capabilities/main.json`.
//
// Rust -> JS (frontend `listen`):
// - close-prep, close-requested
// - open-file, file-dropped
// - tray-open, tray-new, tray-quit
// - shortcut-quick-open, shortcut-command-palette
// - menu-open, menu-new, menu-save, menu-save-as, menu-print, menu-print-preview, menu-export-pdf, menu-close-window, menu-quit,
//   menu-undo, menu-redo, menu-cut, menu-copy, menu-paste, menu-paste-special, menu-select-all,
//   menu-zoom-in, menu-zoom-out, menu-zoom-reset, menu-about, menu-check-updates, menu-open-release-page
// - startup:window-visible, startup:webview-loaded, startup:first-render, startup:tti, startup:metrics
// - update-check-started, update-check-already-running, update-not-available, update-check-error, update-available
// - oauth-redirect
//
// JS -> Rust (frontend `emit`):
// - open-file-ready, oauth-redirect-ready
// - close-prep-done, close-handled
// - updater-ui-ready, coi-check-result
const OPEN_FILE_EVENT: &str = "open-file";
const OPEN_FILE_READY_EVENT: &str = "open-file-ready";
const OAUTH_REDIRECT_EVENT: &str = "oauth-redirect";
const OAUTH_REDIRECT_READY_EVENT: &str = "oauth-redirect-ready";

// Cross-origin isolation headers required for `globalThis.crossOriginIsolated === true`.
//
// We set these on the `tauri://` protocol responses in production builds so Chromium enables
// `SharedArrayBuffer` (required by Pyodide's worker backend).
const CROSS_ORIGIN_OPENER_POLICY: HeaderName =
    HeaderName::from_static("cross-origin-opener-policy");
const CROSS_ORIGIN_EMBEDDER_POLICY: HeaderName =
    HeaderName::from_static("cross-origin-embedder-policy");

fn apply_cross_origin_isolation_headers(response: &mut Response<Vec<u8>>) {
    let headers = response.headers_mut();
    headers.insert(
        CROSS_ORIGIN_OPENER_POLICY,
        HeaderValue::from_static("same-origin"),
    );
    headers.insert(
        CROSS_ORIGIN_EMBEDDER_POLICY,
        HeaderValue::from_static("require-corp"),
    );
}

type SharedOpenFileState = Arc<Mutex<OpenFileState>>;

type SharedOauthRedirectState = Arc<Mutex<OauthRedirectState>>;

#[derive(Debug, Default)]
struct OauthLoopbackState {
    active_redirect_uris: HashSet<String>,
}

type SharedOauthLoopbackState = Arc<Mutex<OauthLoopbackState>>;

type SharedStartupMetrics = Arc<Mutex<StartupMetrics>>;

#[derive(Clone, Debug, Serialize)]
struct StartupTimingsSnapshot {
    #[serde(rename = "window_visible_ms")]
    window_visible_ms: Option<u64>,
    #[serde(rename = "webview_loaded_ms")]
    webview_loaded_ms: Option<u64>,
    #[serde(rename = "first_render_ms")]
    first_render_ms: Option<u64>,
    #[serde(rename = "tti_ms")]
    tti_ms: Option<u64>,
}

#[derive(Debug)]
struct StartupMetrics {
    start: Instant,
    window_visible_ms: Option<u64>,
    /// Monotonic ms since native process start when the *main webview finished its initial page
    /// load/navigation* (Tauri `PageLoadEvent::Finished`).
    ///
    /// This is intentionally recorded from Rust (via a page-load callback) so it is independent
    /// from frontend bootstrap timing. In particular, it does **not** include any JS execution
    /// time, event listener installation, or "time-to-interactive" work in the renderer.
    webview_loaded_ms: Option<u64>,
    webview_loaded_recorded_from_page_load: bool,
    first_render_ms: Option<u64>,
    tti_ms: Option<u64>,
    logged: bool,
}

impl StartupMetrics {
    fn new(start: Instant) -> Self {
        Self {
            start,
            window_visible_ms: None,
            webview_loaded_ms: None,
            webview_loaded_recorded_from_page_load: false,
            first_render_ms: None,
            tti_ms: None,
            logged: false,
        }
    }

    fn elapsed_ms(&self) -> u64 {
        self.start.elapsed().as_millis() as u64
    }

    fn record_window_visible(&mut self) -> u64 {
        if let Some(ms) = self.window_visible_ms {
            return ms;
        }
        let ms = self.elapsed_ms();
        self.window_visible_ms = Some(ms);
        ms
    }

    fn record_webview_loaded(&mut self) -> u64 {
        if let Some(ms) = self.webview_loaded_ms {
            return ms;
        }
        let ms = self.elapsed_ms();
        self.webview_loaded_ms = Some(ms);
        self.webview_loaded_recorded_from_page_load = false;
        ms
    }

    fn record_webview_loaded_from_page_load(&mut self) -> u64 {
        if self.webview_loaded_recorded_from_page_load {
            return self.webview_loaded_ms.unwrap_or_else(|| self.elapsed_ms());
        }
        let ms = self.elapsed_ms();
        self.webview_loaded_ms = Some(ms);
        self.webview_loaded_recorded_from_page_load = true;
        ms
    }

    fn record_first_render(&mut self) -> u64 {
        if let Some(ms) = self.first_render_ms {
            return ms;
        }
        let ms = self.elapsed_ms();
        self.first_render_ms = Some(ms);
        ms
    }

    fn record_tti(&mut self) -> u64 {
        if let Some(ms) = self.tti_ms {
            return ms;
        }
        let ms = self.elapsed_ms();
        self.tti_ms = Some(ms);
        ms
    }

    fn snapshot(&self) -> StartupTimingsSnapshot {
        StartupTimingsSnapshot {
            window_visible_ms: self.window_visible_ms,
            webview_loaded_ms: self.webview_loaded_ms,
            first_render_ms: self.first_render_ms,
            tti_ms: self.tti_ms,
        }
    }

    fn maybe_log(&mut self) {
        if self.logged {
            return;
        }
        if !should_log_startup_metrics() {
            return;
        }
        if let (Some(window_visible), Some(tti)) = (self.window_visible_ms, self.tti_ms) {
            let webview_loaded = self.webview_loaded_ms;
            let first_render = self.first_render_ms;
            println!(
                "[startup] window_visible_ms={window_visible} webview_loaded_ms={} first_render_ms={} tti_ms={tti}",
                webview_loaded
                    .map(|v| v.to_string())
                    .unwrap_or_else(|| "n/a".to_string()),
                first_render
                    .map(|v| v.to_string())
                    .unwrap_or_else(|| "n/a".to_string())
            );
            // `--startup-bench` exits the process shortly after this line is printed. Be explicit
            // about flushing so CI parsers reliably see the metrics line even with piped stdout.
            #[allow(unused_imports)]
            use std::io::Write as _;
            let _ = std::io::stdout().flush();
            self.logged = true;
        }
    }
}

fn should_log_startup_metrics() -> bool {
    if cfg!(debug_assertions) {
        return true;
    }
    env_flag_truthy("FORMULA_STARTUP_METRICS")
}

#[cfg(not(debug_assertions))]
fn should_disable_startup_update_check() -> bool {
    env_flag_truthy("FORMULA_DISABLE_STARTUP_UPDATE_CHECK")
}

fn env_flag_truthy(name: &str) -> bool {
    match std::env::var(name) {
        Ok(raw) => {
            let v = raw.trim().to_ascii_lowercase();
            !(v.is_empty() || v == "0" || v == "false")
        }
        Err(_) => false,
    }
}

fn spawn_post_window_visible_init(app: &tauri::AppHandle) {
    // Defer any non-critical, potentially expensive startup work until after the first
    // window-visible mark has been recorded.
    //
    // Important: keep this best-effort and never block the UI thread; all work runs on a
    // background thread.
    let trust_shared = app.state::<SharedMacroTrustStore>().inner().clone();
    tauri::async_runtime::spawn_blocking(move || {
        let mut trust_store = trust_shared.lock().unwrap();
        trust_store.ensure_loaded();
    });
}

#[tauri::command]
fn report_startup_webview_loaded(app: tauri::AppHandle, state: State<'_, SharedStartupMetrics>) {
    // The frontend calls this once it has installed its startup listeners (so early-emitted
    // events aren't dropped). The primary `webview_loaded_ms` measurement is recorded from Rust
    // via `Builder::on_page_load` when the main webview finishes its initial navigation; this
    // command is idempotent and will not overwrite an earlier host-recorded timestamp.
    //
    // Note: some frontend code may call this extremely early (before the page-load finished
    // callback fires). In that case we may temporarily record a provisional timestamp here, but
    // the page-load callback will still overwrite it with the authoritative
    // `PageLoadEvent::Finished` mark.
    let shared = state.inner().clone();
    let (window_visible_ms, webview_loaded_ms, snapshot, first_window_visible) = {
        let mut metrics = shared.lock().unwrap();
        let first_window_visible = metrics.window_visible_ms.is_none();
        let window_visible_ms = metrics.record_window_visible();
        let webview_loaded_ms = metrics.record_webview_loaded();
        let snapshot = metrics.snapshot();
        (
            window_visible_ms,
            webview_loaded_ms,
            snapshot,
            first_window_visible,
        )
    };

    if first_window_visible {
        spawn_post_window_visible_init(&app);
    }

    if let Some(window) = app.get_webview_window("main") {
        let _ = window.emit("startup:window-visible", window_visible_ms);
        let _ = window.emit("startup:webview-loaded", webview_loaded_ms);
        let _ = window.emit("startup:metrics", snapshot);
    }
}

#[tauri::command]
fn report_startup_first_render(app: tauri::AppHandle, state: State<'_, SharedStartupMetrics>) {
    let shared = state.inner().clone();
    let (first_render_ms, snapshot) = {
        let mut metrics = shared.lock().unwrap();
        // If we somehow never recorded a window-visible timestamp (e.g. we missed the window
        // event), fall back to "at least by first render the window was visible".
        metrics.record_window_visible();
        let first_render_ms = metrics.record_first_render();
        let snapshot = metrics.snapshot();
        (first_render_ms, snapshot)
    };

    if let Some(window) = app.get_webview_window("main") {
        let _ = window.emit("startup:first-render", first_render_ms);
        let _ = window.emit("startup:metrics", snapshot);
    }
}

#[tauri::command]
fn report_startup_tti(app: tauri::AppHandle, state: State<'_, SharedStartupMetrics>) {
    let shared = state.inner().clone();
    let (tti_ms, snapshot, first_window_visible) = {
        let mut metrics = shared.lock().unwrap();
        let first_window_visible = metrics.window_visible_ms.is_none();
        // If we somehow never recorded a window-visible timestamp (e.g. the webview
        // never called `report_startup_webview_loaded`), fall back to "at least by
        // the time we became interactive the window was visible".
        metrics.record_window_visible();
        let tti_ms = metrics.record_tti();
        let snapshot = metrics.snapshot();
        metrics.maybe_log();
        (tti_ms, snapshot, first_window_visible)
    };

    if first_window_visible {
        spawn_post_window_visible_init(&app);
    }

    if let Some(window) = app.get_webview_window("main") {
        let _ = window.emit("startup:tti", tti_ms);
        let _ = window.emit("startup:metrics", snapshot);
    }
}

#[derive(Clone, Debug, Serialize)]
struct CloseRequestedPayload {
    token: String,
    /// Optional cell updates produced by `Workbook_BeforeClose`.
    ///
    /// Note: if the user cancels the close in the frontend (e.g. via an unsaved-changes prompt),
    /// applying these updates keeps the frontend `DocumentController` consistent with the backend.
    updates: Vec<commands::CellUpdate>,
}

fn show_main_window(app: &tauri::AppHandle) {
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.show();
        let _ = window.set_focus();
    }
}

fn emit_open_file_event(app: &tauri::AppHandle, paths: Vec<String>) {
    if paths.is_empty() {
        return;
    }

    show_main_window(app);

    if let Some(window) = app.get_webview_window("main") {
        let _ = window.emit(OPEN_FILE_EVENT, paths);
    } else {
        let _ = app.emit(OPEN_FILE_EVENT, paths);
    }
}

fn emit_oauth_redirect_event(app: &tauri::AppHandle, url: String) {
    let trimmed = url.trim().trim_matches('"');
    if trimmed.is_empty() {
        return;
    }

    show_main_window(app);

    if let Some(window) = app.get_webview_window("main") {
        let _ = window.emit(OAUTH_REDIRECT_EVENT, trimmed.to_string());
    } else {
        let _ = app.emit(OAUTH_REDIRECT_EVENT, trimmed.to_string());
    }
}

fn normalize_open_file_request_paths(paths: Vec<String>) -> Vec<String> {
    // Best-effort de-dupe (avoids double-opens if both argv and macOS open-document events fire).
    let mut seen = HashSet::<String>::new();
    let mut out = Vec::new();
    for path in paths {
        if path.trim().is_empty() {
            continue;
        }
        if seen.insert(path.clone()) {
            out.push(path);
        }
    }
    out
}

fn normalize_oauth_redirect_request_urls(urls: Vec<String>) -> Vec<String> {
    let mut seen = HashSet::<String>::new();
    let mut out = Vec::new();
    for url in urls {
        let trimmed = url.trim().trim_matches('"');
        if trimmed.is_empty() {
            continue;
        }
        let is_formula = trimmed
            .get(..8)
            .map_or(false, |prefix| prefix.eq_ignore_ascii_case("formula:"));

        // Support RFC 8252 loopback redirects:
        // - http://127.0.0.1:<port>/...
        // - http://localhost:<port>/...
        // - http://[::1]:<port>/...
        let is_loopback = if !is_formula {
            desktop::oauth_loopback::parse_loopback_redirect_uri(trimmed).is_ok()
        } else {
            false
        };

        if !is_formula && !is_loopback {
            continue;
        }
        let normalized = trimmed.to_string();
        if seen.insert(normalized.clone()) {
            out.push(normalized);
        }
    }
    out
}

fn handle_open_file_request(app: &tauri::AppHandle, paths: Vec<String>) {
    let paths = normalize_open_file_request_paths(paths);
    if paths.is_empty() {
        // Still focus the existing window on "warm start" launches with no file args.
        show_main_window(app);
        return;
    }

    show_main_window(app);

    let open_file_state = app.state::<SharedOpenFileState>().inner().clone();
    let maybe_emit = {
        let mut state = open_file_state.lock().unwrap();
        state.queue_or_emit(paths)
    };

    if let Some(paths) = maybe_emit {
        emit_open_file_event(app, paths);
    }
}

fn handle_oauth_redirect_request(app: &tauri::AppHandle, urls: Vec<String>) {
    let urls = normalize_oauth_redirect_request_urls(urls);
    if urls.is_empty() {
        return;
    }

    show_main_window(app);

    let redirect_state = app.state::<SharedOauthRedirectState>().inner().clone();
    let maybe_emit = {
        let mut state = redirect_state.lock().unwrap();
        state.queue_or_emit(urls)
    };

    if let Some(urls) = maybe_emit {
        for url in urls {
            emit_oauth_redirect_event(app, url);
        }
    }
}

fn extract_open_file_paths(argv: &[String], cwd: Option<&Path>) -> Vec<String> {
    open_file::extract_open_file_paths_from_argv(argv, cwd)
        .into_iter()
        .map(|path| path.to_string_lossy().to_string())
        .collect()
}

fn extract_oauth_redirect_urls(argv: &[String]) -> Vec<String> {
    argv.iter()
        .filter_map(|arg| {
            let trimmed = arg.trim().trim_matches('"');
            if trimmed
                .get(..8)
                .map_or(false, |prefix| prefix.eq_ignore_ascii_case("formula:"))
            {
                Some(trimmed.to_string())
            } else {
                None
            }
        })
        .collect()
}

fn cwd_from_single_instance_callback(cwd: String) -> Option<PathBuf> {
    let trimmed = cwd.trim();
    if trimmed.is_empty() {
        return None;
    }
    Some(PathBuf::from(trimmed))
}

fn signature_status(
    vba_project_bin: &[u8],
    vba_project_signature_bin: Option<&[u8]>,
) -> commands::MacroSignatureStatus {
    // Match `formula_xlsx::XlsxPackage::verify_vba_digital_signature` behavior:
    // - Prefer the signature-part signature when it cryptographically verifies.
    // - Otherwise, fall back to an embedded signature inside `vbaProject.bin`.
    // - If neither verifies, return the best-effort signature info (parse errors included).
    let mut signature_part_result: Option<formula_vba::VbaDigitalSignature> = None;
    if let Some(sig_part) = vba_project_signature_bin {
        match formula_vba::verify_vba_digital_signature_with_project(vba_project_bin, sig_part) {
            Ok(Some(sig)) => signature_part_result = Some(sig),
            Ok(None) => {}
            Err(_) => {
                // Not an OLE container: fall back to verifying the part bytes as a raw PKCS#7/CMS
                // signature blob.
                let (verification, signer_subject) =
                    formula_vba::verify_vba_signature_blob(sig_part);
                signature_part_result = Some(formula_vba::VbaDigitalSignature {
                    stream_path: "xl/vbaProjectSignature.bin".to_string(),
                    stream_kind: formula_vba::VbaSignatureStreamKind::Unknown,
                    signer_subject,
                    signature: sig_part.to_vec(),
                    verification,
                    binding: formula_vba::VbaSignatureBinding::Unknown,
                });
            }
        }
    }

    if let Some(sig) = signature_part_result.as_mut() {
        if sig.verification == formula_vba::VbaSignatureVerification::SignedVerified
            && sig.binding == formula_vba::VbaSignatureBinding::Unknown
        {
            sig.binding = match formula_vba::verify_vba_project_signature_binding(
                vba_project_bin,
                &sig.signature,
            ) {
                Ok(binding) => match binding {
                    formula_vba::VbaProjectBindingVerification::BoundVerified(_) => {
                        formula_vba::VbaSignatureBinding::Bound
                    }
                    formula_vba::VbaProjectBindingVerification::BoundMismatch(_) => {
                        formula_vba::VbaSignatureBinding::NotBound
                    }
                    formula_vba::VbaProjectBindingVerification::BoundUnknown(_) => {
                        formula_vba::VbaSignatureBinding::Unknown
                    }
                },
                Err(_) => formula_vba::VbaSignatureBinding::Unknown,
            };
        }
    }

    let embedded = formula_vba::verify_vba_digital_signature(vba_project_bin)
        .ok()
        .flatten();

    let parsed = if signature_part_result.as_ref().is_some_and(|sig| {
        sig.verification == formula_vba::VbaSignatureVerification::SignedVerified
    }) {
        signature_part_result
    } else if embedded.as_ref().is_some_and(|sig| {
        sig.verification == formula_vba::VbaSignatureVerification::SignedVerified
    }) {
        embedded
    } else {
        signature_part_result.or(embedded)
    };

    match parsed {
        Some(sig) => match sig.verification {
            formula_vba::VbaSignatureVerification::SignedVerified => match sig.binding {
                formula_vba::VbaSignatureBinding::Bound => {
                    commands::MacroSignatureStatus::SignedVerified
                }
                formula_vba::VbaSignatureBinding::NotBound => {
                    commands::MacroSignatureStatus::SignedInvalid
                }
                formula_vba::VbaSignatureBinding::Unknown => {
                    commands::MacroSignatureStatus::SignedUnverified
                }
            },
            formula_vba::VbaSignatureVerification::SignedInvalid => {
                commands::MacroSignatureStatus::SignedInvalid
            }
            formula_vba::VbaSignatureVerification::SignedParseError => {
                commands::MacroSignatureStatus::SignedParseError
            }
            formula_vba::VbaSignatureVerification::SignedButUnverified => {
                commands::MacroSignatureStatus::SignedUnverified
            }
        },
        None => commands::MacroSignatureStatus::Unsigned,
    }
}

fn macros_trusted_for_before_close(
    state: &mut AppState,
    trust_store: &MacroTrustStore,
) -> Result<bool, String> {
    let workbook = match state.get_workbook_mut() {
        Ok(workbook) => workbook,
        Err(_) => return Ok(false),
    };

    let Some(vba_bin) = workbook.vba_project_bin.as_deref() else {
        return Ok(false);
    };

    let fingerprint = if let Some(fp) = workbook.macro_fingerprint.as_deref() {
        fp.to_string()
    } else {
        let identity = workbook
            .origin_path
            .as_deref()
            .or(workbook.path.as_deref())
            .unwrap_or(WORKBOOK_ID);
        let fp = compute_macro_fingerprint(identity, vba_bin);
        workbook.macro_fingerprint = Some(fp.clone());
        fp
    };

    let trust = trust_store.trust_state(&fingerprint);

    // Prefer the in-memory signature part payload when available, because some workflows drop the
    // original XLSX bytes (forcing regeneration on save) while still preserving VBA payloads.
    let mut sig_part_fallback: Option<Vec<u8>> = None;
    if workbook.vba_project_signature_bin.is_none() {
        sig_part_fallback = workbook.origin_xlsx_bytes.as_deref().and_then(|origin| {
            formula_xlsx::read_part_from_reader(
                std::io::Cursor::new(origin),
                "xl/vbaProjectSignature.bin",
            )
            .ok()
            .flatten()
        });
    }
    let sig_part = workbook
        .vba_project_signature_bin
        .as_deref()
        .or(sig_part_fallback.as_deref());
    let sig_status = signature_status(vba_bin, sig_part);
    Ok(commands::evaluate_macro_trust(trust, sig_status).is_ok())
}

fn cell_update_from_state(update: CellUpdateData) -> commands::CellUpdate {
    commands::CellUpdate {
        sheet_id: update.sheet_id,
        row: update.row,
        col: update.col,
        value: update.value.as_json(),
        formula: update.formula,
        display_value: update.display_value,
    }
}

#[tauri::command]
async fn show_system_notification(
    window: tauri::WebviewWindow,
    title: String,
    body: Option<String>,
) -> Result<(), String> {
    desktop::ipc_origin::ensure_main_window(
        window.label(),
        "notifications",
        desktop::ipc_origin::Verb::Are,
    )?;

    let url = window.url().map_err(|err| err.to_string())?;
    desktop::ipc_origin::ensure_trusted_origin(
        &url,
        "notifications",
        desktop::ipc_origin::Verb::Are,
    )?;

    let mut builder = window.app_handle().notification().builder().title(title);

    if let Some(body) = body {
        builder = builder.body(body);
    }

    builder.show().map_err(|err| err.to_string())?;
    Ok(())
}

#[tauri::command]
async fn oauth_loopback_listen(
    window: tauri::WebviewWindow,
    state: State<'_, SharedOauthLoopbackState>,
    redirect_uri: String,
) -> Result<(), String> {
    desktop::ipc_origin::ensure_main_window(
        window.label(),
        "oauth loopback listeners",
        desktop::ipc_origin::Verb::Are,
    )?;

    let url = window.url().map_err(|err| err.to_string())?;
    desktop::ipc_origin::ensure_trusted_origin(
        &url,
        "oauth loopback listeners",
        desktop::ipc_origin::Verb::Are,
    )?;

    let app = window.app_handle();

    let parsed = desktop::oauth_loopback::parse_loopback_redirect_uri(&redirect_uri)?;
    let host_kind = parsed.host_kind;
    let port = parsed.port;
    let expected_path = parsed.path;
    let redirect_uri = parsed.normalized_redirect_uri;

    let shared = state.inner().clone();
    {
        let mut guard = shared.lock().unwrap();
        if guard.active_redirect_uris.contains(&redirect_uri) {
            return Ok(());
        }
        guard.active_redirect_uris.insert(redirect_uri.clone());
    }

    let mut listeners: Vec<TcpListener> = Vec::new();
    let mut listener_errors: Vec<String> = Vec::new();

    let wants_ipv4 = matches!(
        host_kind,
        desktop::oauth_loopback::LoopbackHostKind::Ipv4Loopback
            | desktop::oauth_loopback::LoopbackHostKind::Localhost
    );
    let wants_ipv6 = matches!(
        host_kind,
        desktop::oauth_loopback::LoopbackHostKind::Ipv6Loopback
            | desktop::oauth_loopback::LoopbackHostKind::Localhost
    );

    if wants_ipv4 {
        match TcpListener::bind((Ipv4Addr::LOCALHOST, port)).await {
            Ok(listener) => listeners.push(listener),
            Err(err) => listener_errors.push(err.to_string()),
        }
    }

    if wants_ipv6 {
        let addr = SocketAddr::from((Ipv6Addr::LOCALHOST, port));
        let listener = (|| -> std::io::Result<TcpListener> {
            let socket = TcpSocket::new_v6()?;
            socket.set_only_v6(true)?;
            socket.bind(addr)?;
            socket.listen(1024)
        })();
        match listener {
            Ok(listener) => listeners.push(listener),
            Err(err) => listener_errors.push(err.to_string()),
        }
    }

    if listeners.is_empty() {
        if let Ok(mut guard) = shared.lock() {
            guard.active_redirect_uris.remove(&redirect_uri);
        }
        let details = listener_errors.join("; ");
        return Err(match host_kind {
            desktop::oauth_loopback::LoopbackHostKind::Ipv4Loopback => format!(
                "Failed to bind loopback OAuth redirect listener on 127.0.0.1:{port}: {details}"
            ),
            desktop::oauth_loopback::LoopbackHostKind::Ipv6Loopback => format!(
                "Failed to bind loopback OAuth redirect listener on [::1]:{port}: {details}"
            ),
            desktop::oauth_loopback::LoopbackHostKind::Localhost => format!(
                "Failed to bind loopback OAuth redirect listener on localhost:{port}: {details}"
            ),
        });
    }

    let app_handle = app.clone();
    tauri::async_runtime::spawn(async move {
        struct ActiveGuard {
            state: SharedOauthLoopbackState,
            key: String,
        }
        impl Drop for ActiveGuard {
            fn drop(&mut self) {
                if let Ok(mut guard) = self.state.lock() {
                    guard.active_redirect_uris.remove(&self.key);
                }
            }
        }
        let _guard = ActiveGuard {
            state: shared.clone(),
            key: redirect_uri.clone(),
        };

        let overall_timeout = Duration::from_secs(5 * 60);
        let handled = Arc::new(AtomicBool::new(false));
        let (stop_tx, stop_rx) = watch::channel(false);

        let mut tasks = Vec::new();
        for listener in listeners {
            let app_handle = app_handle.clone();
            let expected_path = expected_path.clone();
            let redirect_uri = redirect_uri.clone();
            let handled = handled.clone();
            let stop_tx = stop_tx.clone();
            let mut stop_rx = stop_rx.clone();

            tasks.push(tauri::async_runtime::spawn(async move {
                loop {
                    if *stop_rx.borrow() {
                        break;
                    }

                    tokio::select! {
                        _ = stop_rx.changed() => {
                            continue;
                        }
                        res = listener.accept() => {
                            let (mut socket, _) = match res {
                                Ok(v) => v,
                                Err(_) => break,
                            };

                            let mut buf = vec![0_u8; 8192];
                            let n = match timeout(Duration::from_secs(2), socket.read(&mut buf)).await {
                                Ok(Ok(n)) => n,
                                _ => 0,
                            };
                            if n == 0 {
                                continue;
                            }
                            buf.truncate(n);
                            let req = String::from_utf8_lossy(&buf);
                            let line = req.lines().next().unwrap_or("");
                            let mut parts = line.split_whitespace();
                            let method = parts.next().unwrap_or("");
                            let target = parts.next().unwrap_or("");

                            if method != "GET" {
                                let _ = socket
                                    .write_all(b"HTTP/1.1 405 Method Not Allowed\r\nContent-Length: 0\r\n\r\n")
                                    .await;
                                continue;
                            }

                            // The request target should be a path+query (e.g. `/callback?code=...`), but handle
                            // absolute-form targets defensively.
                            let target = if target.starts_with("http://") || target.starts_with("https://") {
                                Url::parse(target)
                                    .ok()
                                    .map(|u| {
                                        let mut out = u.path().to_string();
                                        if let Some(q) = u.query() {
                                            out.push('?');
                                            out.push_str(q);
                                        }
                                        out
                                    })
                                    .unwrap_or_else(|| target.to_string())
                            } else {
                                target.to_string()
                            };

                            let mut split = target.splitn(2, '?');
                            let path = split.next().unwrap_or("");
                            let query = split.next();

                            if path != expected_path {
                                let _ = socket
                                    .write_all(b"HTTP/1.1 404 Not Found\r\nContent-Length: 0\r\n\r\n")
                                    .await;
                                continue;
                            }

                            let mut full = match Url::parse(&redirect_uri) {
                                Ok(u) => u,
                                Err(_) => break,
                            };
                            full.set_query(query);
                            full.set_fragment(None);
                            let full_url = full.to_string();

                            if handled
                                .compare_exchange(false, true, Ordering::AcqRel, Ordering::Acquire)
                                .is_ok()
                            {
                                handle_oauth_redirect_request(&app_handle, vec![full_url.clone()]);
                                let _ = stop_tx.send(true);
                            }

                            let body = "<!doctype html><html><head><meta charset=\"utf-8\" /><title>Formula</title></head><body><h1>Sign-in complete</h1><p>You can close this window and return to Formula.</p></body></html>";
                            let resp = format!(
                                "HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {}\r\n\r\n{}",
                                body.len(),
                                body
                            );
                            let _ = socket.write_all(resp.as_bytes()).await;
                            let _ = socket.shutdown().await;
                            break;
                        }
                    }
                }
            }));
        }

        let mut stop_rx_wait = stop_rx.clone();
        let wait_for_stop = async move {
            loop {
                if *stop_rx_wait.borrow() {
                    break;
                }
                if stop_rx_wait.changed().await.is_err() {
                    break;
                }
            }
        };

        let _ = timeout(overall_timeout, wait_for_stop).await;
        let _ = stop_tx.send(true);

        for task in tasks {
            let _ = task.await;
        }
    });

    Ok(())
}

fn main() {
    // Record a monotonic startup timestamp as early as possible so we can measure
    // cold start time-to-window-visible / time-to-interactive.
    let startup_metrics: SharedStartupMetrics =
        Arc::new(Mutex::new(StartupMetrics::new(Instant::now())));
    let startup_metrics_for_page_load = startup_metrics.clone();

    let state: SharedAppState = Arc::new(Mutex::new(AppState::new()));
    // Create the trust store *without* reading from disk so cold start can get the window on
    // screen faster. We'll load persisted trust decisions asynchronously after the webview has
    // reported that the window is visible.
    //
    // Security note: while the store is not yet loaded, macros are default-deny (blocked).
    let macro_trust: SharedMacroTrustStore =
        Arc::new(Mutex::new(MacroTrustStore::new_unloaded_default()));

    let open_file_state: SharedOpenFileState = Arc::new(Mutex::new(OpenFileState::default()));
    let oauth_redirect_state: SharedOauthRedirectState =
        Arc::new(Mutex::new(OauthRedirectState::default()));
    let oauth_loopback_state: SharedOauthLoopbackState =
        Arc::new(Mutex::new(OauthLoopbackState::default()));
    let initial_argv: Vec<String> = std::env::args().collect();
    let startup_bench = initial_argv.iter().any(|arg| arg == "--startup-bench");
    if startup_bench {
        // In production builds we normally gate startup metrics logging behind
        // `FORMULA_STARTUP_METRICS=1`. The `--startup-bench` mode is explicitly for CI
        // measurement, so opt-in automatically.
        std::env::set_var("FORMULA_STARTUP_METRICS", "1");
    }
    if initial_argv.iter().any(|arg| arg == "--log-process-metrics") {
        process_metrics::log_process_metrics();
    }
    let initial_cwd = std::env::current_dir().ok();
    let initial_paths = normalize_open_file_request_paths(extract_open_file_paths(
        &initial_argv,
        initial_cwd.as_deref(),
    ));
    if !initial_paths.is_empty() {
        let mut guard = open_file_state.lock().unwrap();
        guard.queue_or_emit(initial_paths);
    }

    let initial_oauth_urls =
        normalize_oauth_redirect_request_urls(extract_oauth_redirect_urls(&initial_argv));
    if !initial_oauth_urls.is_empty() {
        let mut guard = oauth_redirect_state.lock().unwrap();
        guard.queue_or_emit(initial_oauth_urls);
    }

    let app = tauri::Builder::default()
        // Override Tauri's default `asset:` protocol handler to attach COEP-friendly headers.
        // See `asset_protocol.rs` for details.
        .register_uri_scheme_protocol("asset", asset_protocol::handler)
        // In production builds, the webview loads `frontendDist` via Tauri's custom
        // asset protocol (`tauri://...`). Unlike the Vite dev/preview servers, those
        // responses don't include COOP/COEP headers by default, which prevents
        // `globalThis.crossOriginIsolated` from becoming true and disables
        // `SharedArrayBuffer` in Chromium.
        //
        // Inject COOP/COEP headers into the `tauri://` protocol responses so we can use
        // `SharedArrayBuffer` (required by Pyodide).
        //
        // Note: Tauri's internal asset protocol handler is not a stable public API, so we
        // implement a minimal handler using the public `AssetResolver`.
        .register_uri_scheme_protocol("tauri", move |_ctx, request| {
            let path = request.uri().path();

            // Lightweight shell-startup benchmark: serve a tiny inline HTML document instead of
            // the real bundled frontend (which may not be present, and is expensive to build).
            if startup_bench && (path == "/" || path == "/index.html") {
                let mut builder = Response::builder()
                    .status(StatusCode::OK)
                    .header(tauri::http::header::CONTENT_TYPE, "text/html; charset=utf-8");

                if let Some(csp) = _ctx.app_handle().config().app.security.csp.as_ref() {
                    builder = builder.header("Content-Security-Policy", csp.as_str());
                }

                let mut response = builder
                    .body(STARTUP_BENCH_HTML.as_bytes().to_vec())
                    .unwrap_or_else(|_| {
                        Response::builder()
                            .status(StatusCode::INTERNAL_SERVER_ERROR)
                            .header(tauri::http::header::CONTENT_TYPE, "text/plain")
                            .body(b"failed to build tauri startup-bench response".to_vec())
                            .expect("build error response")
                    });
                apply_cross_origin_isolation_headers(&mut response);
                return response;
            }

            let path = path.to_string();
            match _ctx.app_handle().asset_resolver().get(path) {
                Some(asset) => {
                    let mut builder = Response::builder()
                        .status(StatusCode::OK)
                        .header(tauri::http::header::CONTENT_TYPE, asset.mime_type);

                    if let Some(csp) = asset.csp_header {
                        builder = builder.header("Content-Security-Policy", csp);
                    }

                    let mut response = builder.body(asset.bytes).unwrap_or_else(|_| {
                        Response::builder()
                            .status(StatusCode::INTERNAL_SERVER_ERROR)
                            .header(tauri::http::header::CONTENT_TYPE, "text/plain")
                            .body(b"failed to build tauri asset response".to_vec())
                            .expect("build error response")
                    });

                    apply_cross_origin_isolation_headers(&mut response);
                    response
                }
                None => Response::builder()
                    .status(StatusCode::NOT_FOUND)
                    .header(tauri::http::header::CONTENT_TYPE, "text/plain")
                    .body(b"asset not found".to_vec())
                    .expect("build not-found response"),
            }
        })
        // Core platform plugins used by the frontend (dialog, clipboard, shell).
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_clipboard_manager::init())
        .plugin(tauri_plugin_shell::init())
        .plugin(tauri_plugin_single_instance::init(|app, argv, cwd| {
            // OAuth PKCE deep-link redirect capture (e.g. `formula://oauth/callback?...`).
            //
            // When an OAuth provider redirects to our custom URI scheme, the OS may attempt to
            // launch a second instance of the application. The single-instance plugin forwards
            // the argv to the running instance; queue/emit the URL to the frontend so it can
            // resolve any pending `DesktopOAuthBroker.waitForRedirect(...)` promises.
            let oauth_urls = extract_oauth_redirect_urls(&argv);
            handle_oauth_redirect_request(app, oauth_urls);

            // File association / open-with handling.
            let cwd = cwd_from_single_instance_callback(cwd);
            let paths = extract_open_file_paths(&argv, cwd.as_deref());
            handle_open_file_request(app, paths);
        }))
        .plugin(
            tauri_plugin_global_shortcut::Builder::new()
                .with_handler(
                    |app, shortcut, _event| match shortcut.to_string().as_str() {
                        "CmdOrCtrl+Shift+O" => {
                            let _ = app.emit("shortcut-quick-open", ());
                        }
                        "CmdOrCtrl+Shift+P" => {
                            let _ = app.emit("shortcut-command-palette", ());
                        }
                        _ => {}
                    },
                )
                .build(),
        )
        .plugin(tauri_plugin_deep_link::init())
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_notification::init())
        .manage(state)
        .manage(macro_trust)
        .manage(open_file_state)
        .manage(oauth_redirect_state)
        .manage(oauth_loopback_state)
        .manage(TrayStatusState::default())
        .manage(startup_metrics)
        .on_page_load(move |window, payload| {
            // Best-effort startup instrumentation: record when the *native webview* finishes
            // loading the initial app page. This is intentionally independent of frontend JS
            // bootstrap timing.
            if window.label() != "main" {
                return;
            }

            // Only record once, and only when the page load finished.
            let finished = matches!(payload.event(), tauri::PageLoadEvent::Finished);
            if !finished {
                return;
            }

            // Some platforms may briefly load an internal blank page during WebView creation.
            // Avoid treating that as "app page loaded".
            let url = payload.url().to_string();
            if url == "about:blank" {
                return;
            }

            let (webview_loaded_ms, snapshot) = {
                let mut metrics = match startup_metrics_for_page_load.lock() {
                    Ok(guard) => guard,
                    Err(poisoned) => poisoned.into_inner(),
                };

                if metrics.webview_loaded_recorded_from_page_load {
                    return;
                }

                let webview_loaded_ms = metrics.record_webview_loaded_from_page_load();
                let snapshot = metrics.snapshot();
                (webview_loaded_ms, snapshot)
            };

            let _ = window.emit("startup:webview-loaded", webview_loaded_ms);
            let _ = window.emit("startup:metrics", snapshot);
        })
        // NOTE: IPC hardening / capabilities (Tauri v2)
        //
        // We avoid `core:default` and instead grant only the plugin APIs the frontend uses
        // (events, dialogs, window ops, clipboard (plain text), updater, ...) in
        // `src-tauri/capabilities/main.json`, scoped to the `main` window.
        //
        // Any new `#[tauri::command]` must be:
        //  1) Implemented in Rust (typically `src/commands.rs`, but may live elsewhere)
        //  2) Registered here in `generate_handler![...]`
        //  3) Added to the explicit JS invoke allowlist in
        //     `src-tauri/permissions/allow-invoke.json` (`allow-invoke` permission)
        //
        // Note: Tauri command invocation can be allowlisted in two ways:
        // - `allow-invoke` (application permission defined in `src-tauri/permissions/allow-invoke.json`)
        // - `core:allow-invoke` (optional core permission supported by some toolchains)
        //
        // IMPORTANT:
        // - Never grant the string form `"core:allow-invoke"` (it enables the default/unscoped allowlist).
        // - If `core:allow-invoke` is present in `src-tauri/capabilities/main.json`, it must use the object form:
        //   `{ "identifier": "core:allow-invoke", "allow": [{ "command": "..." }, ...] }`
        //   and stay explicit + in sync with `allow-invoke.json`.
        //
        // Guardrails:
        // - `apps/desktop/src-tauri/tests/tauri_ipc_allowlist.rs` asserts this
        //   `generate_handler![...]` list matches `src-tauri/permissions/allow-invoke.json`.
        // - `apps/desktop/src-tauri/tests/ipc_origin_guardrails.rs` asserts privileged commands
        //   include runtime origin hardening (`ipc_origin::ensure_main_window` +
        //   `ipc_origin::ensure_trusted_origin`) as defense-in-depth in case untrusted content is
        //   ever loaded into a WebView.
        // - `apps/desktop/src/tauri/__tests__/capabilitiesPermissions.vitest.ts` asserts:
        //   - the `allow-invoke` permission is granted to the main window
        //   - the allowlist stays explicit (no wildcards) and covers frontend `invoke("...")` usage
        //   - we don't grant the unscoped string form `"core:allow-invoke"` (and if `core:allow-invoke` is present, it is the
        //     object form with an explicit per-command allowlist)
        //   - the plugin permission surface stays minimal/explicit (dialogs/window ops/clipboard/updater)
        //
        // Note: we intentionally do not grant the JS shell plugin API (`shell:allow-open`);
        // external URL opening goes through the `open_external_url` Rust command which enforces a
        // scheme allowlist.
        //
        // SECURITY: `allow-invoke` only gates *which command names* can be invoked.
        // Commands touching filesystem/network/etc must still validate inputs and enforce
        // scoping/authorization in Rust (trusted-origin + window-label checks, path/network scopes, etc).
        .invoke_handler(tauri::generate_handler![
            clipboard::clipboard_read,
            clipboard::clipboard_write,
            ed25519_verifier::verify_ed25519_signature,
            commands::open_workbook,
            commands::new_workbook,
            commands::add_sheet,
            commands::add_sheet_with_id,
            commands::rename_sheet,
            commands::move_sheet,
            commands::delete_sheet,
            commands::reorder_sheets,
            commands::set_sheet_visibility,
            commands::set_sheet_tab_color,
            commands::read_text_file,
            commands::read_binary_file,
            commands::read_binary_file_range,
            commands::stat_file,
            commands::list_dir,
            commands::open_external_url,
            commands::read_clipboard,
            commands::write_clipboard,
            commands::power_query_cache_key_get_or_create,
            commands::collab_encryption_key_get,
            commands::collab_encryption_key_set,
            commands::collab_encryption_key_delete,
            commands::collab_encryption_key_list,
            commands::collab_token_get,
            commands::collab_token_set,
            commands::collab_token_delete,
            commands::power_query_credential_get,
            commands::power_query_credential_set,
            commands::power_query_credential_delete,
            commands::power_query_credential_list,
            commands::power_query_refresh_state_get,
            commands::power_query_refresh_state_set,
            commands::power_query_state_get,
            commands::power_query_state_set,
            commands::sql_query,
            commands::sql_get_schema,
            commands::save_workbook,
            commands::mark_saved,
            commands::get_workbook_theme_palette,
            commands::list_defined_names,
            commands::list_tables,
            commands::list_imported_chart_models,
            commands::get_precedents,
            commands::get_dependents,
            commands::set_cell,
            commands::get_range,
            commands::get_sheet_used_range,
            commands::get_sheet_formatting,
            commands::set_range,
            commands::apply_sheet_formatting_deltas,
            commands::create_pivot_table,
            commands::refresh_pivot_table,
            commands::list_pivot_tables,
            commands::get_sheet_print_settings,
            commands::set_sheet_page_setup,
            commands::set_sheet_print_area,
            commands::export_sheet_range_pdf,
            commands::get_vba_project,
            commands::list_macros,
            commands::get_macro_security_status,
            commands::set_macro_trust,
            commands::set_macro_ui_context,
            commands::run_macro,
            commands::validate_vba_migration,
            commands::run_python_script,
            commands::network_fetch,
            commands::marketplace_search,
            commands::marketplace_get_extension,
            commands::marketplace_download_package,
            commands::check_for_updates,
            commands::quit_app,
            commands::restart_app,
            commands::fire_workbook_open,
            commands::fire_workbook_before_close,
            commands::fire_worksheet_change,
            commands::fire_selection_change,
            tray_status::set_tray_status,
            show_system_notification,
            oauth_loopback_listen,
            updater::install_downloaded_update,
            report_startup_webview_loaded,
            report_startup_first_render,
            report_startup_tti,
        ])
        .on_menu_event(|app, event| {
            menu::on_menu_event(app, event);
        })
        .on_window_event(|window, event| match event {
            tauri::WindowEvent::Resized(_)
            | tauri::WindowEvent::Moved(_)
            | tauri::WindowEvent::Focused(true) => {
                if window.label() == "main" {
                    let startup = window.state::<SharedStartupMetrics>().inner().clone();
                    let first_window_visible = {
                        let mut metrics = startup.lock().unwrap();
                        let first = metrics.window_visible_ms.is_none();
                        metrics.record_window_visible();
                        first
                    };
                    if first_window_visible {
                        spawn_post_window_visible_init(window.app_handle());
                    }
                }
            }
            tauri::WindowEvent::CloseRequested { api, .. } => {
                // Keep the process alive so the tray icon stays available.
                api.prevent_close();

                // Avoid running multiple overlapping close flows / macros if the user triggers
                // repeated close requests while a close prompt is still in flight.
                if CLOSE_REQUEST_IN_FLIGHT.swap(true, Ordering::SeqCst) {
                    return;
                }

                let window = window.clone();
                let shared_state = window.state::<SharedAppState>().inner().clone();
                let shared_trust = window.state::<SharedMacroTrustStore>().inner().clone();

                tauri::async_runtime::spawn(async move {
                    struct CloseRequestGuard;
                    impl Drop for CloseRequestGuard {
                        fn drop(&mut self) {
                            CLOSE_REQUEST_IN_FLIGHT.store(false, Ordering::SeqCst);
                        }
                    }
                    let _guard = CloseRequestGuard;

                    // Best-effort Workbook_BeforeClose. We do this in a background task so we
                    // don't block the window event loop. Cancellation isn't supported yet.
                    //
                    // We ask the frontend to drain any pending workbook-sync operations and to
                    // sync the macro UI context before we run the event macro. This avoids
                    // running the macro against stale backend state.
                    let token = Uuid::new_v4().to_string();
                    let (tx, rx) = oneshot::channel::<()>();
                    let tx = Arc::new(Mutex::new(Some(tx)));
                    let token_for_listener = token.clone();
                    let tx_for_listener = tx.clone();

                    let handler = window.listen("close-prep-done", move |event| {
                        let received = event.payload().trim().trim_matches('"');
                        if received != token_for_listener {
                            return;
                        }
                        if let Ok(mut guard) = tx_for_listener.lock() {
                            if let Some(sender) = guard.take() {
                                let _ = sender.send(());
                            }
                        }
                    });

                    let _ = window.emit("close-prep", token.clone());
                    let _ = timeout(Duration::from_millis(750), rx).await;
                    window.unlisten(handler);

                    let state_for_macro = shared_state.clone();
                    let trust_for_macro = shared_trust.clone();
                    let macro_outcome = tauri::async_runtime::spawn_blocking(move || {
                        let mut state = state_for_macro.lock().unwrap();
                        let mut trust_store = trust_for_macro.lock().unwrap();
                        trust_store.ensure_loaded();

                        let should_run = macros_trusted_for_before_close(&mut state, &trust_store)?;
                        drop(trust_store);

                        if !should_run {
                            return Ok::<_, String>(Vec::new());
                        }

                        let options = MacroExecutionOptions {
                            permissions: Vec::new(),
                            timeout_ms: None,
                        };

                        match state.fire_workbook_before_close(options) {
                            Ok(outcome) => {
                                if outcome.permission_request.is_some() {
                                    eprintln!(
                                        "[macro] Workbook_BeforeClose requested additional permissions; refusing to escalate."
                                    );
                                }
                                if !outcome.ok {
                                    let msg = outcome
                                        .error
                                        .unwrap_or_else(|| "unknown macro error".to_string());
                                    eprintln!("[macro] Workbook_BeforeClose failed: {msg}");
                                }
                                let updates = outcome
                                    .updates
                                    .into_iter()
                                    .map(cell_update_from_state)
                                    .collect();
                                return Ok(updates);
                            }
                            Err(err) => {
                                eprintln!("[macro] Workbook_BeforeClose failed: {err}");
                            }
                        }

                        Ok(Vec::new())
                    })
                    .await;

                    let updates = match macro_outcome {
                        Ok(Ok(updates)) => updates,
                        Ok(Err(err)) => {
                            eprintln!("[macro] Workbook_BeforeClose task failed: {err}");
                            Vec::new()
                        }
                        Err(err) => {
                            eprintln!("[macro] Workbook_BeforeClose task panicked: {err}");
                            Vec::new()
                        }
                    };

                    // Delegate the rest of close-handling to the frontend (unsaved changes prompt
                    // + deciding whether to hide the window or keep it open).
                    let payload = CloseRequestedPayload { token: token.clone(), updates };
                    let _ = window.emit("close-requested", payload);

                    // Wait until the frontend finishes its close flow (e.g. after an unsaved
                    // changes prompt). This keeps `CLOSE_REQUEST_IN_FLIGHT` set while the close
                    // prompt is active so repeated close clicks don't rerun macros.
                    let (handled_tx, handled_rx) = oneshot::channel::<()>();
                    let handled_tx = Arc::new(Mutex::new(Some(handled_tx)));
                    let token_for_handled = token.clone();
                    let handled_tx_for_listener = handled_tx.clone();
                    let handled_handler = window.listen("close-handled", move |event| {
                        let received = event.payload().trim().trim_matches('"');
                        if received != token_for_handled {
                            return;
                        }
                        if let Ok(mut guard) = handled_tx_for_listener.lock() {
                            if let Some(sender) = guard.take() {
                                let _ = sender.send(());
                            }
                        }
                    });
                    let _ = timeout(Duration::from_secs(60), handled_rx).await;
                    window.unlisten(handled_handler);
                });
            }
            tauri::WindowEvent::DragDrop(drag_drop) => {
                if let tauri::DragDropEvent::Drop { paths, .. } = drag_drop {
                    let payload: Vec<String> = paths
                        .iter()
                        .map(|p| p.to_string_lossy().to_string())
                        .collect();
                    let _ = window.emit("file-dropped", payload);
                }
            }
            _ => {}
        })
        .setup(move |app| {
            if startup_bench {
                // CI benchmark: measure desktop shell startup without requiring built frontend
                // assets. This mode should be lightweight and exit quickly.
                const TIMEOUT_SECS: u64 = 20;
                std::thread::spawn(|| {
                    std::thread::sleep(Duration::from_secs(TIMEOUT_SECS));
                    eprintln!(
                        "[formula][startup-bench] timed out after {TIMEOUT_SECS}s (webview did not report)"
                    );
                    std::process::exit(2);
                });

                let Some(window) = app.get_webview_window("main") else {
                    eprintln!("[formula][startup-bench] missing main window");
                    std::process::exit(2);
                };

                window
                    .eval(
                        r#"
(() => {
  const deadline = Date.now() + 10_000;

  const raf = () =>
    new Promise((resolve) => {
      if (typeof requestAnimationFrame === "function") {
        requestAnimationFrame(() => resolve(null));
      } else {
        setTimeout(() => resolve(null), 0);
      }
    });

  let started = false;
  const tick = async () => {
    if (started) return;

    const invoke = globalThis.__TAURI__?.core?.invoke;
    if (typeof invoke !== "function") {
      if (Date.now() > deadline) return;
      setTimeout(tick, 10);
      return;
    }

    started = true;

    // "WebView loaded": the earliest point where the JS bridge is ready to invoke into Rust.
    await invoke("report_startup_webview_loaded");

    // Approximate "time to interactive": a microtask + first frame later.
    await Promise.resolve();
    await raf();

    await invoke("report_startup_tti");

    // Hard-exit after the `[startup] ...` line is printed. Add a tiny delay so stdout is
    // reliably flushed when captured via pipes.
    setTimeout(() => {
      invoke("quit_app").catch(() => {});
    }, 25);
  };

  tick().catch(() => {});
})();
"#,
                    )
                    .unwrap_or_else(|err| {
                        eprintln!("[formula][startup-bench] failed to eval script: {err}");
                        std::process::exit(2);
                    });

                // Skip the rest of normal app setup (tray icon, updater, open-file wiring, etc).
                // The benchmark mode should be as lightweight as possible so it can run in CI and
                // exit quickly based on the injected JS invocations.
                return Ok(());
            }

            if std::env::args().any(|arg| arg == "--cross-origin-isolation-check") {
                // CI/developer smoke test: validate cross-origin isolation (COOP/COEP) in the
                // packaged Tauri build by running in a special mode that exits quickly with a
                // status code.
                //
                // This is evaluated inside the WebView so we can check `globalThis.crossOriginIsolated`,
                // `SharedArrayBuffer` availability, and basic Worker support (for the Pyodide worker backend).
                const TIMEOUT_SECS: u64 = 20;
                std::thread::spawn(|| {
                    std::thread::sleep(Duration::from_secs(TIMEOUT_SECS));
                    eprintln!(
                        "[formula][coi-check] timed out after {TIMEOUT_SECS}s (webview did not report)"
                    );
                    std::process::exit(2);
                });

                let Some(window) = app.get_webview_window("main") else {
                    eprintln!("[formula][coi-check] missing main window");
                    std::process::exit(2);
                };

                #[derive(Debug, Deserialize)]
                struct CrossOriginIsolationCheckResult {
                    cross_origin_isolated: bool,
                    shared_array_buffer: bool,
                    worker_ok: bool,
                }

                window.listen("coi-check-result", |event| {
                    let payload = event.payload();
                    if payload.trim().is_empty() {
                        eprintln!("[formula][coi-check] missing payload");
                        std::process::exit(2);
                    }
                    let parsed: CrossOriginIsolationCheckResult =
                        match serde_json::from_str(payload) {
                            Ok(parsed) => parsed,
                            Err(err) => {
                                eprintln!(
                                    "[formula][coi-check] invalid payload {payload:?}: {err}"
                                );
                                std::process::exit(2);
                            }
                        };

                    println!(
                        "[formula][coi-check] crossOriginIsolated={}, SharedArrayBuffer={}, workerOk={}",
                        parsed.cross_origin_isolated, parsed.shared_array_buffer, parsed.worker_ok
                    );

                    let ok = parsed.cross_origin_isolated && parsed.shared_array_buffer && parsed.worker_ok;
                    std::process::exit(if ok { 0 } else { 1 });
                });

                window
                    .eval(
                        r#"
 (() => {
   const deadline = Date.now() + 10_000;
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const runWorker = (url, opts) =>
    new Promise((resolve) => {
      let worker;
      try {
        worker = opts ? new Worker(url, opts) : new Worker(url);
      } catch {
        resolve(false);
        return;
      }

      const timeout = setTimeout(() => {
        try {
          worker.terminate();
        } catch {}
        resolve(false);
      }, 1000);

      worker.onmessage = () => {
        clearTimeout(timeout);
        try {
          worker.terminate();
        } catch {}
        resolve(true);
      };
      worker.onerror = () => {
        clearTimeout(timeout);
        try {
          worker.terminate();
        } catch {}
        resolve(false);
      };

      try {
        worker.postMessage(null);
      } catch {
        clearTimeout(timeout);
        try {
          worker.terminate();
        } catch {}
        resolve(false);
      }
    });

  const checkWorker = async () => {
    if (typeof Worker === "undefined") return false;

    // Use a real `self` URL (not `blob:`) so we validate the packaged asset protocol + CSP.
    const workerUrl = new URL("coi-check-worker.js", globalThis.location.href).toString();

    // Prefer module workers (used by the engine + extension host). Fall back to classic workers
    // (used by the Pyodide worker backend).
    try {
      if (await runWorker(workerUrl, { type: "module" })) return true;
    } catch {}

    try {
      if (await runWorker(workerUrl)) return true;
    } catch {}

    return false;
  };

  let started = false;
  const tick = async () => {
    if (started) return;

    const emit = globalThis.__TAURI__?.event?.emit;
    if (typeof emit !== "function") {
      if (Date.now() > deadline) return;
      setTimeout(tick, 50);
      return;
    }

    started = true;
    const crossOriginIsolated = globalThis.crossOriginIsolated === true;
    const sharedArrayBuffer = typeof SharedArrayBuffer !== "undefined";
    let workerOk = false;
    try {
      workerOk = await checkWorker();
    } catch {}

    const payload = {
      cross_origin_isolated: crossOriginIsolated,
      shared_array_buffer: sharedArrayBuffer,
      worker_ok: workerOk,
    };

    // Emit can reject if the Tauri event bridge isn't fully ready yet; retry briefly.
    while (Date.now() <= deadline) {
      try {
        await emit("coi-check-result", payload);
        return;
      } catch {
        await sleep(50);
      }
    }
  };

  tick().catch(() => {});
})();
"#,
                    )
                    .unwrap_or_else(|err| {
                        eprintln!("[formula][coi-check] failed to eval script: {err}");
                        std::process::exit(2);
                    });

                // Skip the rest of normal app setup (tray icon, updater, open-file wiring, etc).
                // The check mode should be as lightweight as possible so it can run in headless
                // environments and exit quickly based on the WebView evaluation result.
                return Ok(());
            }

            tray::init(app)?;
            menu::init(app)?;

            // Register the `formula://` deep-link URL scheme with the OS so OAuth PKCE redirects
            // can round-trip back into the running desktop app instance.
            //
            // This is best-effort because some platforms may deny registration (e.g. sandboxed
            // environments) and we still want the app to launch normally.
            if let Err(err) = app.handle().deep_link().register("formula") {
                eprintln!("[deep-link] failed to register formula:// handler: {err}");
            }

            // Register global shortcuts (handled by the frontend via the Tauri plugin).
            shortcuts::register(app.handle())?;

            // Auto-update is configured via `tauri.conf.json`. We do a lightweight startup check
            // in release builds; users can also trigger checks from the tray menu.
            #[cfg(not(debug_assertions))]
            {
                if !should_disable_startup_update_check() {
                    // Tauri does not guarantee that emitted events are queued before JS listeners are
                    // registered. To avoid dropping fast startup update notifications, the frontend
                    // emits `updater-ui-ready` once it has installed its updater event listeners.
                    // Only then do we run the startup update check.
                    let handle = app.handle().clone();
                    let started = Arc::new(AtomicBool::new(false));
                    let listener = Arc::new(Mutex::new(None));

                    let started_for_listener = started.clone();
                    let listener_for_listener = listener.clone();
                    let handle_for_listener = handle.clone();

                    let id = handle.listen("updater-ui-ready", move |_| {
                        if started_for_listener.swap(true, Ordering::SeqCst) {
                            return;
                        }

                        if let Some(id) = listener_for_listener.lock().unwrap().take() {
                            handle_for_listener.unlisten(id);
                        }

                        updater::spawn_update_check(
                            &handle_for_listener,
                            updater::UpdateCheckSource::Startup,
                        );
                    });

                    *listener.lock().unwrap() = Some(id);

                    // Extremely defensive: if the readiness signal fires before we store `id`, make
                    // sure the listener is still unregistered.
                    if started.load(Ordering::SeqCst) {
                        if let Some(id) = listener.lock().unwrap().take() {
                            handle.unlisten(id);
                        }
                    }
                }
            }

            // Queue `open-file` requests until the frontend has installed its event listeners.
            if let Some(window) = app.get_webview_window("main") {
                let handle = app.handle().clone();
                window.listen(OPEN_FILE_READY_EVENT, move |_event| {
                    let state = handle.state::<SharedOpenFileState>().inner().clone();
                    let pending = {
                        let mut guard = state.lock().unwrap();
                        guard.mark_ready_and_drain()
                    };
                    let pending = normalize_open_file_request_paths(pending);

                    if !pending.is_empty() {
                        emit_open_file_event(&handle, pending);
                    }
                });
            }

            // Queue `oauth-redirect` requests until the frontend has installed its event listeners.
            if let Some(window) = app.get_webview_window("main") {
                let handle = app.handle().clone();
                window.listen(OAUTH_REDIRECT_READY_EVENT, move |_event| {
                    let state = handle.state::<SharedOauthRedirectState>().inner().clone();
                    let pending = {
                        let mut guard = state.lock().unwrap();
                        guard.mark_ready_and_drain()
                    };

                    let pending = normalize_oauth_redirect_request_urls(pending);
                    for url in pending {
                        emit_oauth_redirect_event(&handle, url);
                    }
                });
            }

            Ok(())
        })
        .build(tauri::generate_context!())
        .expect("error while building tauri application");

    app.run(|_app_handle, event| match event {
        // macOS/iOS: when the app is already running and the user opens a file via the OS,
        // the running instance receives an "open documents" event. Route it through the
        // same open-file pipeline used by argv / single-instance callbacks.
        #[cfg(any(target_os = "macos", target_os = "ios"))]
        tauri::RunEvent::Opened { urls, .. } => {
            let classified = opened_urls::classify_opened_urls(&urls);
            handle_oauth_redirect_request(_app_handle, classified.oauth_redirects);

            // File association / open-with handling.
            let paths = extract_open_file_paths(&classified.file_open_candidates, None);
            handle_open_file_request(_app_handle, paths);
        }
        _ => {}
    });
}
