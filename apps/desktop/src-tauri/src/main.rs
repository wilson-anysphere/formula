#![cfg_attr(
    all(not(debug_assertions), target_os = "windows"),
    windows_subsystem = "windows"
)]

mod shortcuts;
mod asset_protocol;
mod tray;
mod updater;

use desktop::commands;
use desktop::macro_trust::{compute_macro_fingerprint, MacroTrustStore, SharedMacroTrustStore};
use desktop::macros::MacroExecutionOptions;
use desktop::open_file;
use desktop::state::{AppState, CellUpdateData, SharedAppState};
use desktop::tray_status::{self, TrayStatusState};
use serde::Serialize;
use std::collections::HashSet;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use tauri::{Emitter, Listener, Manager};
use tauri_plugin_notification::NotificationExt;
use tokio::sync::oneshot;
use tokio::time::{timeout, Duration};
use uuid::Uuid;

const WORKBOOK_ID: &str = "local-workbook";

static CLOSE_REQUEST_IN_FLIGHT: AtomicBool = AtomicBool::new(false);

const OPEN_FILE_EVENT: &str = "open-file";
const OPEN_FILE_READY_EVENT: &str = "open-file-ready";

#[derive(Debug, Default)]
struct OpenFileState {
    ready: bool,
    pending_paths: Vec<String>,
}

type SharedOpenFileState = Arc<Mutex<OpenFileState>>;

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

fn handle_open_file_request(app: &tauri::AppHandle, paths: Vec<String>) {
    let paths = normalize_open_file_request_paths(paths);
    if paths.is_empty() {
        // Still focus the existing window on "warm start" launches with no file args.
        show_main_window(app);
        return;
    }

    show_main_window(app);

    let open_file_state = app.state::<SharedOpenFileState>().inner().clone();
    let mut state = open_file_state.lock().unwrap();

    if state.ready {
        drop(state);
        emit_open_file_event(app, paths);
    } else {
        state.pending_paths.extend(paths);
    }
}

fn extract_open_file_paths(argv: &[String], cwd: Option<&Path>) -> Vec<String> {
    open_file::extract_open_file_paths_from_argv(argv, cwd)
        .into_iter()
        .map(|path| path.to_string_lossy().to_string())
        .collect()
}

fn cwd_from_single_instance_callback(cwd: String) -> Option<PathBuf> {
    let trimmed = cwd.trim();
    if trimmed.is_empty() {
        return None;
    }
    Some(PathBuf::from(trimmed))
}

fn signature_status(vba_project_bin: &[u8]) -> commands::MacroSignatureStatus {
    let parsed = formula_vba::verify_vba_digital_signature(vba_project_bin)
        .ok()
        .flatten();

    match parsed {
        Some(sig) => match sig.verification {
            formula_vba::VbaSignatureVerification::SignedVerified => match sig.binding {
                formula_vba::VbaSignatureBinding::Bound => commands::MacroSignatureStatus::SignedVerified,
                formula_vba::VbaSignatureBinding::NotBound => commands::MacroSignatureStatus::SignedInvalid,
                formula_vba::VbaSignatureBinding::Unknown => commands::MacroSignatureStatus::SignedUnverified,
            },
            formula_vba::VbaSignatureVerification::SignedInvalid => commands::MacroSignatureStatus::SignedInvalid,
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
    let sig_status = signature_status(vba_bin);
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
    // Restrict notification triggers to the main application window. This avoids
    // accidental abuse if we ever embed untrusted content in secondary webviews.
    if window.label() != "main" {
        return Err("notifications are only allowed from the main window".to_string());
    }

    let mut builder = window
        .app_handle()
        .notification()
        .builder()
        .title(title);

    if let Some(body) = body {
        builder = builder.body(body);
    }

    builder.show().map_err(|err| err.to_string())?;
    Ok(())
}

fn main() {
    let state: SharedAppState = Arc::new(Mutex::new(AppState::new()));
    let macro_trust: SharedMacroTrustStore = Arc::new(Mutex::new(
        MacroTrustStore::load_default().unwrap_or_else(|_| {
            // Backend startup should not fail if the trust store is unreadable; fall back
            // to an ephemeral store (macros will remain blocked by default).
            MacroTrustStore::new_ephemeral()
        }),
    ));

    let open_file_state: SharedOpenFileState = Arc::new(Mutex::new(OpenFileState::default()));
    let initial_argv: Vec<String> = std::env::args().collect();
    let initial_cwd = std::env::current_dir().ok();
    let initial_paths = normalize_open_file_request_paths(extract_open_file_paths(
        &initial_argv,
        initial_cwd.as_deref(),
    ));
    if !initial_paths.is_empty() {
        let mut guard = open_file_state.lock().unwrap();
        guard.pending_paths.extend(initial_paths);
    }

    let app = tauri::Builder::default()
        // Override Tauri's default `asset:` protocol handler to attach COEP-friendly headers.
        // See `asset_protocol.rs` for details.
        .register_uri_scheme_protocol("asset", asset_protocol::handler)
        .plugin(tauri_plugin_single_instance::init(|app, argv, cwd| {
            // OAuth PKCE deep-link redirect capture (e.g. `formula://oauth/callback?...`).
            //
            // When an OAuth provider redirects to our custom URI scheme, the OS may attempt to
            // launch a second instance of the application. The single-instance plugin forwards
            // the argv to the running instance; emit the URL to the frontend so it can resolve
            // any pending `DesktopOAuthBroker.waitForRedirect(...)` promises.
            for arg in &argv {
                let url = arg.trim().trim_matches('"');
                if url.starts_with("formula://") {
                    let _ = app.emit("oauth-redirect", url.to_string());
                }
            }

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
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_notification::init())
        .manage(state)
        .manage(macro_trust)
        .manage(open_file_state)
        .manage(TrayStatusState::default())
        .invoke_handler(tauri::generate_handler![
            commands::open_workbook,
            commands::new_workbook,
            commands::add_sheet,
            commands::read_text_file,
            commands::read_binary_file,
            commands::read_binary_file_range,
            commands::stat_file,
            commands::list_dir,
            commands::power_query_cache_key_get_or_create,
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
            commands::get_cell,
            commands::get_precedents,
            commands::get_dependents,
            commands::set_cell,
            commands::get_range,
            commands::get_sheet_used_range,
            commands::set_range,
            commands::create_pivot_table,
            commands::refresh_pivot_table,
            commands::list_pivot_tables,
            commands::recalculate,
            commands::undo,
            commands::redo,
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
            commands::quit_app,
            commands::exit_process,
            commands::report_cross_origin_isolation,
            commands::fire_workbook_open,
            commands::fire_workbook_before_close,
            commands::fire_worksheet_change,
            commands::fire_selection_change,
            tray_status::set_tray_status,
            show_system_notification,
        ])
        .on_window_event(|window, event| match event {
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
                        let Some(payload) = event.payload() else {
                            return;
                        };
                        let received = payload.trim().trim_matches('"');
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
                        let trust_store = trust_for_macro.lock().unwrap();

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
                        let Some(payload) = event.payload() else {
                            return;
                        };
                        let received = payload.trim().trim_matches('"');
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
        .setup(|app| {
            if std::env::args().any(|arg| arg == "--cross-origin-isolation-check") {
                // CI/developer smoke test: validate cross-origin isolation (COOP/COEP) in the
                // packaged Tauri build by running in a special mode that exits quickly with a
                // status code.
                //
                // This is evaluated inside the WebView so we can check `globalThis.crossOriginIsolated`
                // and `SharedArrayBuffer` availability.
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

                window
                    .eval(
                        r#"
(() => {
  const deadline = Date.now() + 10_000;
  const tick = () => {
    const invoke = globalThis.__TAURI__?.core?.invoke;
    if (typeof invoke === "function") {
      const crossOriginIsolated = globalThis.crossOriginIsolated === true;
      const sharedArrayBuffer = typeof SharedArrayBuffer !== "undefined";
      const ok = crossOriginIsolated && sharedArrayBuffer;

      invoke("report_cross_origin_isolation", {
        cross_origin_isolated: crossOriginIsolated,
        shared_array_buffer: sharedArrayBuffer,
      }).catch(() => {});

      invoke("exit_process", { code: ok ? 0 : 1 });
      return;
    }
    if (Date.now() > deadline) return;
    setTimeout(tick, 50);
  };
  tick();
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

            // Register global shortcuts (handled by the frontend via the Tauri plugin).
            shortcuts::register(app.handle())?;

            // Auto-update is configured via `tauri.conf.json`. We do a lightweight startup check
            // in release builds; users can also trigger checks from the tray menu.
            #[cfg(not(debug_assertions))]
            updater::spawn_update_check(app.handle(), updater::UpdateCheckSource::Startup);

            // Best-effort: if the app was launched via a deep-link URL (e.g. the first
            // instance after an OAuth redirect), forward it to the frontend.
            //
            // When the app is already running, `tauri_plugin_single_instance` forwards URLs
            // to this process via its callback above.
            for arg in std::env::args() {
                let url = arg.trim().trim_matches('"');
                if url.starts_with("formula://") {
                    let _ = app.emit("oauth-redirect", url.to_string());
                }
            }

            // Queue `open-file` requests until the frontend has installed its event listeners.
            if let Some(window) = app.get_webview_window("main") {
                let handle = app.handle().clone();
                window.listen(OPEN_FILE_READY_EVENT, move |_event| {
                    let state = handle.state::<SharedOpenFileState>().inner().clone();
                    let pending = {
                        let mut guard = state.lock().unwrap();
                        if guard.ready {
                            return;
                        }
                        guard.ready = true;
                        std::mem::take(&mut guard.pending_paths)
                    };
                    let pending = normalize_open_file_request_paths(pending);

                    if !pending.is_empty() {
                        emit_open_file_event(&handle, pending);
                    }
                });
            }

            Ok(())
        })
        .build(tauri::generate_context!())
        .expect("error while building tauri application");

    app.run(|app_handle, event| match event {
        // macOS: when the app is already running and the user opens a file via Finder,
        // the running instance receives an "open documents" event. Route it through the
        // same open-file pipeline.
        tauri::RunEvent::Opened { urls, .. } => {
            let argv: Vec<String> = urls.iter().map(|url| url.to_string()).collect();
            let paths = extract_open_file_paths(&argv, None);
            handle_open_file_request(app_handle, paths);
        }
        _ => {}
    });
}
