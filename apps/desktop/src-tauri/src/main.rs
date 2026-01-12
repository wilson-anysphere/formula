#![cfg_attr(
    all(not(debug_assertions), target_os = "windows"),
    windows_subsystem = "windows"
)]

mod shortcuts;
mod tray;
mod updater;

use formula_desktop_tauri::commands;
use formula_desktop_tauri::macro_trust::{compute_macro_fingerprint, MacroTrustStore, SharedMacroTrustStore};
use formula_desktop_tauri::macros::MacroExecutionOptions;
use formula_desktop_tauri::state::{AppState, CellUpdateData, SharedAppState};
use serde::Serialize;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use tauri::{Emitter, Listener, Manager};
use tokio::sync::oneshot;
use tokio::time::{timeout, Duration};
use uuid::Uuid;

const WORKBOOK_ID: &str = "local-workbook";

static CLOSE_REQUEST_IN_FLIGHT: AtomicBool = AtomicBool::new(false);

#[derive(Clone, Debug, Serialize)]
struct CloseRequestedPayload {
    token: String,
    /// Optional cell updates produced by `Workbook_BeforeClose`.
    ///
    /// Note: if the user cancels the close in the frontend (e.g. via an unsaved-changes prompt),
    /// applying these updates keeps the frontend `DocumentController` consistent with the backend.
    updates: Vec<commands::CellUpdate>,
}

fn signature_status(vba_project_bin: &[u8]) -> commands::MacroSignatureStatus {
    let parsed = formula_vba::verify_vba_digital_signature(vba_project_bin)
        .ok()
        .flatten();

    match parsed {
        Some(sig) => match sig.verification {
            formula_vba::VbaSignatureVerification::SignedVerified => {
                commands::MacroSignatureStatus::SignedVerified
            }
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

fn main() {
    let state: SharedAppState = Arc::new(Mutex::new(AppState::new()));
    let macro_trust: SharedMacroTrustStore = Arc::new(Mutex::new(
        MacroTrustStore::load_default().unwrap_or_else(|_| {
            // Backend startup should not fail if the trust store is unreadable; fall back
            // to an ephemeral store (macros will remain blocked by default).
            MacroTrustStore::new_ephemeral()
        }),
    ));

    tauri::Builder::default()
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
        .manage(state)
        .manage(macro_trust)
        .invoke_handler(tauri::generate_handler![
            commands::open_workbook,
            commands::new_workbook,
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
            commands::fire_workbook_open,
            commands::fire_workbook_before_close,
            commands::fire_worksheet_change,
            commands::fire_selection_change,
        ])
        .setup(|app| {
            tray::init(app)?;

            // Register global shortcuts (handled by the frontend via the Tauri plugin).
            shortcuts::register(app.handle())?;

            // Auto-update is configured via `tauri.conf.json`. We do a lightweight startup check
            // in release builds; users can also trigger checks from the tray menu.
            #[cfg(not(debug_assertions))]
            updater::spawn_update_check(app.handle());

            Ok(())
        })
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
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
