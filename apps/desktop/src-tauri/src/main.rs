#![cfg_attr(
    all(not(debug_assertions), target_os = "windows"),
    windows_subsystem = "windows"
)]

mod shortcuts;
mod tray;
mod updater;

use formula_desktop_tauri::commands;
use formula_desktop_tauri::macro_trust::{MacroTrustStore, SharedMacroTrustStore};
use formula_desktop_tauri::state::{AppState, SharedAppState};
use std::sync::{Arc, Mutex};
use tauri::{Emitter, Manager};

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
             commands::power_query_credential_get,
             commands::power_query_credential_set,
             commands::power_query_credential_delete,
             commands::power_query_credential_list,
               commands::save_workbook,
               commands::mark_saved,
               commands::get_workbook_theme_palette,
               commands::list_defined_names,
              commands::list_tables,
             commands::get_cell,
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
                // Delegate close-handling to the frontend so it can:
                // - fire `Workbook_BeforeClose` macros
                // - prompt for unsaved changes
                // - decide whether to hide the window or keep it open
                api.prevent_close();
                let _ = window.emit("close-requested", ());
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
